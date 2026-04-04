import numpy as np
import ssm
from RegimeDetector.base_detector import BaseRegimeDetector


class HMMDetector(BaseRegimeDetector):
    """
    HMM Gaussien (Classic Hidden Markov Model) basé sur la librairie ssm (Linderman Lab).

    Le modèle est fitté sur des variables d'observation (ex: macro standardisées).
    Les métriques de régime (moyennes, covariances) sont calculées sur les log returns financiers.

    Paramètres
    ----------
    n_states : int
        Nombre de régimes latents.
    n_iter : int
        Nombre d'itérations EM pour l'apprentissage.
    random_state : int
        Graine aléatoire pour la reproductibilité du modèle final.
    kappa : float
        Force de persistance des régimes (transitions "sticky").
        0 = transitions standard sans biais, >0 = biais de persistance croissant.
    n_restarts : int
        Nombre de restarts avec initialisation K-Means.
        Le modèle retenu est celui maximisant la log-vraisemblance.
    """

    def __init__(self, n_states: int = 2, n_iter: int = 100, random_state: int = 42,
                 kappa: float = 5.0, n_restarts: int = 10):
        self.n_states = n_states
        self.n_iter = n_iter
        self.seed = random_state
        self.kappa = kappa
        self.n_restarts = n_restarts
        self.model = None
        self._fitted_probs = None       # Cache des probabilités lissées (in-sample)
        self.viterbi_states = None      # Séquence d'états la plus probable (Viterbi)
        self._observations_fitted = None  # Observations d'entraînement (ex: macro)
        self._log_returns_fitted = None   # Log returns financiers alignés sur le fit

    # ------------------------------------------------------------------
    # Méthode privée : instanciation du modèle SSM
    # ------------------------------------------------------------------

    def _build_model(self, D: int):
        """
        Instancie un HMM ssm selon la configuration de l'objet.

        Paramètres
        ----------
        D : int
            Dimension des observations.
        """
        if self.kappa > 0:
            return ssm.HMM(
                self.n_states, D,
                observations="gaussian",
                transitions="sticky",
                transitions_kwargs=dict(alpha=1.0, kappa=self.kappa)
            )
        return ssm.HMM(self.n_states, D, observations="gaussian")

    # ------------------------------------------------------------------
    # fit
    # ------------------------------------------------------------------

    def fit(self, observations, inputs=None, log_returns=None):
        """
        Entraîne le HMM sur les observations via l'algorithme EM avec initialisation K-Means.

        Effectue n_restarts runs indépendants et retient le modèle maximisant
        la log-vraisemblance pour éviter les optima locaux.

        Paramètres
        ----------
        observations : List[np.ndarray] ou np.ndarray
            Variables d'observation du HMM (ex : macro standardisées).
            Accepte une liste (contrat RegimeDetector) ou un tableau (T, D).
        inputs : ignoré
            HMMDetector est un modèle non-conditionnel sur les transitions.
        log_returns : List[np.ndarray] ou np.ndarray, optionnel
            Log returns financiers alignés sur observations.
            Stockés pour le calcul des métriques de régime (moyennes, covariances).
        """
        # Extraction si format liste
        if isinstance(observations, list):
            observations = observations[0]
        observations = np.array(observations, dtype=float)
        if observations.ndim == 1:
            observations = observations[:, None]

        # Stockage des log returns financiers si fournis
        if log_returns is not None:
            lr = log_returns[0] if isinstance(log_returns, list) else log_returns
            self._log_returns_fitted = np.array(lr, dtype=float)
        else:
            self._log_returns_fitted = None

        T, D = observations.shape

        # Multi-restart K-Means : on retient le modèle avec la meilleure log-vraisemblance
        best_ll = -np.inf

        for seed in range(self.n_restarts):
            np.random.seed(seed)
            candidate = self._build_model(D)
            candidate.fit(observations, method="em", num_iters=self.n_iter,
                          init_method="kmeans", verbose=0)
            ll = candidate.log_probability(observations)
            if ll > best_ll:
                best_ll = ll
                self.model = candidate

        # Probabilités lissées P(z_t | Y_{1:T}) via Forward-Backward
        self._fitted_probs = self.model.expected_states(observations)[0]

        # Séquence optimale d'états via l'algorithme de Viterbi
        self.viterbi_states = self.model.most_likely_states(observations)

        self._observations_fitted = observations
        return self

    # ------------------------------------------------------------------
    # regime_probabilities
    # ------------------------------------------------------------------

    def regime_probabilities(self, series_index: int = 0) -> np.ndarray:
        """
        Retourne les probabilités lissées P(z_t | Y_{1:T}) in-sample — shape (T, K).

        Paramètres
        ----------
        series_index : ignoré (conservé pour compatibilité avec le contrat abstrait).
        """
        if self.model is None:
            raise ValueError("Le modèle n'a pas été entraîné.")
        return self._fitted_probs

    # ------------------------------------------------------------------
    # regime_covariances — sur les log returns, pas les observations macro
    # ------------------------------------------------------------------

    def regime_covariances(self, log_returns: np.ndarray = None):
        """
        Retourne les matrices de covariance empiriques des log returns par régime — shape (K, D, D).

        Les covariances sont calculées sur les log returns financiers conditionnellement
        aux états Viterbi, et non sur les observations macro du HMM.

        Paramètres
        ----------
        log_returns : np.ndarray, optionnel
            Log returns financiers (T, D). Si None, utilise ceux stockés lors du fit.
        """
        if self.model is None:
            raise ValueError("Le modèle n'a pas été entraîné.")

        lr = self._resolve_log_returns(log_returns)

        K = self.n_states
        D = lr.shape[1]
        covs = []

        for k in range(K):
            mask = (self.viterbi_states == k)
            subset = lr[mask]
            if subset.shape[0] < 2:
                covs.append(np.eye(D))
            else:
                covs.append(np.cov(subset.T, ddof=1))

        return covs

    # ------------------------------------------------------------------
    # regime_means — sur les log returns, pas les observations macro
    # ------------------------------------------------------------------

    def regime_means(self, log_returns: np.ndarray = None):
        """
        Retourne les moyennes empiriques des log returns par régime — shape (K, D).

        Les moyennes sont calculées sur les log returns financiers conditionnellement
        aux états Viterbi, et non sur les observations macro du HMM.

        Paramètres
        ----------
        log_returns : np.ndarray, optionnel
            Log returns financiers (T, D). Si None, utilise ceux stockés lors du fit.
        """
        if self.model is None:
            raise ValueError("Le modèle n'a pas été entraîné.")

        lr = self._resolve_log_returns(log_returns)

        K = self.n_states
        means = []

        for k in range(K):
            mask = (self.viterbi_states == k)
            subset = lr[mask]
            if subset.shape[0] == 0:
                means.append(np.zeros(lr.shape[1]))
            else:
                means.append(np.mean(subset, axis=0))

        return means

    # ------------------------------------------------------------------
    # conditional_covariance — délégation à BaseRegimeDetector
    # ------------------------------------------------------------------

    def conditional_covariance(self, probs: np.ndarray, log_returns: np.ndarray = None) -> np.ndarray:
        """
        Calcule la covariance conditionnelle totale via la loi des covariances totales.

        Délègue à BaseRegimeDetector.conditional_covariance en lui fournissant
        les métriques de régime calculées sur les log returns.

        Paramètres
        ----------
        probs : np.ndarray (T, K)
            Probabilités de régime à chaque date.
        log_returns : np.ndarray, optionnel
            Log returns financiers (T, D). Si None, utilise ceux stockés lors du fit.
        """
        lr = self._resolve_log_returns(log_returns)
        return super().conditional_covariance(probs=probs, log_returns=lr)

    # ------------------------------------------------------------------
    # predict_probabilities
    # ------------------------------------------------------------------

    def predict_probabilities(self, observations: np.ndarray, inputs: np.ndarray = None,
                              oos_start: int = 0, oos_end: int = None) -> np.ndarray:
        """
        Calcule les probabilités prédictives causales P(z_t | Y_{1:t-1}) sur l'OOS.

        Utilise le filtrage Forward (causal) — pas de look-ahead bias.

        Paramètres
        ----------
        observations : np.ndarray
            Observations complètes in-sample + OOS (T_total, D).
        inputs : ignoré
        oos_start : int
            Indice de début de l'OOS dans observations.
        oos_end : int, optionnel
            Indice de fin de l'OOS. Si None, va jusqu'à la fin.

        Retourne
        --------
        np.ndarray (T_oos, K) : probabilités prédictives sur l'OOS uniquement.
        """
        if self.model is None:
            raise ValueError("Le modèle n'a pas été entraîné.")

        observations = np.array(observations, dtype=float)
        if observations.ndim == 1:
            observations = observations[:, None]

        if oos_end is None:
            oos_end = len(observations)

        # Filtrage causal sur toute la séquence (in-sample + OOS pour contexte)
        _, filtered = self.model.filter(observations)

        # On ne retourne que les probabilités filtrées sur l'OOS
        return filtered[oos_start:oos_end]

    # ------------------------------------------------------------------
    # get_transition_matrix
    # ------------------------------------------------------------------

    def get_transition_matrix(self) -> np.ndarray:
        """
        Retourne la matrice de transition stochastique A (K, K),
        où A[i, j] = P(z_{t+1} = j | z_t = i).
        """
        if self.model is None:
            raise ValueError("Le modèle n'a pas été entraîné.")
        return np.exp(self.model.transitions.log_Ps)

    # ------------------------------------------------------------------
    # Méthode privée utilitaire
    # ------------------------------------------------------------------

    def _resolve_log_returns(self, log_returns: np.ndarray = None) -> np.ndarray:
        """
        Résout les log returns à utiliser : argument explicite ou cache du fit.

        Paramètres
        ----------
        log_returns : np.ndarray ou None

        Retourne
        --------
        np.ndarray (T, D)
        """
        if log_returns is not None:
            lr = np.array(log_returns, dtype=float)
            return lr
        if self._log_returns_fitted is not None:
            return self._log_returns_fitted
        raise ValueError(
            "Aucun log_returns disponible. "
            "Passez log_returns au fit() ou directement à cette méthode."
        )
