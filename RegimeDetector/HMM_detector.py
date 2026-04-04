import numpy as np
import ssm
from RegimeDetector.base_detector import BaseRegimeDetector


class HMMDetector(BaseRegimeDetector):
    """
    HMM Gaussien (Classic Hidden Markov Model) basé sur la librairie ssm (Linderman Lab).

    Le modèle est fitté sur des variables d'observation (ex: macro standardisées).
    Les métriques de régime (moyennes, covariances) sont calculées sur les log returns financiers,
    en mode statique (empirique par état Viterbi) ou dynamique (EWMA pondéré par les probabilités).

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
        self.random_state = random_state
        self.seed = random_state            # alias conservé pour compatibilité interne
        self.kappa = kappa
        self.n_restarts = n_restarts

        # Attributs d'interface attendus par TransitionStatisticalAnalyzer / CoherenceChecker
        self.K = n_states                   # nombre de régimes (alias de n_states)
        self.D = None                       # dimension des observations (fixée au fit)
        self.M = 0                          # nombre d'inputs covariables (0 = HMM non-conditionnel)

        self.model = None
        self.is_fitted = False              # booléen simple, mis à True dans fit()
        self._fitted_probs = None           # Cache des probabilités lissées (in-sample)
        self.viterbi_states = None          # Séquence d'états la plus probable (Viterbi)
        self._observations_fitted = None    # Observations d'entraînement (ex: macro)
        self._log_returns_fitted = None     # Log returns financiers alignés sur le fit
        self.fit_observations = None        # Alias liste-wrappé attendu par TransitionStatisticalAnalyzer
        self.fit_input = None               # Inputs covariables (None pour HMM non-conditionnel)

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
    # Méthode privée : résolution des log returns
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
            return np.array(log_returns, dtype=float)
        if self._log_returns_fitted is not None:
            return self._log_returns_fitted
        raise ValueError(
            "Aucun log_returns disponible. "
            "Passez log_returns au fit() ou directement à cette méthode."
        )

    # ------------------------------------------------------------------
    # Méthode privée : calcul des poids EWMA
    # ------------------------------------------------------------------

    def _ewma_weights(self, t: int, regime_probs: np.ndarray, halflife: int) -> np.ndarray:
        """
        Calcule les poids EWMA combinés pour le régime k jusqu'à l'instant t.

        Poids = décroissance exponentielle × probabilité d'être dans le régime k.

        Paramètres
        ----------
        t : int
            Instant courant (exclusif).
        regime_probs : np.ndarray (t, K)
            Probabilités de régime jusqu'à t.
        halflife : int
            Demi-vie en mois pour la décroissance EWMA.

        Retourne
        --------
        np.ndarray (t, K) : poids normalisés pour chaque régime.
        """
        alpha = 1 - np.exp(-np.log(2) / halflife)
        decay = np.array([(1 - alpha) ** i for i in range(t)])[::-1]  # (t,)

        # Poids combinés : décroissance × probabilité de régime
        w = decay[:, None] * regime_probs[:t]  # (t, K)

        # Normalisation par régime
        w_sum = w.sum(axis=0, keepdims=True)  # (1, K)
        w_sum = np.where(w_sum < 1e-8, 1.0, w_sum)
        return w / w_sum  # (t, K)

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
            Accepte une liste de la forme [obs_array] ou un tableau (T, D).
            Si une liste est fournie, observations[0] est utilisé pour le fit
            et observations[1] (si présent) est stocké comme log_returns financiers.
        inputs : ignoré
            HMMDetector est un modèle non-conditionnel sur les transitions.
        log_returns : List[np.ndarray] ou np.ndarray, optionnel
            Log returns financiers alignés sur observations.
            Stockés pour le calcul des métriques de régime (moyennes, covariances).
            Prioritaire sur observations[1] si les deux sont fournis.
        """
        # Extraction si format liste :
        # observations[0] → données du fit HMM
        # observations[1] → log returns financiers (optionnel, si log_returns non fourni)
        if isinstance(observations, list):
            if log_returns is None and len(observations) > 1:
                log_returns = observations[1]
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
        self.fit_observations = [observations]  # format liste attendu par TransitionStatisticalAnalyzer
        self.fit_input = None                    # HMM non-conditionnel : pas d'inputs
        self.D = D          # dimension des observations (disponible post-fit)
        self.K = self.n_states  # redondant mais garantit la cohérence si n_states change
        self.is_fitted = True
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
    # regime_covariances — statique ou dynamique EWMA
    # ------------------------------------------------------------------

    def regime_covariances(self, log_returns: np.ndarray = None,
                           use_ewma: bool = True,
                           halflife: int = 24):
        """
        Retourne les matrices de covariance des log returns par régime.

        Mode statique (use_ewma=False) :
            Covariances empiriques fixes par état Viterbi — shape (K, D, D).
            Cas statique au sens de base.py : paramètres fixes estimés lors du fit.

        Mode dynamique EWMA (use_ewma=True) :
            Covariances EWMA pondérées par les probabilités de régime — shape (T, K, D, D).
            Cas dynamique au sens de base.py : séquence temporelle des matrices.
            Chaque matrice C[t, k] est estimée sur l'historique jusqu'à t,
            pondéré exponentiellement et par P(z_t = k).

        Paramètres
        ----------
        log_returns : np.ndarray, optionnel (T, D)
            Log returns financiers. Si None, utilise ceux stockés lors du fit.
        use_ewma : bool
            True → mode dynamique EWMA, False → mode statique empirique.
        halflife : int
            Demi-vie en mois pour la décroissance EWMA (utilisé si use_ewma=True).
        """
        if self.model is None:
            raise ValueError("Le modèle n'a pas été entraîné.")

        lr = self._resolve_log_returns(log_returns)
        T, D = lr.shape
        K = self.n_states

        if not use_ewma:
            # Mode statique : covariance empirique par état Viterbi
            covs = []
            for k in range(K):
                mask = (self.viterbi_states == k)
                subset = lr[mask]
                covs.append(np.cov(subset.T, ddof=1) if subset.shape[0] >= 2 else np.eye(D))
            return covs

        # Mode dynamique EWMA — shape (T, K, D, D)
        covs_dynamic = np.zeros((T, K, D, D))

        for t in range(1, T):
            w = self._ewma_weights(t, self._fitted_probs, halflife)  # (t, K)

            for k in range(K):
                w_k = w[:, k]  # (t,)

                # Moyenne pondérée
                mu_k = (lr[:t] * w_k[:, None]).sum(axis=0)

                # Covariance pondérée
                diff = lr[:t] - mu_k
                covs_dynamic[t, k] = (diff * w_k[:, None]).T @ diff

        return covs_dynamic

    # ------------------------------------------------------------------
    # regime_means — statique ou dynamique EWMA
    # ------------------------------------------------------------------

    def regime_means(self, log_returns: np.ndarray = None,
                     use_ewma: bool = True,
                     halflife: int = 24):
        """
        Retourne les moyennes des log returns par régime.

        Mode statique (use_ewma=False) :
            Moyennes empiriques fixes par état Viterbi — shape (K, D).

        Mode dynamique EWMA (use_ewma=True) :
            Moyennes EWMA pondérées par les probabilités de régime — shape (T, K, D).

        Paramètres
        ----------
        log_returns : np.ndarray, optionnel (T, D)
            Log returns financiers. Si None, utilise ceux stockés lors du fit.
        use_ewma : bool
            True → mode dynamique EWMA, False → mode statique empirique.
        halflife : int
            Demi-vie en mois pour la décroissance EWMA (utilisé si use_ewma=True).
        """
        if self.model is None:
            raise ValueError("Le modèle n'a pas été entraîné.")

        lr = self._resolve_log_returns(log_returns)
        T, D = lr.shape
        K = self.n_states

        if not use_ewma:
            # Mode statique : moyenne empirique par état Viterbi
            means = []
            for k in range(K):
                mask = (self.viterbi_states == k)
                subset = lr[mask]
                means.append(np.mean(subset, axis=0) if subset.shape[0] > 0 else np.zeros(D))
            return means

        # Mode dynamique EWMA — shape (T, K, D)
        means_dynamic = np.zeros((T, K, D))

        for t in range(1, T):
            w = self._ewma_weights(t, self._fitted_probs, halflife)  # (t, K)

            for k in range(K):
                w_k = w[:, k]  # (t,)
                means_dynamic[t, k] = (lr[:t] * w_k[:, None]).sum(axis=0)

        return means_dynamic

    # ------------------------------------------------------------------
    # conditional_covariance — délégation à BaseRegimeDetector
    # ------------------------------------------------------------------

    def conditional_covariance(self, probs: np.ndarray,
                                log_returns: np.ndarray = None) -> np.ndarray:
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
    # get_transition_params / set_transition_params
    # Requis par TransitionStatisticalAnalyzer (save → perturb → restore)
    # ------------------------------------------------------------------

    def get_transition_params(self):
        """
        Retourne les paramètres bruts de transition pour sauvegarde.

        Pour un HMM standard (non-conditionnel), seuls les log-probabilités
        de transition log_Ps sont des paramètres libres. Ws est retourné comme
        un array vide de shape (K, 0) — sans covariables — pour rester compatible
        avec TransitionStatisticalAnalyzer qui suppose toujours un array indexable.

        Retourne
        --------
        log_Ps : np.ndarray (K, K) — copie des log-probabilités de transition.
        Ws     : np.ndarray (K, 0) — array vide (pas de covariables).
        """
        if self.model is None:
            raise ValueError("Le modèle n'a pas été entraîné.")
        log_Ps = self.model.transitions.log_Ps.copy()
        Ws = np.zeros((self.K, 0))   # (K, M=0) : compatible avec toutes les ops array
        return log_Ps, Ws

    def set_transition_params(self, log_Ps: np.ndarray, Ws=None):
        """
        Restaure les paramètres de transition (après perturbation numérique).

        Paramètres
        ----------
        log_Ps : np.ndarray (K, K) — log-probabilités de transition à restaurer.
        Ws     : ignoré (pas de covariables dans un HMM standard).
        """
        if self.model is None:
            raise ValueError("Le modèle n'a pas été entraîné.")
        self.model.transitions.log_Ps = log_Ps.copy()

    # ------------------------------------------------------------------
    # compute_confidence_index — requise par HMMCoherenceChecker
    # ------------------------------------------------------------------

    def compute_confidence_index(self, series_index: int = 0) -> np.ndarray:
        """
        Calcule l'indice de confiance à chaque pas de temps comme (1 - entropie normalisée).

        Une valeur proche de 1 indique que le modèle est très certain du régime courant
        (une probabilité domine). Une valeur proche de 0 indique une forte ambiguïté.

        Paramètres
        ----------
        series_index : int
            Ignoré — conservé pour compatibilité avec le contrat BaseRegimeDetector.

        Retourne
        --------
        np.ndarray (T,) : indice de confiance ∈ [0, 1] pour chaque date.
        """
        if self._fitted_probs is None:
            raise ValueError("Le modèle n'a pas été entraîné.")

        probs = self._fitted_probs  # (T, K)
        K = probs.shape[1]

        # Entropie de Shannon normalisée : H ∈ [0, 1]
        # On clip pour éviter log(0)
        log_probs = np.log(np.clip(probs, 1e-12, 1.0))
        entropy = -np.sum(probs * log_probs, axis=1)          # (T,)
        max_entropy = np.log(K)                                # entropie max = log(K) (uniforme)
        normalized_entropy = entropy / max_entropy             # (T,) ∈ [0, 1]

        return 1.0 - normalized_entropy                        # (T,) ∈ [0, 1]
