import numpy as np
import pandas as pd
from sklearn.linear_model import LogisticRegression
from RegimeDetector.HMMMarketRegimeDetector import HMMMarketRegimeDetector
from typing import List
import ssm
from RegimeDetector.EmissionStatisticalAnalyzer import EmissionStatisticalAnalyzer


class HMMMarketWithMacroTransitionsRegimeDetector(HMMMarketRegimeDetector):
    """
    Wrappeur de la librairie ssm pour détecter les régimes de marché dont les transitions sont impactées par des covariables macroéconomiques
    Hérite de HMMRegimeDetector pour conserver les méthodes communes

    Fonctionnement générique : fit un HMM sur des données jusqu'à T (Y_{1:t}, X_{1:t}) puis prédit les régimes en T+1

    Note importante : la librairie ssm utilise X_{t+1} pour passer de l'état en t à l'état en t+1, donc l'utilisateur doit lagguer
    ces variables de 1
    """
    def __init__(self, n_regimes : int = 2, n_dim : int = 2, n_input : int = 0, **kwargs : dict):
        """
        Paramètres:
        - n_regimes : nombre de régimes (défaut : 2)
        - n_dim : nombre de dimensions (défaut : 2)
        - n_input : nombre d'inputs (défaut : 0)
        - **kwargs : dictionnaire dynamique qui permet de rajouter des paramètres spécifiques au IOHMM (ces paramètres seront
        ensuite passés à la classe mère sous forme key=value)
        """
        kwargs['transitions'] = "inputdriven" # pas besoin que l'utilisateur précise
        super().__init__(n_regimes=n_regimes, n_dim=n_dim, n_input=n_input, **kwargs)

    def _display_logistic_weights(self, clf : LogisticRegression, feature_names : List[str]):
        """
        Affiche les poids de la régression logistique servant à l'initialisation.

        Paramètres :
        - clf : la régression logistique sur laquelle on veut afficher les poids
        - feature_names : les noms des features utilisées dans la régression logistique
        """
        print("\n=== POIDS INITIAUX DE LA RÉGRESSION LOGISTIQUE (X_{t-1} -> Z_t) ===")
        weights_df = pd.DataFrame(clf.coef_, columns=feature_names) # clf.coef_ est de taille (K, M)
        weights_df.index = [f"Probabilité Régime {i}" for i in range(self.K)]
        weights_df.insert(0, 'Intercept', clf.intercept_) # On ajoute l'intercept (le biais de base de chaque régime)

        print(weights_df.round(4))
        print("Note: Un poids positif augmente la probabilité de transition vers ce régime.")
        print("-" * 62)

    def fit(self, observations : List[np.ndarray], inputs : List[np.ndarray], log_returns : List[np.ndarray] = None, asset_names : List[str] = None, feature_names : List[str] = None, series_index : int = 0,
            num_iters : int = 200, tolerance : float = 1e-4, method : str = "em", warm_start : bool = True, initialize : bool = True, init_method: str = "kmeans"):
        """
        Fit le modèle.
        Lorsque warm_start=True, pour initialiser les paramètres d'émission, on lance un HMM sur les séries de
        log returns. Ensuite, pour initialiser les paramètres de transition, on lance un régression logistique multinomiale
        telle que l'output est le régime prédit à la date t et les features sont les covariables
        à la date t-1 (suppose un laggue d'ordre 1).
        Sinon, les paramètres d'émissions sont initialisés par K-means et les paramètres de transition de manière aléatoire.

        Paramètres :
        - observations : listes de tableaux de log returns
        - inputs : listes de tableaux de covariables
        - log_returns : ignoré dans ce modèle car déjà dans observations (mais présent par soucis de généralité)
        - feature_names : noms des covariables
        - series_index : index du tableau sur lequel on veut stocker les probabilités (défaut : 0)
        - num_iters : nombre d'itérations (défaut : 200)
        - tolerance : tolérance d'arrêt de l'algo d'optimisation
        - method : méthode d'optimisation (défaut : em)
        - warm_start : True si on veut lancer une init via HMM et Régression Logistique
        """
        if feature_names is None: # noms par défaut
            feature_names = [f"Input_{i}" for i in range(self.M)]

        if asset_names is None: # noms par défaut
            asset_names = [f"Asset_{i}" for i in range(self.D)]


        if warm_start: # on init avec un HMM pour la partie émission et avec une regression logistique entre la macro et les états du HMM pour la partie transition
            # 1. Fit du HMM stationnaire pour l'initialisation
            transition_kwargs_hmm = {key : value for (key, value) in self.transition_kwargs.items() if key != "l2_penalty"}
            np.random.seed(self.random_state)
            simple_hmm = ssm.HMM(self.K, self.D, M=0,
                                 observations=self.observations_type,
                                 transitions="sticky",
                                 transition_kwargs=transition_kwargs_hmm)

            print("Initialisation : Fit du HMM stationnaire...")
            np.random.seed(self.random_state)
            simple_hmm.fit(observations, method=method, num_iters=num_iters,
                           init_method=init_method, verbose=0, tolerance=tolerance, initialize=initialize)

            # 2. Affichage des stats de validation (pour vérifier l'initialisation à vue d'oeil)
            initial_states = simple_hmm.most_likely_states(observations[series_index])

            print("\n" + "="*70)
            print("AFFICHAGE DE L'INITIALISATION AVEC UN HMM STATIONNAIRE (EMISSION)")
            print("="*70)

            init_analyzer = EmissionStatisticalAnalyzer(
                log_returns=observations[series_index],
                states=initial_states,
                asset_names=asset_names
            )
            df_init = init_analyzer.get_descriptive_stats_df()
            print(df_init.to_string(index=False)) # pour affichage lisible

            # 3. Régression logistique pour les poids W
            print("\n" + "="*70)
            print("AFFICHAGE DE L'INITIALISATION AVEC UNE REGRESSION LOGISTIQUE (TRANSITION)")
            print("="*70)

            labels_list = [simple_hmm.most_likely_states(y) for y in observations]
            z_stacked = np.concatenate(labels_list) # transforme en une seule liste qui représente la suite des états à chaque date d'in-sample
            X_stacked = np.vstack(inputs) # transforme la liste de matrices en une seule grosse matrice

            clf = LogisticRegression(multi_class='multinomial', penalty='l2', C=1.0) # TO DO: justifier les hyperpamètres de cette Log Reg?
            clf.fit(X_stacked, z_stacked)

            # 4. Affichage des poids de l'initialisation (pour vérifier l'initialisation à vue d'oeil)
            self._display_logistic_weights(clf, feature_names)

            # 5. Initialisation
            self.model.transitions.Ws = clf.coef_ # ssm attend Ws de forme (K, M) pour InputDrivenTransitions
            self.model.transitions.log_Ps = np.tile(clf.intercept_, (self.K, 1)) # init aussi l'intercept (transforme de K à (1, K))
            self.model.observations.params = simple_hmm.observations.params

            Y_to_fit = observations
            X_to_fit = inputs
            final_initialize = False
        else:
            final_initialize = initialize
            Y_to_fit = observations
            X_to_fit = inputs

        # 7. Fit final du modèle Macro
        np.random.seed(self.random_state)
        self.model.fit(Y_to_fit, inputs=X_to_fit, method=method, num_iters=num_iters,
                       init_method=init_method, verbose=0, tolerance=tolerance,
                       initialize=final_initialize)

        self.is_fitted = True
        self.fit_observations = Y_to_fit
        self.fit_input = X_to_fit
        self.viterbi_states = self.model.most_likely_states(Y_to_fit[series_index], input=X_to_fit[series_index]) # conserve les régimes les + probables sur l'in-sample
        return self

    def get_transition_weights(self):
        """
        Renvoie les poids du modèle de transition.
        """
        return self.model.transitions.Ws