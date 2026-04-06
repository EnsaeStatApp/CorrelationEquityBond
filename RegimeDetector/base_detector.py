import numpy as np
from RegimeDetector.base import RegimeDetector
import ssm
from typing import List

class BaseRegimeDetector(RegimeDetector):
    """
    Cette classe (abstraite toujours) donne le comportement commun de chacun de nos détecteurs de régimes. 
    
    But : Ne pas réimplémenter à chaque fois les mêmes fonctions 

    NB : tout modele qui hérite de cette classe est de fait supposée utilisatrice de ssm
    """

    def conditional_covariance(self, probs: np.ndarray, log_returns: np.ndarray = None) -> np.ndarray:
        """
        Calcule la covariance totale par la formule de la covariance totale

        Paramètres :
        - probs : tableaux des probabilités (probs[t] = probabilité en t)
        - log_returns (optionnel) : tableaux des logs returns dans le cas d'un calcul des métriques des régimes dynamiques
        """
        means = np.array(self.regime_means(log_returns)) # (K, D) ou (T, K, D)
        covs = np.array(self.regime_covariances(log_returns)) # (K, D, D) ou (T, K, D, D)
        
        T, K = probs.shape
        D = means.shape[-1]
        total_covs = np.zeros((T, D, D))

        for t in range(T):
            p_t = probs[t] # Vecteur (K,)
            
            # 1. Moyenne estimée = sum_k (p_k * mu_kt)
            m_t = means if means.ndim == 2 else means[t] # gestion cas statique ou dynamique
            mu_bar_t = p_t @ m_t 

            # 2. Partie intra-régime
            c_t = covs if covs.ndim == 3 else covs[t] # gestion cas statique ou dynamique
            intra_t = np.average(c_t, weights=p_t, axis=0) # moyenne des covs pondérée par les probas

            # 3. Partie inter-régime
            inter_t = np.zeros((D, D))
            for k in range(K):
                diff = m_t[k] - mu_bar_t
                inter_t += p_t[k] * np.outer(diff, diff)

            total_covs[t] = intra_t + inter_t
            
        return total_covs
    

    def regime_correlations(self, log_returns: np.ndarray = None):
        """
        Calcule les corrélations de chaque régime k

        Paramètre : 
        - log_returns (Optionnel) : tableaux des logs returns dans le cas d'un calcul des métriques des régimes dynamiques
        """
        sigmas = np.array(self.regime_covariances(log_returns))
        
        # CAS STATIQUE (K, D, D)
        if sigmas.ndim == 3:    
            corrs = []
            for k in range(len(sigmas)):
                S = sigmas[k]
                vols = np.sqrt(np.diag(S))
                inv_v = 1.0 / (vols + 1e-16)
                R = S * np.outer(inv_v, inv_v)
                corrs.append(R)
            return corrs
        
        # CAS DYNAMIQUE (T, K, D, D)
        else:
            T, K, D, _ = sigmas.shape
            corrs_dynamic = np.zeros_like(sigmas)
            for t in range(T):
                for k in range(K):
                    S = sigmas[t, k]
                    vols = np.sqrt(np.diag(S))
                    inv_v = 1.0 / (vols + 1e-16)
                    corrs_dynamic[t, k] = S * np.outer(inv_v, inv_v)
            return corrs_dynamic
        
    def compute_confidence_index(self, series_index: int = 0):
        """
        Calcule l'indice de confiance du modèle à chaque date via l'entropie de Shannon.
        
        Résultat :
        - 1.0 : Certitude absolue (une probabilité est à 100%, les autres à 0%).
        - 0.0 : Incertitude totale (tous les régimes ont une probabilité de 1/K).
        """
        if not self.is_fitted:
            raise ValueError("Modèle non fitté.")
            
        # 1. Récupérer les probabilités (T, K)
        probs = self.regime_probabilities(series_index)
        
        # 2. Calcul de l'entropie H = -sum(p * log(p))
        # On ajoute un epsilon pour éviter log(0)
        probs = np.maximum(probs, 1e-12)
        entropy = -np.sum(probs * np.log(probs), axis=1)
        
        # 3. Normalisation : l'entropie max d'un système à K états est log(K)
        # On transforme l'entropie en "Index de Confiance" (0 à 1)
        max_entropy = np.log(self.K)
        confidence = 1 - (entropy / max_entropy)
        
        return confidence

    def get_params(self):
        """
        Renvoie les paramètres du modèle (Init + Trans + Obs)
        """
        if not self.is_fitted:
            raise ValueError("Modèle non fitté.")
        return self.model.params # Renvoie un tuple complet de ssm

    def set_params(self, params):
        """
        Injecte des paramètres dans le modèle
        """
        self.model.params = params
        self.is_fitted = True

    def get_transition_params(self):
        """
        Renvoie les paramètres des transitions (log_Ps et Ws)

        Si M = 0 (i.e. HMM classique) renvoie log_Ps, Ws = vecteur nul. Ainsi les paramètres de transitions des modèles ne sont plus que
        les intercepts (les log_Ps)
        """
        log_Ps = self.model.transitions.log_Ps.copy()
        if self.M > 0: # i.e. on a des covariables i.e. on a des poids
            Ws = self.model.transitions.Ws.copy()
        else:
            Ws = np.zeros((self.K, self.M)) # Renvoie une matrice vide (K x 0) pour ne pas casser les concaténations
        return log_Ps, Ws

    def update_transition_params(self, log_Ps : np.ndarray, Ws : np.ndarray):
        """
        Permet de modifier les paramètres de transitions (modifie les poids uniquement lorsqu'ils existent i.e. avec des covariables)

        Paramètres :
        - log_Ps : Matrice (de taille K*K) des log-probabilités de transition de base (intercepts)
        - Ws : Matrice (de taille K*M) des poids associés aux covariables macro (poids Log Regression).
        """
        self.model.transitions.log_Ps = log_Ps
        if self.M > 0: # on a des covariables donc des poids donc on met à jour
            self.model.transitions.Ws = Ws

    def compute_ll(self, Y : List[np.ndarray], X : List[np.ndarray]):
        """
        Calcule la vraisemblance sur le jeu de données en entrée

        Paramètres :
        """
        return self.model.log_likelihood(Y, inputs=X)

    def get_transition_matrices(self, X: np.ndarray = None):
        """
        Calcule les matrices de transition selon la logique demandée :
        - Si M == 0 : Modèle stationnaire, X est ignoré.
        - Si M > 0 : Modèle Macro, X est obligatoire.
        """
        if not self.is_fitted:
            raise ValueError("Modèle non fitté.")

        if self.M == 0: # HMM
            P = self.model.transitions.transition_matrix
            return P

        else: # cas M > 0 (IOHMM Macro)
            if X is None:
                raise ValueError(f"Ce modèle a été fitté avec M={self.M} variables macro, donc vous devez fournir un X pour calculer les transitions.")

            X_input = X
            T = X_input.shape[0]
            dummy_Y = np.zeros((T, self.D)) # On crée un dummy Y minimal pour satisfaire la signature technique

            return self.model.transitions.transition_matrices( # respecte la signature de ssm
                data=dummy_Y,
                input=X_input,
                mask=np.ones((T, self.D), dtype=bool),
                tag=None
            )
        
    def predict_probabilities(self, observations: np.ndarray, inputs: np.ndarray = None, oos_start : int = 0, oos_end : int = None):
        """
        Calcule les probabilités prédictives sur tout l'in-sample + l'OOS (pour le contexte) mais ne renvoie les probabilités que
        sur l'OOS

        A la ligne t, P(z_t | Y_{1:t-1}, X_{1:t-1})

        Paramètres :
        - observations : tableaux des logs returns sur l'in-sample + l'OOS (pour contexte)
        - inputs : tableaux des covariables impactant les transitions sur l'in-sample + l'OOS (pour contexte)
        - oos_start : indice de début de l'OOS
        - oos_end : indice de fin de l'OOS
        """
        if not self.is_fitted:
            raise ValueError("Modèle non fitté.")

        all_preds = self.model.filter(observations, input=inputs) # all_preds[t] = P(z_t | Y_{1:t-1}, X_{1:t-1})

        return all_preds[oos_start:oos_end]