import numpy as np
from RegimeDetector.base import RegimeDetector

class BaseRegimeDetector(RegimeDetector):
    """
    Cette classe (abstraite toujours) donne le comportement commun de chacun de nos détecteurs de régimes. 
    
    But : Ne pas réimplémenter à chaque fois les mêmes fonctions. 
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
