import numpy as np
from RegimeDetector.base import RegimeDetector


class BaseRegimeDetector(RegimeDetector):
    """
    Cette classe (abstraite toujours) donne le comportement commun de chacun de nos détecteurs de régimes. 
    But : Ne pas réimplémenter à chaque fois les mêmes fonctions. 
    """
    def conditional_covariance(self, t):
        """
        Calcule la matrice de covariance conditionnelle à un instant t par la formule de la covariance totale
        """
        # 1. Partie Intra-Régime
        Sigmas = self.regime_covariances()
        probs = self.regime_probabilities()[t] # vecteur des probabilités d'être dans chaque régime en t
        intra_cov = sum(p * S for p, S in zip(probs, Sigmas)) # somme des produits vectoriels ligne par ligne

        # 2. Partie Inter-Régime (Ajustement des moyennes)
        means = self.regime_means() # Liste de vecteurs mu_k
        mu_expected = sum(p * m for p, m in zip(probs, means))
        
        inter_cov = np.zeros_like(intra_cov)
        for k in range(len(probs)):
            diff = (means[k] - mu_expected).reshape(-1, 1)
            inter_cov += probs[k] * (diff @ diff.T)
            
        return intra_cov + inter_cov

    def regime_correlations(self):
        """
        Corrélations par régime : 
        regime_correlations[k] = matrice de corrélations du régime k 
        """
        corrs = []
        for S in self.regime_covariances():
            d = np.sqrt(np.diag(S)) # vecteurs des vols pour ce régime
            product_vols_matrix = np.outer(d, d) # product_vols_matrix[i,j] = sigma[i] sigma[j]
            corrs.append(S / product_vols_matrix) 
        return corrs
