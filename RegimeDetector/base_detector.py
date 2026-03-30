import numpy as np
from RegimeDetector.base import RegimeDetector


class BaseRegimeDetector(RegimeDetector):
    """
    Cette classe (abstraite toujours) donne le comportement commun de chacun de nos détecteurs de régimes. 
    But : Ne pas réimplémenter à chaque fois les mêmes fonctions. 
    """
    def conditional_covariance(self, t):
        """
        Covariance conditionnelle à t :
        sigma_t = sum_k (p_k(t) sigma_k)
        """
        Sigmas = self.regime_covariances()
        probs = self.regime_probabilities()[t] # vecteur des probabilités d'être dans chaque régime en t

        return sum(p * S for p, S in zip(probs, Sigmas)) # somme des produits vectoriels ligne par ligne

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
