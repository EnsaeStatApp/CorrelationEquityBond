import numpy as np
from Allocator.base import Allocator

class TargetVolAllocator(Allocator):
    """
    Allocateur Target Volatility avec pondération par l'inverse de la variance.
    """

    def __init__(self, target_vol_annual: float = 0.10, max_leverage: float = 3.0):
        self.target_vol_annual = target_vol_annual
        self.max_leverage = max_leverage
        self._mu = None

    def set_mu(self, mu: np.ndarray):
        self._mu = mu

    def allocate(self, Sigma: np.ndarray) -> np.ndarray:
        n_assets = Sigma.shape[0]

        if Sigma is None or np.isnan(Sigma).any() or np.isinf(Sigma).any():
            # Si les données sont invalides, on renvoie des poids à zéro (100% Cash)
            return np.zeros(n_assets)

        # 1. Calcul de la répartition relative intelligente (Inverse de la Variance)
        # On extrait la diagonale (les variances de chaque actif)
        variances = np.diag(Sigma)
        
        # Poids inversement proportionnels au risque
        inv_variances = 1.0 / (variances + 1e-12)
        w_base = inv_variances / np.sum(inv_variances)

        # 2. Calcul de la volatilité mensuelle du portefeuille de base
        var_monthly = float(w_base.T @ Sigma @ w_base)
        vol_monthly = np.sqrt(max(var_monthly, 1e-12))
        
        # 3. Annualisation
        vol_annual = vol_monthly * np.sqrt(12)

        # 4. Calcul du levier pour atteindre la Target Vol
        exposure = self.target_vol_annual / vol_annual
        exposure = min(exposure, self.max_leverage)

        # On affiche uniquement tous les 12 mois pour ne pas saturer l'écran
        print(f"\n--- Diagnostic Stratégie ---")
        print(f"Vols annuelles par actif : {np.sqrt(variances * 12)}")
        print(f"Vol du panier de base : {vol_annual:.2%}")
        print(f"Levier calculé : {exposure:.2f}x")
        print(f"Poids finaux (SP, AAA, BAA, T10) : {w_base * exposure}")

        # 5. Poids finaux
        return w_base * exposure