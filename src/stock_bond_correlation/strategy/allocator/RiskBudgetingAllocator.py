import numpy as np
from scipy.optimize import minimize

class RiskBudgetingAllocator:
    """
    Allocator class implementing the Risk Budgeting (Risk Parity) optimization logic.
    """
    
    @staticmethod
    def get_risk_budget_weights(sigma: np.ndarray, budgets: np.ndarray) -> np.ndarray:
        """
        Computes the portfolio weights where each asset contributes a specific 
        percentage to the total risk.

        Args:
            sigma (np.ndarray): Annualized covariance matrix (D, D).
            budgets (np.ndarray): Target risk contributions (must sum to 1).

        Returns:
            np.ndarray: Normalized portfolio weights.
        """
        n = len(budgets)
        budgets = np.array(budgets)

        # Objective function using log-utility for convexity (Spinu approach)
        def obj(w):
            return 0.5 * w @ sigma @ w - np.sum(budgets * np.log(w))

        # Positive bounds to ensure strictly positive weights
        bounds = [(1e-8, None)] * n
        x0 = np.ones(n) / n 

        # L-BFGS-B is used for its efficiency in bounded optimization
        res = minimize(obj, x0, method='L-BFGS-B', bounds=bounds)

        # Normalize to ensure sum(w) = 1
        w_optimized = res.x
        return w_optimized / np.sum(w_optimized)