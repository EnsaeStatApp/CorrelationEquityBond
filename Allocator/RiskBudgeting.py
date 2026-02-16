import numpy as np
from scipy.optimize import minimize
from Allocator.base import Allocator

class RiskBudgetingAllocator(Allocator):
    """
    Cette classe utilise la methode de riskbudgeting pour faire de l'alloc d'actif
    Idée : étant donné le budget de risque de chaque actif, on alloue des poids de manière à ce que la contribution 
    de chaque actif à la vol globale du portefeuille soit égale à son budget de risque (qui somment à 1)
    Pour cela, on résout un probleme de minimisation convexe dont la solution respecte l'equation de riskbudgeting
    """

    def __init__(self, risk_budgets: np.ndarray, tol: float = 1e-8, maxiter: int= 500):
        b = np.asarray(risk_budgets, dtype=float)

        if np.any(b <= 0):
            raise ValueError("Tous les budgets de risque doivent etre strict positifs")

        self.b = b / b.sum()
        self.tol = tol
        self.maxiter = maxiter

    def allocate(self, Sigma: np.ndarray) -> np.ndarray:
        Sigma = np.asarray(Sigma, dtype=float)
        n = Sigma.shape[0]

        # TO DO : verifier que sigma est bien def pos 

        def objective(x): # c'est cette fonction qu'on minimise (la solution vérifie l'équation de riskbudgeting)
            return 0.5 * x @ Sigma @ x - np.sum(self.b * np.log(x))

        def grad(x): # le gradient de la fonction (convexe)
            return Sigma @ x - self.b / x

       
        ### Initialiasation 
        x0 = 1 / np.sqrt(np.diag(Sigma))
        x0 /= x0.sum()


        bounds = [(1e-16, None)] * n

        res = minimize(
            objective,
            x0,
            jac=grad,
            method="L-BFGS-B",
            bounds=bounds,
            options=dict(ftol=self.tol, maxiter=self.maxiter)
        )

        if not res.success:
            raise RuntimeError(res.message)

        x = res.x
        w = x / x.sum()

        ### Verification de l'equation de riskbudgeting 

        Sigma_w = Sigma @ w
        port_vol = (w @ Sigma_w)

        rc = w * Sigma_w                 # risk contributions
        rc_normalized = rc / port_vol    # contributions en part de risque

        err = np.max(np.abs(rc_normalized - self.b))

        if err > 1e-2: # TO DO : prendre quelle erreur ? 
            raise RuntimeError(
                f"Risk budgeting non satisfait : erreur max = {err:.2e}"
            )

        return w

