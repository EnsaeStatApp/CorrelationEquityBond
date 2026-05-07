import numpy as np
import matplotlib.pyplot as plt
from RegimeDetector.HMMMarketWithMacroTransitionsRegimeDetector import HMMMarketWithMacroTransitionsRegimeDetector
from typing import List

class TransitionExplainer:
    """
    Class designed to explain the change in the transition matrix between time $t$ and $t+1$.
    
    It attributes a contribution to each macro-economic covariate using the 
    Integrated Gradients (IG) method.

    Note:
        This tool is specifically designed for Input-Driven HMMs (IOHMMs).
        It assumes that covariates are already lagged (order 1) to ensure causality.
    """
    def __init__(self, detector : HMMMarketWithMacroTransitionsRegimeDetector, feature_names : List[str]):
        """
        Initializes the explainer with a fitted detector.

        Args:
            detector (HMMMarketWithMacroTransitionsRegimeDetector): A regime detector with covariates.
            feature_names (List[str]): List of names for the macro covariates.
        """
        self.detector = detector
        self.feature_names = feature_names

    def explain_delta(self, X_t : np.ndarray, X_tp1 : np.ndarray, from_state : int, to_state : int , steps : int = 50):
        """
        Explains the transition probability change:
        $$\Delta = P(z_{t+1} = \text{to\_state} \mid z_t = \text{from\_state}, X_{t+1}) - P(z_{t+1} = \text{to\_state} \mid z_t = \text{from\_state}, X_t)$$

        Attributes the contributions of each macro variable via Integrated Gradients.

        Args:
            X_t (np.ndarray): Vector of size $M$ containing covariate values at time $t$.
            X_tp1 (np.ndarray): Vector of size $M$ containing covariate values at time $t+1$.
            from_state (int): Origin state index.
            to_state (int): Destination state index.
            steps (int): Number of steps for integral approximation (higher is more accurate).

        Returns:
            np.ndarray: Attribution vector of size $M$.
        """
        # 1. Extract transition parameters
        log_Ps, Ws = self.detector.get_transition_params()

        # delta_X: Variation in macro variables between t and t+1
        delta_X = X_tp1 - X_t

        # Interpolation (path from X_t to X_tp1)
        alphas = np.linspace(0, 1, steps)
        avg_grads = np.zeros_like(X_t)

        for a in alphas:
            x_interp = X_t + a * delta_X # Intermediate point on the path

            # Manual Softmax calculation for the 'from_state' row
            logits = log_Ps[from_state] + (Ws @ x_interp)
            probs = np.exp(logits - np.max(logits)) # Softmax transformation
            probs /= np.sum(probs)

            p_target = probs[to_state]
            expected_W = probs @ Ws
            # Softmax Gradient: P_j * (W_target - sum(P_l * W_l))
            grad = p_target * (Ws[to_state] - expected_W) 

            avg_grads += grad

        # Final attribution using the Integrated Gradients formula
        attributions = delta_X * (avg_grads / steps) 
        return attributions

    def plot_explanation(self, attributions : np.ndarray, from_regime : int, to_regime : int, date_label : str, from_regime_name, to_regime_name):
        """
        Plots the macro-variable contributions calculated by explain_delta using a bar chart.

        Args:
            attributions (np.ndarray): Result vector from explain_delta.
            from_regime (int): Origin state index.
            to_regime (int): Destination state index.
            date_label (str): Date at which the change was observed ($t+1$).
            from_regime_name (str): Descriptive name of the origin regime.
            to_regime_name (str): Descriptive name of the destination regime.
        """
        plt.figure(figsize=(10, 6))
        # Sort variables to have the most impactful ones at the top
        indices = np.argsort(np.abs(attributions))

        plt.barh(np.array(self.feature_names)[indices],
                 attributions[indices],
                 color=['red' if x < 0 else 'green' for x in attributions[indices]])

        plt.axvline(0, color='black', lw=0.8)
        plt.title(f"Transition {from_regime_name} -> {to_regime_name} on {date_label}\n"
                  f"Macro Variable Attribution of the Probability Shift")
        plt.xlabel("Contribution to $\Delta$ Probability")
        plt.grid(axis='x', alpha=0.3)
        plt.tight_layout()
        plt.show()