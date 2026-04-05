import numpy as np
import matplotlib.pyplot as plt
from RegimeDetector.HMMMarketWithMacroTransitionsRegimeDetector import HMMMarketWithMacroTransitionsRegimeDetector
from typing import List

class TransitionExplainer:
    """
    Cette classe permet d'expliquer la différence entre la matrice de transition entre la date t et celle en t+1 en 
    attribuant une contribution à chaque covariable par la méthode des Integrated Gradients.

    NB : cet outil n'a du sens que pour des IOHMMs
    Attention : on suppose toujours que l'utilisateur laggue ses covariables (ordre 1) pour assurer la causalité
    """
    def __init__(self, detector : HMMMarketWithMacroTransitionsRegimeDetector, feature_names : List[str]):
        """
        Paramètres : 
        - detector : detector de régimes avec des covariables  
        - feature_names : liste des noms des covariables
        """
        self.detector = detector
        self.feature_names = feature_names

    def explain_delta(self, X_t : np.ndarray, X_tp1 : np.ndarray, from_state : int, to_state : int , steps : int = 50):
        """
        Explique 
        delta =
        P(z_t+1 = to_state | z_t = from_state, X_{t+1}) - P(z_t+1 = to_state | z_t = from_state, X_t)
        
        Attribue les contributions de chaque variable macro via Integrated Gradients 

        Paramètres : 
        - X_t : vecteur de taille M (nombre de covariables) contenant la valeur des covariables à la date t
        - X_tp1 : vecteur de taille M (nombre de covariables) contenant la valeur des covariables à la date t+1
        - from_state : état d'origine 
        - to_state : état d'arrivée
        - steps : nombre de petits pas entre X_t et X_{t+1} (précision de l'intégrale, + grand => + précis mais + coûteux)
        """
        # 1. Extraction des paramètres
        log_Ps, Ws = self.detector.get_transition_params()
        
        # delta_X : variation des variables macro entre t et t+1
        delta_X = X_tp1 - X_t
        
        # Interpolation (chemin de X_t à X_tp1)
        alphas = np.linspace(0, 1, steps)
        avg_grads = np.zeros_like(X_t)
        
        for a in alphas:
            x_interp = X_t + a * delta_X # chemin intermédiaire
            
            # Calcul manuel du Softmax pour la ligne 'from_state'
            logits = log_Ps[from_state] + (Ws @ x_interp)
            probs = np.exp(logits - np.max(logits)) # transfo en proba via softmax
            probs /= np.sum(probs)
            
            p_target = probs[to_state]
            expected_W = probs @ Ws
            grad = p_target * (Ws[to_state] - expected_W) # Gradient du Softmax : P_j * (W_target - sum(P_l * W_l))
            
            avg_grads += grad
            
        attributions = delta_X * (avg_grads / steps) # Attribution finale (IG)
        return attributions

    def plot_explanation(self, attributions : np.ndarray, from_regime : int, to_regime : int, date_label : str):
        """
        Trace en batons les contributions des variables macros calculées par explain_delta

        Paramètres : 
        - attributions : vecteur de taille M qui est le résultat direct de explain_delta
        - froùm_regime : état d'origine
        - to_regime : état d'arrivée
        - date_label : date à laquelle le changement a été observé (t+1) 
        """
        plt.figure(figsize=(10, 6))
        # On trie pour avoir les variables les plus importantes en haut
        indices = np.argsort(np.abs(attributions))
        
        plt.barh(np.array(self.feature_names)[indices], 
                 attributions[indices], 
                 color=['red' if x < 0 else 'green' for x in attributions[indices]])
        
        plt.axvline(0, color='black', lw=0.8)
        plt.title(f"Transition {from_regime} -> {to_regime} le {date_label}\n"
                  f"Attribution du changement de probabilité par variable macro")
        plt.xlabel("Contribution à la $\Delta$ Probabilité")
        plt.grid(axis='x', alpha=0.3)
        plt.tight_layout()
        plt.show()