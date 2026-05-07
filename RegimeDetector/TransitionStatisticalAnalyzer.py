import numpy as np
import pandas as pd
from scipy import stats as sp_stats
from typing import List
from RegimeDetector.base_detector import BaseRegimeDetector

class TransitionStatisticalAnalyzer:
    """
    Classe qui effectue des tests statistiques descriptifs sur les paramètres de transition d'un detector

    - Renvoie les paramètres de transitions de chaque régime (incertept + poids associés aux covariables si présentes)
    - Renvoie les standard errors associés à ces paramètres en calculant numériquement l'information de fisher du modèle
    - Renvoie les valeurs propres de la Hessienne au point optimal (uniquement sur les paramètres de transitions pour simplifier)
    pour déterminer "la forme du sommet du MLE"
    - Renvoie des matrices de transitions associés aux covariables si présentes, sinon sont constantes
    - Renvoie la durée moyenne de chaque régime liée à une matrice de transition
    """

    def __init__(self, detector : BaseRegimeDetector):
        """
        Paramètre :
        - detector : un détecteur de régime général (part du principe que detector.fit_data sont des log returns)
        """
        self.detector = detector
        self.K = detector.K
        self.M = detector.M

    def compute_transition_inference(self, covariate_names : List[str] = None, series_index : int = 0, alpha : float = 0.05,
                                     ref_idx : int = 0, epsilon : float = 1e-4):
        """
        Calcule les incertitudes sur les poids des covariables sur les probabilités de transitions

        Pour cela, on calcule l'information de Fisher en approximant la Hessienne de la LL par différences finies

        Paramètres :
        - covariate_names : nom des covariables
        - series_index : index de la série sur laquelle on veut regarder les paramètres (rappel : self.detector.fit_data est une liste de tableaux de log returns)
        - alpha : seuil de significativité
        - ref_idx : index de référence du régime par rapport auquel on regarde la significativité
        - epsilon : pas de différences finies
        """

        # 1. On sauvegarde les données car à la fin on veut que le modèle revienne à son état d'origine
        data = np.array(self.detector.fit_observations[series_index])
        if self.detector.fit_input is not None:
            inp = np.array(self.detector.fit_input[series_index])
        else:
            inp = None

        # 2. On sauvegarde les paramètres actuels de transitions car à la fin on veut que le modèle revienne à son état d'origine
        log_Ps_orig, Ws_orig = self.detector.get_transition_params()

        # 3. Ici, on va centrer les paramètres autour d'un régime de référence
        reference_column = log_Ps_orig[:, [ref_idx]] # on prend la colonne du régime en question (avec la liste pour garder la matrice de K*1)
        log_Ps_centered = log_Ps_orig - reference_column
        log_Ps_r = np.delete(log_Ps_centered, ref_idx, axis=1)

        reference_column = Ws_orig[[ref_idx], :] # pareil
        Ws_centered = Ws_orig - reference_column
        Ws_r = np.delete(Ws_centered, ref_idx, axis=0)

        n_logPs = self.K * (self.K - 1)
        p_red = n_logPs + (self.K - 1) * self.M # nombre de paramètres réduits de ceux du régime de réf
        params_opt = np.concatenate([log_Ps_r.flatten(), Ws_r.flatten()]) # Pour pouvoir calculer la hessienne, on applatit les paramètres en 2 listes séparées puis on concat en une grosse variable theta

        # 4. LL
        def get_ll(params_flat : np.ndarray):
            """
            Prends un vecteur de paramètres puis reconstruit les log_Ps et Ws et calcule la LL

            Paramètre :
            - params_flat : vecteur de paramètres
            """
            lP_r = params_flat[:n_logPs].reshape(self.K, self.K - 1)
            W_r = params_flat[n_logPs:].reshape(self.K - 1, self.M)

            lP_f = np.concatenate([lP_r[:, :ref_idx], np.zeros((self.K, 1)), lP_r[:, ref_idx:]], axis=1)
            W_f = np.concatenate([W_r[:ref_idx, :], np.zeros((1, self.M)), W_r[ref_idx:, :]], axis=0)

            self.detector.update_transition_params(lP_f, W_f) # Mise à jour via le wrapper

            return self.detector.compute_ll([data], [inp]) # Calcul de la LL via le wrapper

        # 5. Calcul de la hessienne par diff finies (différence finie centrée de second ordre)
        ll_base = get_ll(params_opt)
        H = np.zeros((p_red, p_red))
        for i in range(p_red): # Diagonale
            p_plus = params_opt.copy()
            p_plus[i] += epsilon
            p_minus = params_opt.copy()
            p_minus[i] -= epsilon
            ll_p = get_ll(p_plus)
            ll_m = get_ll(p_minus)

            H[i, i] = (ll_p - 2 * ll_base + ll_m) / (epsilon ** 2) # formule diff centrée

        for i in range(p_red): # Termes croisés
            for j in range(i + 1, p_red): # on commence à i+1 car H est symétrique
                p_pp = params_opt.copy()
                p_pp[i] += epsilon
                p_pp[j] += epsilon
                p_mm = params_opt.copy()
                p_mm[i] -= epsilon
                p_mm[j] -= epsilon
                p_pm = params_opt.copy()
                p_pm[i] += epsilon
                p_pm[j] -= epsilon
                p_mp = params_opt.copy()
                p_mp[i] -= epsilon
                p_mp[j] += epsilon
                val = (get_ll(p_pp) - get_ll(p_pm) - get_ll(p_mp) + get_ll(p_mm)) / (4 * epsilon ** 2) # formule diff centrée
                H[i, j], H[j, i] = val, val

        self.detector.update_transition_params(log_Ps_orig, Ws_orig)  # On remet les paramètres initiaux

        # 6. Analyse du spectre (pour stabilité numérique) et inversion de la matrice
        I_obs = - H
        eigvals = np.linalg.eigvalsh(I_obs)
        cond_num = eigvals.max() / (np.abs(eigvals.min()) + 1e-12)

        try: # pour inverser
            I_inv = np.linalg.inv(I_obs)
        except np.linalg.LinAlgError:
            I_inv = np.linalg.pinv(I_obs)

        # 7. Calcul des standard errors
        se = np.sqrt(np.maximum(np.diag(I_inv), 1e-12))
        se_Ws_r = se[n_logPs:].reshape(self.K - 1, self.M)
        se_lPs_r = se[:n_logPs].reshape(self.K, self.K - 1)

        z_Ws = Ws_r / se_Ws_r
        pv_Ws = 2 * sp_stats.norm.sf(np.abs(z_Ws))

        # Préparation du dataframe final
        rows = []
        regimes_dest = [r for r in range(self.K) if r != ref_idx]
        cov_names = covariate_names or [f"X{m}" for m in range(self.M)]

        # A. Ajout des Intercepts (log_Ps)
        for i, k_origin in enumerate(range(self.K)):
            for j, k_dest in enumerate(regimes_dest):
                coef = log_Ps_r[i, j]
                stderr = se_lPs_r[i, j]
                z = coef / stderr
                pval = 2 * sp_stats.norm.sf(np.abs(z))
                rows.append({
                    'Type': 'Intercept',
                    'Transition': f"R{k_origin} -> R{k_dest}",
                    'Variable': 'Base_Prob',
                    'Coef': coef, 'Std_Err': stderr, 'P_Value': pval,
                    'Significatif': pval < alpha
                })

        # B. Ajout des Poids Macro (Ws)
        for i, k_dest in enumerate(regimes_dest):
            for m in range(self.M):
                coef = Ws_r[i, m]
                stderr = se_Ws_r[i, m]
                z = coef / stderr
                pval = 2 * sp_stats.norm.sf(np.abs(z))
                rows.append({
                    'Type': 'Macro_Weight',
                    'Transition': f"Vers R{k_dest} (vs R{ref_idx})",
                    'Variable': cov_names[m],
                    'Coef': coef, 'Std_Err': stderr, 'P_Value': pval,
                    'Significatif': pval < alpha
                })

        return pd.DataFrame(rows), {'eigvals': eigvals, 'cond_num': cond_num}


    def get_transition_matrix(self, inputs: np.ndarray = None, mode: str = "average"):
        """
        Calcule la matrice de transition.
        """
        matrices = self.detector.get_transition_matrices(inputs)

        # AJUSTEMENT : Si le détecteur renvoie une seule matrice (K, K) au lieu de (T, K, K)
        if matrices.ndim == 2:
            return matrices  # Déjà une matrice, rien à moyenner !

        # Si c'est une pile (T, K, K), on applique la logique demandée
        if mode == "average":
            return np.mean(matrices, axis=0)
        elif mode == "last":
            return matrices[-1]
        elif mode == "all":
            return matrices
        else:
            raise ValueError("Le mode doit être 'average', 'last' ou 'all'.")

    def get_expected_durations(self, transition_matrix: np.ndarray):
        """
        Calcule la durée attendue dans chaque régime à partir d'une matrice A
        Formule : 1 / (1 - P(rester)) (loi géometrique de paramètre transition_matrix[i, i] pour le régime i)

        Paramètres :
        - transition_matrix : matrice de transition
        """
        return 1 / (1 - np.diag(transition_matrix))

    def get_average_transition_dynamics(self, inputs: np.ndarray = None):
        """
        Renvoie : (Matrice Moyenne, Durées, Fréquences Long Terme)

        Paramètres :
        - inputs : tableaux de covariables impactant les transitions
        """
        # 1. On récupère la matrice moyenne
        avg_A = self.get_transition_matrix(inputs, mode="average")

        # 2. On calcule les durées attendues
        durations = self.get_expected_durations(avg_A)

        # 3. Calcul de la distribution stationnaire (Fréquences Long Terme)
        vals, vecs = np.linalg.eig(avg_A.T) # On cherche le vecteur propre à gauche associé à la valeur propre 1
        stationary = np.real(vecs[:, np.isclose(vals, 1)])
        stationary = stationary[:, 0] / stationary.sum() # Normalisation

        return avg_A, durations, stationary
   