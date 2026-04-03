import pandas as pd
import numpy as np
from RegimeDetector.EmissionStatisticalAnalyzer import EmissionStatisticalAnalyzer
from RegimeDetector.TransitionStatisticalAnalyzer import TransitionStatisticalAnalyzer
from RegimeDetector.base import RegimeDetector
from typing import List, Dict, Tuple
from IPython.display import display


class HMMCoherenceChecker:
    """
    Tour de contrôle pour valider la robustesse d'un fit HMM.
    Automatise les tests de séparabilité, stabilité numérique et persistance.
    """
    def __init__(self, detector: RegimeDetector, log_returns: np.ndarray, asset_names: List[str] = None, covariate_names: List[str] = None,
                 min_duration: float = 3.0, min_freq: float = 0.05, sig_threshold: float = 0.5, sig_max_cond_number: float = 1e8,
                 min_confidence: float = 0.7, weights: Dict[str, float] = None):
        """
        Paramètres :
        - detector : Instance de RegimeDetector (Market ou Macro).
        - log_returns : Rendements financiers (T, D) pour les tests d'émissions.
        - asset_names : Noms des actifs.
        - covariate_names : Noms des variables macro (si IOHMM).
        - min_duration : Durée moyenne minimale acceptée (ex: 3 mois).
        - min_freq : Fréquence long-terme minimale (ex: 5%).
        - sig_threshold : Ratio minimal de tests de séparabilité significatifs.
        - min_confidence : Moyenne minimale de l'indice de confiance (1-Entropie).
        """
        self.detector = detector
        self.log_returns = log_returns
        self.asset_names = asset_names
        self.covariate_names = covariate_names

        # Seuils de validation
        self.min_duration = min_duration
        self.min_freq = min_freq
        self.sig_threshold = sig_threshold
        self.sig_max_cond_number = sig_max_cond_number
        self.min_confidence = min_confidence

        # Poids du score global (somme = 1)
        self.weights = weights or {
            "separability": 0.3,
            "stability": 0.2,
            "persistence": 0.2,
            "confidence": 0.3
        }

    def _refresh_analyzers(self):
        """
        Instancie les analyseurs avec l'état actuel (post-fit) du détecteur.
        """
        if not self.detector.is_fitted:
            raise ValueError("Le détecteur doit être fitté avant d'utiliser le checker.")

        # On récupère les états Viterbi pour l'analyseur d'émissions
        states = self.detector.viterbi_states

        stat_ana = EmissionStatisticalAnalyzer(
            log_returns=self.log_returns,
            states=states,
            asset_names=self.asset_names
        )
        trans_ana = TransitionStatisticalAnalyzer(detector=self.detector)

        return stat_ana, trans_ana

    def check_separability(self):
        """
        Vérifie si les régimes sont statistiquement distincts (Vol & Correl) grâce à EmissionStatisticalAnalyzer
        """
        stat_ana, _ = self._refresh_analyzers()
        levene = stat_ana.test_pairwise_levene()
        fisher = stat_ana.test_fisher_z()

        total = len(levene) + len(fisher)
        sigs = levene['significant'].sum() + fisher['significant'].sum()
        ratio = sigs / total if total > 0 else 0

        return {"ratio": ratio, "is_ok": ratio >= self.sig_threshold}

    def check_stability(self, ref_idx: int = 0):
        """
        Analyse la forme du maximum de vraisemblance (Hessienne) grâce à TransitionStatisticalAnalyzer 

        Paramètre : 
        - ref_idx : indice de réf pour calculer l'incertitude dansn TransitionStatisticalAnalyzer
        """
        _, trans_ana = self._refresh_analyzers()
        # On calcule l'inférence numérique
        _, diag = trans_ana.compute_transition_inference(
            covariate_names=self.covariate_names,
            ref_idx=ref_idx
        )

        eigvals = diag['eigvals']
        cond_num = diag['cond_num']

        # Un modèle est stable si eigvals de Fisher > 0 et CondNum raisonnable
        is_stable = (np.all(eigvals > 0)) and (cond_num < self.sig_max_cond_number)

        return {"cond_num": cond_num, "is_stable": is_stable, "min_eig": np.min(eigvals)}


    def report_emission_stats(self, n_boot : int = 5000, alpha : float = 0.05, annualise : float = np.sqrt(12)):
        """
        Affiche le diagnostic complet des émissions : 
        Volatilités et Corrélations avec Intervalles de Confiance (Bootstrap).

        Paramètres : 
        - n_boot : nombre d'itérations dans les bootstrap
        - alpha : seuil de significativité
        - annalualise : facteur d'anualisation 
        """
        stat_ana, _ = self._refresh_analyzers()
        
        print("\n" + "="*70)
        print("DIAGNOSTIC COMPLET DES ÉMISSIONS (BOOTSTRAP & TESTS)")
        print("="*70)
        
        # 1. Récupération du DataFrame global de Bootstrap
        boot_df = stat_ana.get_bootstrap_ci_df(n_boot=n_boot, alpha=alpha, annualise=annualise)
        
        # Affichage propre des vols
        print("\n1. VOLATILITÉS ANNUELLES AVEC INTERVALLES DE CONFIANCE :")
        vols_boot = boot_df.xs('Volatility', level='Type')
        display(vols_boot.round(4))

        # Affichage propre des correls
        print("\n2. CORRÉLATIONS AVEC INTERVALLES DE CONFIANCE :")
        corrs_boot = boot_df.xs('Correlation', level='Type')
        display(corrs_boot.round(4))

        # Affichage des résultats aux tests de significativité
        print("\n3. TESTS DE SÉPARABILITÉ (TOUTES LES PAIRES) :")
        
        levene = stat_ana.test_pairwise_levene(alpha=alpha)
        fisher = stat_ana.test_fisher_z(alpha=alpha)
        
        # on concatène les tests importants
        print("- Significativité des Volatilités (Levene) :")
        display(levene.sort_values("p_value").round(4))
        
        print("\n- Significativité des Corrélations (Fisher-Z) :")
        display(fisher.sort_values("p_value").round(4))

    def report_transition_stats(self, ref_idx : int = 0, inputs : np.ndarray = None):
        """
        Affiche le diagnostic complet des transitions :
        Inférence sur les poids (p-values), matrices de transition et durées.

        Paramètres : 
        - ref_idx : indice de réf pour calculer les incertitudes avec la hessienne
        - inputs : covariables pouvant impacter les transitions
        """
        _, trans_ana = self._refresh_analyzers()

        print("\n" + "="*70)
        print("DIAGNOSTIC DES TRANSITIONS (DYNAMIQUE DU MODÈLE)")
        print("="*70)

        # 1. Inférence statistique (Hessienne / Fisher Info)
        print(f"\n1. SIGNIFICATIVITÉ DES PARAMÈTRES (Réf: Régime {ref_idx}) :")
        inf_df, _ = trans_ana.compute_transition_inference(
            covariate_names=self.covariate_names,
            ref_idx=ref_idx
        )
        display(inf_df.round(4))

        # 2. Dynamique moyenne
        print("\n2. DYNAMIQUE MOYENNE ET PERSISTANCE :")
        avg_A, durations, stationary = trans_ana.get_average_transition_dynamics(inputs=inputs)

        dyn_df = pd.DataFrame({
            "Durée_Moyenne (mois)": durations,
            "Prob_Stationnaire": stationary,
            "Prob_Rester": np.diag(avg_A)
        })
        display(dyn_df.round(4))

    def compute_coherence_score(self, inputs: np.ndarray = None):
        """
        Calcule un score pondéré de 0 à 1 reflétant la qualité globale du fit.
        """
        sep = self.check_separability()
        stab = self.check_stability()

        _, trans_ana = self._refresh_analyzers()
        _, durations, stationary = trans_ana.get_average_transition_dynamics(inputs=inputs)

        # Confiance via l'entropie (méthode héritée de BaseRegimeDetector)
        avg_conf = np.mean(self.detector.compute_confidence_index(series_index=0))

        # --- Normalisation des sous-scores ---
        s_sep = sep['ratio']

        # Stabilité (échelle logarithmique pour le Condition Number)
        log_cond = np.log10(max(stab['cond_num'], 1.0))
        s_stab = max(0, 1 - (log_cond / 12.0))
        if not stab['is_stable']: s_stab *= 0.1 # Pénalité critique

        # Persistance (Durée et Fréquence)
        s_dur = min(durations.min() / self.min_duration, 1.0)
        s_freq = min(stationary.min() / self.min_freq, 1.0)
        s_pers = (s_dur + s_freq) / 2

        s_conf = avg_conf

        score = (
            self.weights["separability"] * s_sep +
            self.weights["stability"] * s_stab +
            self.weights["persistence"] * s_pers +
            self.weights["confidence"] * s_conf
        )

        details = {
            "sep_ratio": s_sep,
            "is_stable": stab['is_stable'],
            "min_dur": durations.min(),
            "min_freq": stationary.min(),
            "avg_conf": avg_conf,
            "durations": durations,
            "stationary": stationary
        }

        return score, details

    def final_verdict(self, inputs: np.ndarray = None):
        """
        Affiche un rapport complet et tranche sur la validité du modèle.
        """
        score, d = self.compute_coherence_score(inputs=inputs)

        # Critères Go/No-Go
        c_sep = d["sep_ratio"] >= self.sig_threshold
        c_stab = d["is_stable"]
        c_pers = (d["min_dur"] >= self.min_duration) and (d["min_freq"] >= self.min_freq)
        c_conf = d["avg_conf"] >= self.min_confidence

        print("\n" + "="*50)
        print(f"RAPPORT DE COHÉRENCE HMM : {score:.2%}")
        print("="*50)
        print(f"1. SÉPARABILITÉ : {d['sep_ratio']:.1%} (Seuil: {self.sig_threshold:.0%}) -> {'OK' if c_sep else 'FAIBLE'}")
        print(f"2. STABILITÉ    : {'VALIDE' if c_stab else 'ÉCHEC (Singularité)'}")
        print(f"3. PERSISTANCE  : {d['min_dur']:.1f} mois (min {self.min_duration}) -> {'OK' if c_pers else 'TROP COURT'}")
        print(f"4. CONFIANCE    : {d['avg_conf']:.1%} (Entropie) -> {'SOLIDE' if c_conf else 'HÉSITANT'}")
        print("-" * 50)

        df_dyn = pd.DataFrame({"Durée": d['durations'], "Fréq": d['stationary']})
        print("DYNAMIQUE DES RÉGIMES :")
        print(df_dyn.round(3).to_string())

        if c_sep and c_stab and c_pers and c_conf:
            print("\nVERDICT : MODÈLE VALIDÉ (Robuste)")
            return score
        else:
            print("\nVERDICT : MODÈLE REJETÉ (Instable ou non-significatif)")
            return 0.0 # pour l'instant on renvoie 0