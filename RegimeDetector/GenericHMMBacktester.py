import pandas as pd
import numpy as np
from tqdm import tqdm
import copy
from typing import List, Union, Dict
from RegimeDetector.base import RegimeDetector
from RegimeDetector.HMMCoherenceChecker import HMMCoherenceChecker

class GenericHMMBacktester:
    """
    Backtest sur une période donnée en mode expanding un detector quelconque. A chaque refit (param refit_freq), teste plusieurs initialisations et 
    utilise celui qui a eu le meilleur score du HMMCoherenceChecker
    """
    def __init__(self, detector: RegimeDetector, data_df: pd.DataFrame, obs_vars: List[str], asset_vars: List[str], 
                 input_vars: List[str] = None, refit_freq: int = 10, scaling_factor: float = 100.0, min_duration: float = 3.0, 
                 min_freq: float = 0.05, sig_threshold: float = 0.5, sig_max_cond_number: float = 1e8, 
                 min_confidence: float = 0.7, weights: Dict[str, float] = None):
        """
        Paramètres : 
        - detector : detecteur de régimes
        - data_df : df contenant les observations, les covariables/inputs et les log returns
        - obs_vars : nom des observations du HMM dans data_df
        - input_vars : nom des covariables/inputs dans data_df
        - asset_vars : nom des log returns dans data_df
        - refit_freq : fréquence de refit
        - scaling_factor : facteur de scaling des observations pour stabilité numérique
        - min_duration : paramètre de HMMCoherenceChecker
        - min_freq : paramètre de HMMCoherenceChecker
        - sig_threshold : paramètre de HMMCoherenceChecker
        - sig_max_cond_number : paramètre de HMMCoherenceChecker
        - min_confidence : paramètre de HMMCoherenceChecker
        - weights : paramètre de HMMCoherenceChecker
        """
        
        self.df = data_df
        self.obs_vars = obs_vars      
        self.asset_vars = asset_vars  
        self.input_vars = input_vars if input_vars is not None else []
        
        self.template = detector
        self.refit_freq = refit_freq
        self.scaling = scaling_factor
        
        self.min_duration = min_duration
        self.min_freq = min_freq
        self.sig_threshold = sig_threshold
        self.sig_max_cond_number = sig_max_cond_number
        self.min_confidence = min_confidence
        self.weights = weights 
        
        self.current_detector = None
        self.results = {} # stockage de la matrice de covariance prévue en date t connaissant tout jusqu'à t-1
        self.history_logs = [] # stockage de (date, méthode d'init, seed)
        self.regime_sigmas = {} # stockage des matrices de cov de chaque régime à chaque refit
        self.regime_probs = {} # stockage des vecteurs de probas (T, K)

    def _clone_detector(self, n_regimes: int, seed: int):
        """
        Crée une copie du detector 

        Paramètres :
        - n_regimes : nombre de régimes
        - seed : seed sur lequel on travaille
        """
        new_det = copy.deepcopy(self.template)
        new_det.__init__(n_regimes=n_regimes, n_dim=len(self.obs_vars), 
                         n_input=len(self.input_vars), random_state=seed) #
        return new_det

    def _create_checker(self, detector: RegimeDetector, observations_scaled: np.ndarray):
        """
        Instancie le checker avec les données à l'échelle du modèle.

        Paramètres : 
        - detector : detecteur de régimes
        - observations_scaled : observations (scaled) du HMM
        """
        return HMMCoherenceChecker(
            detector, observations_scaled, self.asset_vars, 
            min_duration=self.min_duration,
            min_freq=self.min_freq,
            sig_threshold=self.sig_threshold,
            sig_max_cond_number=self.sig_max_cond_number,
            min_confidence=self.min_confidence,
            weights=self.weights
        )

    def _find_best_model(self, observations_train: np.ndarray, inputs_train: np.ndarray = None):
        """
        Moteur de recherche hiérarchique : Warm Start -> K-means -> Random

        Paramètres :
        - observations_train : observations du HMM sur le dataset de train
        - inputs_train : covariables/inputs du HMM sur le dataset de train
        """
        seeds = [120, 42, 2024, 7, 88, 999, 13, 1, 555, 777] # serie de seeds diff
        
        for k in range(self.template.K, 1, -1):

            # recherche du meilleur score en fonction de self.weights
            best_det = None
            best_score = -np.inf
            best_metadata = None

            # 1. Warm start : on teste d'initialiser le nouveau HMM avec celui du fit d'avant
            if self.current_detector is not None and self.current_detector.K == k:
                old_seed = self.current_detector.random_state
                det = self._clone_detector(k, old_seed)
                # Transfert du cerveau via BaseRegimeDetector
                det.set_params(self.current_detector.get_params()) #
                
                # Fit sans réinitialiser les paramètres (initialize=False) 
                det.fit(observations=[observations_train], 
                        inputs=[inputs_train] if self.input_vars else None, 
                        initialize=False, num_iters=50) # peu d'iters car on veut pas trop s'éloigner de l'ancien optima 
                
                if np.min(np.bincount(det.viterbi_states, minlength=k)) >= 2: # vérifie juste qu'on a au moins 2 observations par régime avant de faire quoi que ce soit
                    checker = self._create_checker(det, observations_train)
                    score, details = checker.compute_coherence_score(inputs=inputs_train)

                    # 4 piliers du checker
                    valid = (details["is_stable"] and details["sep_ratio"] >= self.sig_threshold and 
                             details["min_dur"] >= self.min_duration and details["avg_conf"] >= self.min_confidence)

                    # Stockage du meilleur score
                    if valid:
                        best_score = score
                        best_det = det
                        best_metadata = (k, "Warm-Start", old_seed) 

            # 2. Recherche standard (kmeans, random) : on teste d'initialiser le nouveau HMM avec des inits k-means et random 
            for method_name in ["kmeans", "random"]:
                trials = 5 if method_name == "kmeans" else 10
                for i in range(trials):
                    seed = seeds[i % len(seeds)]
                    det = self._clone_detector(k, seed)
                    det.fit(observations=[observations_train], 
                            inputs=[inputs_train] if self.input_vars else None, 
                            initialize=True, init_method=method_name, num_iters=200)
                    
                    if np.min(np.bincount(det.viterbi_states, minlength=k)) >= 2: # vérifie juste qu'on a au moins 2 observations par régime avant de faire quoi que ce soit
                        checker = self._create_checker(det, observations_train)
                        score, details = checker.compute_coherence_score(inputs=inputs_train)

                        # 4 piliers du checker
                        valid = (details["is_stable"] and details["sep_ratio"] >= self.sig_threshold and 
                                 details["min_dur"] >= self.min_duration and details["avg_conf"] >= self.min_confidence)

                        # Stockage du meilleur score
                        if valid and score > best_score:
                            best_score = score
                            best_det = det
                            best_metadata = (k, method_name, seed)

            if best_det is not None:
                return best_det, best_metadata[0], best_metadata[1], best_metadata[2]

        det = self._clone_detector(2, 120)
        det.fit(observations=[observations_train], inputs=[inputs_train] if self.input_vars else None)
        return det, 2, "Fallback", 120

    def run(self, start_date: Union[str, pd.Timestamp], end_date: Union[str, pd.Timestamp] = "2025-01-01"):
        """
        Boucle principale du Backtest

        Paramètres :
        - start_date : début du backtest
        - end_date : fin du backtest
        """

        test_indices = self.df.index[(self.df.index >= start_date) & (self.df.index <= end_date)] 
        pbar = tqdm(test_indices)
        
        for i, current_date in enumerate(pbar):
            train_mask = (self.df.index < current_date) & (self.df.index >= "1971-01-01") # Fenêtre d'entraînement
            observations_train = self.df.loc[train_mask, self.obs_vars].values.astype(float) * self.scaling 
            inputs_train = self.df.loc[train_mask, self.input_vars].values.astype(float) if self.input_vars else None
            
            # Refit périodique
            if i % self.refit_freq == 0 or self.current_detector is None:
                det, k, method, seed = self._find_best_model(observations_train, inputs_train)
                self.current_detector = det
                self.history_logs.append({'date': current_date, 'K': k, 'method': method, 'seed': seed})
                
                # Affichage à chaque refit
                pbar.write(f"\n >> [REFIT] {current_date.date()} | K={k} | Method={method}")
                sigmas_scaled = det.regime_covariances()
                for idx, s_scaled in enumerate(sigmas_scaled):
                    cov = s_scaled / (self.scaling**2)
                    vol_sp = np.sqrt(cov[0,0]) * np.sqrt(12)
                    vol_bond = np.sqrt(cov[1,1]) * np.sqrt(12)
                    corr = cov[0,1] / (np.sqrt(cov[0,0]) * np.sqrt(cov[1,1]) + 1e-16)
                    pbar.write(f"    Régime {idx}: Vol S&P={vol_sp:.1%}, Vol Bond={vol_bond:.1%}, Corr={corr:.2f}")

            # Prédiction OOS
            observations_full = self.df.loc[self.df.index <= current_date, self.obs_vars].values * self.scaling
            inputs_full = self.df.loc[self.df.index <= current_date, self.input_vars].values if self.input_vars else None
            
            # 1. On récupère les probas pour aujourd'hui sachant tout jusqu'à hier
            proba_now = self.current_detector.predict_probabilities(
                observations_full, inputs=inputs_full, oos_start=len(observations_full)-1
            )[0]
            
            # 2. On récupère et on corrige les matrices de cov de chaque régime 
            self.regime_probs[current_date] = proba_now 
            sigmas_scaled = self.current_detector.regime_covariances()
            sigmas_corrected = [s / (self.scaling**2) for s in sigmas_scaled]
            self.regime_sigmas[current_date] = sigmas_corrected 
            
            # 3. On calcule la covariance espérée
            self.results[current_date] = sum(proba_now[k] * sigmas_corrected[k] for k in range(self.current_detector.K))

    def get_risk_df(self):
        """
        Extrait les volatilités annualisées et corrélations de self.results
        """
        dates = sorted(self.results.keys())
        data = []
        for d in dates:
            cov = self.results[d]
            vol_sp = np.sqrt(cov[0,0]) * np.sqrt(12)
            vol_bond = np.sqrt(cov[1,1]) * np.sqrt(12)
            corr = cov[0,1] / (np.sqrt(cov[0,0]) * np.sqrt(cov[1,1]) + 1e-16)
            data.append([vol_sp, vol_bond, corr])
        return pd.DataFrame(data, index=dates, columns=['Vol_SP500', 'Vol_Tbond', 'Correlation'])
    
