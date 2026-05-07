import pandas as pd
import numpy as np
from tqdm import tqdm
from typing import List, Dict
from src.stock_bond_correlation.strategy.allocator.RiskBudgetingAllocator import RiskBudgetingAllocator
from src.stock_bond_correlation.strategy.utils.performance import calculate_net_ret

class StrategySimulator:
    """
    Engine to simulate a regime-aware investment strategy based on HMM outputs.
    """
    
    def __init__(self, target_vol: float = 0.07, max_leverage: float = 1.0, 
                 lev_normal: float = 1.0, lev_min_stress: float = 0.10,
                 cost_per_trade: float = 0.0010, borrow_spread: float = 0.01,
                 smooth_alpha: float = 1.0):
        """
        Initializes the simulator with risk and cost parameters.
        """
        self.target_vol = target_vol
        self.max_leverage = max_leverage
        self.lev_normal = lev_normal
        self.lev_min_stress = lev_min_stress
        self.cost_per_trade = cost_per_trade
        self.borrow_spread = borrow_spread
        self.smooth_alpha = smooth_alpha

    def run(self, backtester, df: pd.DataFrame, asset_vars: List[str], 
            start_year: int = 1995, w_ref_6040: np.ndarray = np.array([0.6, 0.4]),
            sp_budget_range: tuple = (0.10, 0.95)) -> Dict:
        """
        Runs the strategy simulation over the backtester's dates.
        """
        all_dates = sorted(backtester.regime_probs.keys())
        dates = [d for d in all_dates if d in df.index and d.year >= start_year]
        
        # Result containers
        portfolio_rets, rets_hmm_pure, rets_6040, rets_rp = [], [], [], []
        history_rf, history_weights = [], []
        
        p_smooth = None
        n_assets = len(asset_vars)
        w_prev = np.zeros(n_assets)
        w_prev_pure = np.zeros(n_assets)
        w_prev_rp = np.zeros(n_assets)
        w_prev_6040 = w_ref_6040.copy()
        
        sp_min, sp_max = sp_budget_range

        for d in dates:
            r_t = df.loc[d, asset_vars].values
            rf_raw = df.loc[d, "TB3MS"]
            rf_monthly = (rf_raw / 100.0) / 12.0
            history_rf.append(rf_monthly)

            # 1. HMM Data Extraction
            p_raw = backtester.regime_probs[d]
            sigmas_k = backtester.regime_sigmas[d]
            p_smooth = p_raw if p_smooth is None else (self.smooth_alpha * p_raw + (1 - self.smooth_alpha) * p_smooth)
            n_regimes = len(p_raw)

            # 2. Stress Regime Detection (Logic based on 60/40 volatility)
            vols_6040_k = [np.sqrt(w_ref_6040.T @ (sigmas_k[k]*12) @ w_ref_6040) for k in range(n_regimes)]
            v_min, v_max = min(vols_6040_k), max(vols_6040_k)

            k_stress, max_danger = -1, -1e9
            for k in range(n_regimes):
                rho = sigmas_k[k][0,1] / (np.sqrt(sigmas_k[k][0,0]*sigmas_k[k][1,1]) + 1e-15)
                danger = (vols_6040_k[k] - v_min) / (v_max - v_min + 1e-9)
                if rho > 0 and danger > max_danger:
                    max_danger, k_stress = danger, k

            # 3. Allocation Building
            w_final, w_final_pure = np.zeros(n_assets), np.zeros(n_assets)
            
            for k in range(n_regimes):
                sig_k_ann = sigmas_k[k] * 12
                danger = (vols_6040_k[k] - v_min) / (v_max - v_min + 1e-9)
                
                # Dynamic Risk Budgeting
                b_sp = sp_max - danger * (sp_max - sp_min)
                w_rb = RiskBudgetingAllocator.get_risk_budget_weights(sigmas_k[k], np.array([b_sp, 1 - b_sp]))

                # Leverage Calculation
                vol_exp = np.sqrt(w_rb.T @ sig_k_ann @ w_rb)
                lev = min(self.target_vol / (vol_exp + 1e-9), self.max_leverage)
                
                # Stress Regime cap
                if k == k_stress: 
                    lev = min(lev, self.lev_normal - danger * (self.lev_normal - self.lev_min_stress))

                w_final += p_smooth[k] * (w_rb * lev)
                w_final_pure += p_smooth[k] * w_rb

            # 4. Benchmarks (Rolling Risk Parity)
            hist_mask = df.index < d
            if len(df.loc[hist_mask]) >= 12:
                cov_roll = df.loc[hist_mask, asset_vars].tail(12).cov().values * 12
                w_rp = RiskBudgetingAllocator.get_risk_budget_weights(cov_roll, np.array([0.5, 0.5]))
            else:
                w_rp = w_ref_6040.copy()

            # 5. Net Performance Calculation
            portfolio_rets.append(calculate_net_ret(w_final, w_prev, r_t, rf_monthly, self.borrow_spread, self.cost_per_trade))
            rets_hmm_pure.append(calculate_net_ret(w_final_pure, w_prev_pure, r_t, rf_monthly, 0.0, self.cost_per_trade))
            rets_6040.append(calculate_net_ret(w_ref_6040, w_prev_6040, r_t, rf_monthly, 0.0, self.cost_per_trade))
            rets_rp.append(calculate_net_ret(w_rp, w_prev_rp, r_t, rf_monthly, 0.0, self.cost_per_trade))

            # Store history
            expo = np.sum(w_final)
            history_weights.append([w_final[0], w_final[1], max(1.0 - expo, 0.0)])
            w_prev, w_prev_pure, w_prev_rp = w_final.copy(), w_final_pure.copy(), w_rp.copy()

        return {
            "dates": dates,
            "portfolio_rets": portfolio_rets,
            "rets_hmm_pure": rets_hmm_pure,
            "rets_6040": rets_6040,
            "rets_rp": rets_rp,
            "history_rf": history_rf,
            "history_weights": history_weights
        }