import numpy as np
from typing import List

def calculate_net_ret(weights: np.ndarray, prev_weights: np.ndarray, asset_returns_log: np.ndarray,
                      rf_monthly: float, borrow_spread_annual: float, cost_per_trade: float) -> float:
    """
    Calculates net monthly performance accounting for costs and leverage.
    """
    arith_rets = np.exp(asset_returns_log) - 1
    expo_risky = np.sum(weights)
    cash_weight = max(1.0 - expo_risky, 0.0)

    gross_arith_risky = np.dot(weights, arith_rets)
    cash_ret_arith = cash_weight * rf_monthly

    turnover = np.sum(np.abs(weights - prev_weights))
    trading_cost = turnover * cost_per_trade

    borrowed = max(expo_risky - 1.0, 0.0)
    financing_cost = borrowed * (rf_monthly + borrow_spread_annual / 12.0)

    net_arith_ret = gross_arith_risky + cash_ret_arith - trading_cost - financing_cost
    return np.log(max(1 + net_arith_ret, 1e-10))

def get_metrics(r: List[float], rf: List[float]) -> List[float]:
    """
    Computes key performance indicators (Annual Return, Vol, Sharpe, MDD, Calmar).
    """
    r, rf = np.array(r), np.array(rf)
    wealth = np.exp(np.cumsum(r))
    
    ann_ret = wealth[-1]**(12/len(r)) - 1
    arith = np.exp(r) - 1
    ann_vol = np.std(arith) * np.sqrt(12)
    sharpe = (np.mean(arith - rf) * 12) / ann_vol if ann_vol > 0 else 0
    
    mdd = (wealth / np.maximum.accumulate(wealth) - 1).min()
    calmar = ann_ret / abs(mdd) if mdd != 0 else 0
    
    return [ann_ret, ann_vol, sharpe, mdd, calmar]
