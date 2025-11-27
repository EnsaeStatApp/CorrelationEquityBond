"""
portfolio_hmm_monthly.py

Usage: pip install pandas numpy scipy scikit-learn hmmlearn openpyxl

Place the CSV files in the same folder:
- S&P500.csv
- AAA.csv
- BAA.csv
- 10Year_Treasury_Yield.csv
- AWHMAN.csv, BUSLOANS.csv, CUMFNS.csv, DGS10.csv, INDPRO.csv, M2SL.csv,
  M0516BUSM163SNBR.csv, PAYEMS.csv, PPIACO.csv, TB3MS.csv, TCU.csv,
  UMCSENT.csv, UNRATE.csv, USRECD.csv

The script:
- reads them,
- aligns on the common monthly date range,
- computes monthly asset returns (S&P500 total return approx, bonds from yields),
- runs a rolling HMM monthly and outputs monthly allocations for t+1.
"""

import os
import numpy as np
import pandas as pd
from sklearn.preprocessing import StandardScaler
from hmmlearn.hmm import GaussianHMM
from scipy.optimize import minimize
from datetime import datetime

# -----------------------------
# Paramètres (modifie ici si besoin)
# -----------------------------
ASSET_FILES = {
    'SP': 'S&P500.csv',
    'AAA_yield': 'AAA.csv',
    'BAA_yield': 'BAA.csv',
    'T10_yield': '10Year_Treasury_Yield.csv'
}

MACRO_FILES = [
    'AWHMAN.csv', 'BUSLOANS.csv', 'CUMFNS.csv', 'DGS10.csv', 'INDPRO.csv',
    'M2SL.csv', 'M0516BUSM163SNBR.csv', 'PAYEMS.csv', 'PPIACO.csv',
    'TB3MS.csv', 'TCU.csv', 'UMCSENT.csv', 'UNRATE.csv', 'USRECD.csv'
]

# HMM + optimisation
N_STATES = 3
COV_TYPE = 'full'
RANDOM_STATE = 42
N_ITER = 200
ROLLING_FIT = True   # True = fit HMM on data up to t (causal); False = fit once on whole sample

# Bond durations (approx) used to convert yield -> price sensitivity
DURATIONS = {
    'T10': 7.0,   # approx duration for 10-year treasury (modifiable)
    'AAA': 7.0,
    'BAA': 7.0
}

ALLOW_SHORT = False
TARGET_VOL_ANN = 0.10   # example: 10% annual target (modifiable)
MIN_TRAIN_PERIOD_MONTHS = 120  # min months to start making allocations (e.g. 10 years)

OUT_DIR = 'hmm_monthly_output'
os.makedirs(OUT_DIR, exist_ok=True)

# -----------------------------
# Fonctions utilitaires IO
# -----------------------------
def read_monthly_csv(path):
    """
    Lit un CSV mensuel en essayant de détecter la colonne date.
    Retourne DataFrame index DatetimeIndex (month start).
    """
    df = pd.read_csv(path)
    # find a date-like column
    date_col = None
    for c in df.columns:
        if c.lower() in ('date', 'date.', 'datum', 'month', 'period', 'periodo', 'fecha'):
            date_col = c
            break
    if date_col is None:
        # try first column
        date_col = df.columns[0]
    # parse
    df[date_col] = pd.to_datetime(df[date_col], errors='coerce')
    df = df.dropna(subset=[date_col])
    df = df.set_index(date_col)
    # convert to month start (normalize)
    df.index = pd.to_datetime(df.index).to_period('M').to_timestamp()
    # if multiple numerical columns, keep numeric ones except index
    return df

def try_get_series_from_df(df):
    """Return a single series from df: if multiple numeric columns, pick the most plausible (Close, VALUE, etc.)"""
    numeric_cols = df.select_dtypes(include=[np.number]).columns.tolist()
    if not numeric_cols:
        raise ValueError("No numeric columns found in dataframe.")
    # prefer column names
    for prefer in ['Adj Close', 'AdjClose', 'Close', 'VALUE', 'value', 'PRICE', 'price', 'P', 'p']:
        if prefer in df.columns:
            return df[prefer]
    # else return first numeric
    return df[numeric_cols[0]]

# -----------------------------
# Conversion yields -> bond returns (monthly)
# -----------------------------
def yields_to_monthly_bond_returns(yields_series, duration):
    """
    yields_series: pandas Series of yields in percent or decimals
    duration: approximate duration in years
    returns: monthly returns in decimal (e.g. 0.002 for 0.2%)
    Method: r_t ≈ carry + price_change ≈ yield_t/12 - duration * (yield_t - yield_{t-1})
    Expects yields in decimals (0.03). If percentages (>1), divides by 100.
    """
    ys = yields_series.copy().astype(float)
    # fix percent vs decimal
    if ys.median() > 1.5:  # likely in % (e.g. 3.5)
        ys = ys / 100.0
    # carry
    carry = ys / 12.0
    # delta yield
    dy = ys.diff()
    price_change = - duration * dy
    r = carry + price_change
    r.iloc[0] = carry.iloc[0]  # first obs: only carry
    return r

# -----------------------------
# S&P500 total monthly return (approx)
# -----------------------------
def sp500_monthly_total_return(df_sp):
    """
    df_sp: DataFrame for SP file. Try to detect price column and dividend column if present.
    Returns monthly returns as decimal.
    """
    s = try_get_series_from_df(df_sp).astype(float)
    s = s.sort_index()
    # look for dividend column
    div_col = None
    for c in df_sp.columns:
        if 'div' in c.lower():
            div_col = c
            break
    if div_col is not None:
        div = df_sp[div_col].reindex(s.index).fillna(0).astype(float)
        # compute total return: (P_t - P_{t-1} + div_t) / P_{t-1}
        ret = (s.diff() + div) / s.shift(1)
    else:
        # simple price returns
        ret = s.pct_change()
    return ret

# -----------------------------
# Align data on common monthly index
# -----------------------------
def build_common_dataframe(asset_files, macro_files):
    # read asset series
    asset_series = {}
    for key, fname in asset_files.items():
        df = read_monthly_csv(fname)
        asset_series[key] = df

    # read macro
    macro_series = {}
    for fname in macro_files:
        df = read_monthly_csv(fname)
        # choose first numeric column
        try:
            macro_series[fname] = try_get_series_from_df(df)
        except Exception:
            macro_series[fname] = df.iloc[:, 0]

    # Build index intersection
    all_indices = []
    for v in list(asset_series.values()) + list(macro_series.values()):
        all_indices.append(v.index)
    idx_intersection = all_indices[0]
    for idx in all_indices[1:]:
        idx_intersection = idx_intersection.intersection(idx)
    idx_intersection = idx_intersection.sort_values()

    # Put everything into aligned DataFrames
    assets_aligned = {}
    for k, df in asset_series.items():
        # pick main series
        s = try_get_series_from_df(df)
        assets_aligned[k] = s.reindex(idx_intersection)

    macros_aligned = {}
    for k, s in macro_series.items():
        macros_aligned[k] = s.reindex(idx_intersection)

    assets_df = pd.DataFrame(assets_aligned, index=idx_intersection)
    macros_df = pd.DataFrame(macros_aligned, index=idx_intersection)

    return assets_df, macros_df

# -----------------------------
# HMM helpers & state moments (causal: up to time t)
# -----------------------------
def fit_hmm(features_df, n_states=N_STATES, cov_type=COV_TYPE, random_state=RANDOM_STATE, n_iter=N_ITER):
    scaler = StandardScaler()
    X = scaler.fit_transform(features_df.values)
    model = GaussianHMM(n_components=n_states, covariance_type=cov_type,
                        n_iter=n_iter, random_state=random_state)
    model.fit(X)
    return model, scaler

def get_posteriors_and_states(model, scaler, features_df):
    X = scaler.transform(features_df.values)
    # try predict_proba
    try:
        posteriors = model.predict_proba(X)
    except Exception:
        logprob, posteriors = model.score_samples(X)
    states = model.predict(X)
    return posteriors, states

def estimate_state_moments_from_states(states, asset_returns_df):
    """
    states: array length T (states for each observation in training)
    asset_returns_df: DataFrame aligned with same index as states, returns up to t
    returns dicts of means and covs for each state (k -> vector/matrix)
    """
    n_states = int(states.max() + 1)
    state_means = {}
    state_covs = {}
    for k in range(n_states):
        mask = (states == k)
        subset = asset_returns_df.values[mask, :]
        if subset.shape[0] < 2:
            # fallback to overall
            state_means[k] = np.nanmean(asset_returns_df.values, axis=0)
            state_covs[k] = np.cov(asset_returns_df.values.T, ddof=1)
        else:
            state_means[k] = np.mean(subset, axis=0)
            state_covs[k] = np.cov(subset.T, ddof=1)
    return state_means, state_covs

# -----------------------------
# Moments mixture & optimisation
# -----------------------------
def regime_weighted_moments(predictive_probs, state_means, state_covs):
    mus = np.array([state_means[k] for k in sorted(state_means.keys())])
    covs = np.array([state_covs[k] for k in sorted(state_covs.keys())])
    mu_pred = predictive_probs @ mus
    E_Cov = np.tensordot(predictive_probs, covs, axes=(0,0))
    diffs = mus - mu_pred[None, :]
    Cov_of_means = np.zeros_like(E_Cov)
    for k in range(len(predictive_probs)):
        Cov_of_means += predictive_probs[k] * np.outer(diffs[k], diffs[k])
    Sigma_pred = E_Cov + Cov_of_means
    return mu_pred, Sigma_pred

def optimize_portfolio(mu, Sigma, target_vol_annual, allow_short=ALLOW_SHORT):
    """
    Optimize to maximize expected return with variance <= target_vol^2 (annual), no shorts by default.
    Work on monthly variance (target_vol_annual^2 / 12).
    Returns weights or fallback min-var weights.
    """
    n = len(mu)
    if not allow_short:
        bounds = [(0.0, 1.0)] * n
    else:
        bounds = [(-1.0, 1.0)] * n
    target_var_month = (target_vol_annual ** 2) / 12.0

    cons = [{'type': 'eq', 'fun': lambda w: np.sum(w) - 1.0},
            {'type': 'ineq', 'fun': lambda w: target_var_month - float(w @ Sigma @ w)}]

    def obj(w):
        return -float(w @ mu)

    x0 = np.array([1.0 / n] * n)
    res = minimize(obj, x0, method='SLSQP', bounds=bounds, constraints=cons,
                   options={'ftol':1e-9, 'maxiter':1000})
    if res.success:
        w = res.x
        ach_vol_annual = np.sqrt(w @ Sigma @ w * 12.0)
        return {'weights': w, 'achieved_vol_annual': ach_vol_annual, 'success': True, 'message': res.message}
    # fallback to min variance
    def var_obj(w): return float(w @ Sigma @ w)
    cons_mv = [{'type': 'eq', 'fun': lambda w: np.sum(w) - 1.0}]
    res_mv = minimize(var_obj, x0, method='SLSQP', bounds=bounds, constraints=cons_mv,
                      options={'ftol':1e-9, 'maxiter':1000})
    if not res_mv.success:
        return {'weights': None, 'success': False, 'message': 'optimisation failed'}
    w = res_mv.x
    ach_vol_annual = np.sqrt(w @ Sigma @ w * 12.0)
    return {'weights': w, 'achieved_vol_annual': ach_vol_annual, 'success': True, 'message': 'min-variance fallback'}

# -----------------------------
# Processus principal
# -----------------------------
def run_monthly_pipeline(asset_files, macro_files, target_vol_ann=TARGET_VOL_ANN,
                         n_states=N_STATES, durations=DURATIONS, rolling_fit=ROLLING_FIT):
    # 1) Read and align raw data
    print("Lecture et alignement des fichiers...")
    assets_df_raw, macros_df_raw = build_common_dataframe(asset_files, macro_files)

    # 2) Compute monthly returns for assets
    print("Calcul des rendements mensuels des actifs...")
    # S&P
    sp_df = read_monthly_csv(asset_files['SP'])
    sp_ret = sp500_monthly_total_return(sp_df).reindex(assets_df_raw.index)

    # Yields -> bond returns
    aaa_df = read_monthly_csv(asset_files['AAA_yield'])
    baa_df = read_monthly_csv(asset_files['BAA_yield'])
    t10_df = read_monthly_csv(asset_files['T10_yield'])

    aaa_s = try_get_series_from_df(aaa_df).reindex(assets_df_raw.index)
    baa_s = try_get_series_from_df(baa_df).reindex(assets_df_raw.index)
    t10_s = try_get_series_from_df(t10_df).reindex(assets_df_raw.index)

    r_aaa = yields_to_monthly_bond_returns(aaa_s, durations['AAA'])
    r_baa = yields_to_monthly_bond_returns(baa_s, durations['BAA'])
    r_t10 = yields_to_monthly_bond_returns(t10_s, durations['T10'])

    asset_returns = pd.DataFrame({
        'SP': sp_ret,
        'AAA': r_aaa,
        'BAA': r_baa,
        'T10': r_t10
    }, index=assets_df_raw.index).sort_index()

    # Drop leading NaNs
    combined = pd.concat([asset_returns, macros_df_raw], axis=1)
    combined = combined.dropna(how='all')  # keep rows where at least one value is present
    # Now restrict to rows where we have at least asset returns present (for modelling & estimation)
    combined = combined.dropna(subset=['SP', 'AAA', 'BAA', 'T10'], how='any')

    # final aligned series & features
    features_all = macros_df_raw.reindex(combined.index)
    assets_returns_aligned = asset_returns.reindex(combined.index)

    print(f"Plage commune mensuelle: {features_all.index.min().date()} -> {features_all.index.max().date()}")
    print(f"Nombre de mois disponibles: {len(features_all)}")

    # Storage
    allocations = []
    diagnostics = []
    transmats = []

    # We will start after having at least MIN_TRAIN_PERIOD_MONTHS months
    idx = features_all.index
    start_pos = MIN_TRAIN_PERIOD_MONTHS if MIN_TRAIN_PERIOD_MONTHS < len(idx) else int(len(idx)*0.2)
    print(f"Début des allocations après {start_pos} mois de données (index {idx[start_pos].date()})")

    # Rolling loop: for each t from start_pos to len(idx)-2 we will create allocation to apply for t+1
    for pos in range(start_pos, len(idx)-1):
        cutoff_date = idx[pos]
        # training data up to and including cutoff_date
        train_features = features_all.loc[:cutoff_date].dropna(axis=1, how='all')  # keep
        train_assets = assets_returns_aligned.loc[:cutoff_date]

        # Fit HMM on train_features
        if rolling_fit:
            model, scaler = fit_hmm(train_features, n_states=n_states)
        else:
            if pos == start_pos:
                model, scaler = fit_hmm(features_all, n_states=n_states)

        # compute posteriors & states on training set
        posteriors, states = get_posteriors_and_states(model, scaler, train_features)
        # posterior at last obs (cutoff)
        posterior_t = posteriors[-1, :]

        # predictive next-step probs
        predictive_next = posterior_t @ model.transmat_

        # estimate state moments from states & returns in training period (causal)
        state_means, state_covs = estimate_state_moments_from_states(states, train_assets)

        # combine moments
        mu_pred, Sigma_pred = regime_weighted_moments(predictive_next, state_means, state_covs)
        # ensure PSD
        Sigma_pred = Sigma_pred + 1e-8 * np.eye(Sigma_pred.shape[0])

        # optimize portfolio given monthly Sigma (we handle conversion inside)
        opt = optimize_portfolio(mu_pred, Sigma_pred, target_vol_ann, allow_short=ALLOW_SHORT)

        # save allocation for cutoff_date (apply in next month)
        alloc = {
            'date': cutoff_date,
            'weights': opt['weights'],
            'achieved_vol_annual': opt.get('achieved_vol_annual'),
            'success': opt['success'],
            'message': opt['message'],
            'predictive_probs': predictive_next
        }
        allocations.append(alloc)
        diagnostics.append({
            'date': cutoff_date,
            'mu_pred': mu_pred,
            'Sigma_trace': np.trace(Sigma_pred),
            'predictive_probs': predictive_next,
            'n_train_months': len(train_features)
        })
        transmats.append(model.transmat_)

        # optional quick progress
        if (pos - start_pos) % 12 == 0:
            print(f"Processed {pos - start_pos} months (cutoff {cutoff_date.date()})")

    # Aggregate results to DataFrames
    dates = [a['date'] for a in allocations]
    weight_matrix = np.vstack([a['weights'] for a in allocations])
    weights_df = pd.DataFrame(weight_matrix, index=dates, columns=['w_SP', 'w_AAA', 'w_BAA', 'w_T10'])
    diag_df = pd.DataFrame([{
        'date': d['date'],
        'achieved_vol_annual': d.get('achieved_vol_annual'),
        'success': d['success'],
        'message': d['message'],
        'p0': d['predictive_probs'][0],
        'p1': d['predictive_probs'][1],
        'p2': d['predictive_probs'][2] if len(d['predictive_probs'])>2 else np.nan
    } for d in allocations]).set_index('date')

    # Average transition matrix over rolling fits
    mean_transmat = np.mean(np.array(transmats), axis=0)

    # save
    weights_df.to_csv(os.path.join(OUT_DIR, 'monthly_allocations.csv'))
    diag_df.to_csv(os.path.join(OUT_DIR, 'monthly_diagnostics.csv'))
    pd.DataFrame(mean_transmat).to_csv(os.path.join(OUT_DIR, 'mean_transition_matrix.csv'))

    print("Terminé. Résultats sauvés dans:", OUT_DIR)
    return {
        'weights_df': weights_df,
        'diag_df': diag_df,
        'mean_transmat': mean_transmat
    }

# -----------------------------
# Si lancé en tant que script
# -----------------------------
if __name__ == '__main__':
    out = run_monthly_pipeline(ASSET_FILES, MACRO_FILES, target_vol_ann=TARGET_VOL_ANN,
                               n_states=N_STATES, durations=DURATIONS, rolling_fit=ROLLING_FIT)

    print("\nMatrice de transition (moyenne rolling):\n", out['mean_transmat'])
    print("\nExtrait allocations (dernier 5):\n", out['weights_df'].tail())

