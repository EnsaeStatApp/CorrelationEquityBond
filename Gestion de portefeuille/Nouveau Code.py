"""
portfolio_hmm_monthly.py (VERSION FINALE)

Usage: pip install pandas numpy scipy scikit-learn hmmlearn openpyxl

Place the CSV files in the folder defined in DATA_DIR.
The script:
- reads them,
- aligns on the common monthly date range,
- TRANSFORMS MACRO DATA (YoY changes) to ensure stationarity,
- CLEANS DATA (removes Infinities),
- computes monthly asset returns,
- runs a rolling HMM monthly and outputs monthly allocations for t+1.
- NO LEVERAGE: Strict 100% investment cap.
- DYNAMIC EXPORT: Handles any number of states (2, 3, 4...) for visualization.
"""

import os
import numpy as np
import pandas as pd
from sklearn.preprocessing import StandardScaler
from hmmlearn.hmm import GaussianHMM
from scipy.optimize import minimize
from datetime import datetime

# -----------------------------
# Paramètres
# -----------------------------
# ATTENTION : Modifie ce chemin selon ton dossier local
DATA_DIR = "C:/Users/skydr/Desktop/ENSAE/2eme Année/stat'app/Gestion portefeuille/"

ASSET_FILES = {
    'SP': DATA_DIR + 'SP500.csv',
    'AAA_yield': DATA_DIR + 'AAA.csv',
    'BAA_yield': DATA_DIR + 'BAA.csv',
    'T10_yield': DATA_DIR + '10Year_Treasury_Yield.csv'
}

MACRO_FILES = [
    DATA_DIR + 'AWHMAN.csv', DATA_DIR + 'BUSLOANS.csv', DATA_DIR + 'CUMFNS.csv', DATA_DIR + 'DGS10.csv', DATA_DIR + 'INDPRO.csv',
    DATA_DIR + 'M2SL.csv', DATA_DIR + 'PAYEMS.csv', DATA_DIR + 'PPIACO.csv',
    DATA_DIR + 'TB3MS.csv', DATA_DIR + 'TCU.csv', DATA_DIR + 'UMCSENT.csv', DATA_DIR + 'UNRATE.csv', DATA_DIR + 'USRECD.csv'
]

# HMM + optimisation
N_STATES = 4  # Tu peux changer ce chiffre (2, 3, 4), le code s'adaptera.
COV_TYPE = 'full'
RANDOM_STATE = 42
N_ITER = 500
ROLLING_FIT = True

# Bond durations (approx) used to convert yield -> price sensitivity
DURATIONS = {
    'T10': 8.5,
    'AAA': 8.0,
    'BAA': 7.0
}

ALLOW_SHORT = False
TARGET_VOL_ANN = 0.10 # Agit comme un PLAFOND de volatilité
MIN_TRAIN_PERIOD_MONTHS = 120

OUT_DIR = 'hmm_monthly_output'
os.makedirs(OUT_DIR, exist_ok=True)

# -----------------------------
# Fonctions utilitaires IO
# -----------------------------
def read_monthly_csv(path):
    """Lit un fichier CSV mensuel et renvoie DataFrame avec index Année-Mois et colonne Value"""
    if not os.path.exists(path):
        raise FileNotFoundError(f"Fichier non trouvé : {path}")

    # Déterminer séparateur et format de date
    if 'SP500' in os.path.basename(path):
        sep = ';'
        date_format = '%m/%d/%Y'
    else:
        sep = ','
        date_format = '%Y-%m-%d'

    df = pd.read_csv(path, sep=sep, engine='python', header=0)
    df.columns = [c.replace('"','').strip() for c in df.columns]

    if df.empty:
        raise ValueError(f"Le fichier {path} est vide après lecture.")

    # Détection colonne date
    date_col = next((c for c in df.columns if 'date' in c.lower()), df.columns[0])
    df[date_col] = pd.to_datetime(df[date_col], format=date_format, errors='coerce')
    df = df.dropna(subset=[date_col])
    df = df.set_index(date_col)

    # Conversion en Année-Mois (pour alignement propre)
    df.index = df.index.to_period('M')

    # Supprimer doublons éventuels
    df = df[~df.index.duplicated(keep='first')]

    # Détection colonne numérique
    num_col = next((c for c in df.columns if 'value' in c.lower() or pd.api.types.is_numeric_dtype(df[c])), df.columns[0])

    # Nettoyage valeurs (virgules, guillemets)
    if df[num_col].dtype == object:
        df[num_col] = df[num_col].astype(str).str.replace(',', '.', regex=False).str.replace('"','', regex=False).str.strip()
    
    df[num_col] = pd.to_numeric(df[num_col], errors='coerce')
    df = df.dropna(subset=[num_col])

    df = df[[num_col]]
    df.columns = ['Value']
    return df

def try_get_series_from_df(df):
    """Retourne une série numérique à partir du DataFrame"""
    if 'Value' in df.columns:
        s = df['Value']
        return pd.to_numeric(s, errors='coerce')
    
    # Fallback
    num_cols = df.select_dtypes(include=['float64','int64']).columns
    if len(num_cols) == 0:
        first_col = df.columns[1] if df.columns[0].lower()=='date' and len(df.columns)>1 else df.columns[0]
        s = pd.to_numeric(df[first_col], errors='coerce')
    else:
        s = df[num_cols[0]]
    return s

# -----------------------------
# Conversion yields -> bond returns (monthly)
# -----------------------------
def yields_to_monthly_bond_returns(yields_series, duration):
    ys = yields_series.copy().astype(float)
    # Heuristique : si yield > 1.5 (ex: 5.0), c'est probablement en %, on convertit en décimal
    if ys.median() > 1.5:
        ys = ys / 100.0
    carry = ys / 12.0
    dy = ys.diff()
    price_change = - duration * dy
    r = carry + price_change
    r.iloc[0] = carry.iloc[0]
    return r

# -----------------------------
# S&P500 total monthly return
# -----------------------------
def sp500_monthly_total_return(df_sp):
    s = try_get_series_from_df(df_sp).astype(float)
    s = s.sort_index()
    div_col = next((c for c in df_sp.columns if 'div' in c.lower()), None)
    if div_col is not None:
        div = df_sp[div_col].reindex(s.index).fillna(0).astype(float)
        ret = (s.diff() + div)/s.shift(1)
    else:
        ret = s.pct_change()
    return ret

# -----------------------------
# Align data on common monthly index
# -----------------------------
def build_common_dataframe(asset_files, macro_files):
    # 1. Lire les actifs
    asset_series = {k: read_monthly_csv(f) for k,f in asset_files.items()}
    
    # 2. Lire et TRANSFORMER les macros (FEATURE ENGINEERING)
    macro_series = {}
    for fname in macro_files:
        df = read_monthly_csv(fname)
        s = try_get_series_from_df(df)
        
        # --- MODIFICATION STATIONNARITÉ (YoY) ---
        name_upper = os.path.basename(fname).upper()
        
        # Liste heuristique des variables qui sont des Taux (Rate) ou binaires (USRECD)
        is_rate = any(x in name_upper for x in ['UNRATE', 'RATE', 'YIELD', 'DGS', 'TB3', 'AAA', 'BAA', 'FEDFUNDS', 'USRECD'])
        
        if is_rate:
            # Pour un taux ou binaire, diff absolue sur 12 mois
            s_transformed = s.diff(12)
        else:
            # Pour un niveau, variation en % sur 12 mois
            s_transformed = s.pct_change(12)
            
        macro_series[fname] = s_transformed

    # 3. Intersection des index
    all_indices = [v.index for v in list(asset_series.values()) + list(macro_series.values())]
    idx_intersection = all_indices[0]
    for idx in all_indices[1:]:
        idx_intersection = idx_intersection.intersection(idx)
    idx_intersection = idx_intersection.sort_values()

    # 4. Reindex final et Nettoyage INFINITY
    assets_aligned = {k: try_get_series_from_df(df).reindex(idx_intersection) for k, df in asset_series.items()}
    macros_aligned = {k: s.reindex(idx_intersection) for k, s in macro_series.items()}

    df_macros = pd.DataFrame(macros_aligned, index=idx_intersection)
    df_assets = pd.DataFrame(assets_aligned, index=idx_intersection)

    # --- SÉCURITÉ CRITIQUE : Remplacer l'infini par NaN ---
    df_macros = df_macros.replace([np.inf, -np.inf], np.nan)
    
    # On supprime les lignes qui contiennent maintenant des NaN
    df_macros = df_macros.dropna()
    
    # On réaligne les actifs sur les macros survivantes
    df_assets = df_assets.reindex(df_macros.index)

    return df_assets, df_macros

# -----------------------------
# HMM helpers
# -----------------------------
def fit_hmm(features_df, n_states=N_STATES, cov_type='full', random_state=RANDOM_STATE, n_iter=N_ITER):
    scaler = StandardScaler()
    X = scaler.fit_transform(features_df.values)
    model = GaussianHMM(n_components=n_states, covariance_type=cov_type, n_iter=n_iter, random_state=random_state)
    model.fit(X)
    return model, scaler

def get_posteriors_and_states(model, scaler, features_df):
    X = scaler.transform(features_df.values)
    try:
        posteriors = model.predict_proba(X)
    except Exception:
        logprob, posteriors = model.score_samples(X)
    states = model.predict(X)
    return posteriors, states

def estimate_state_moments_from_states(states, asset_returns_df):
    n_states = int(states.max()+1)
    state_means = {}
    state_covs = {}
    for k in range(n_states):
        mask = (states == k)
        subset = asset_returns_df.values[mask,:]
        if subset.shape[0] < 2:
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
    diffs = mus - mu_pred[None,:]
    Cov_of_means = np.zeros_like(E_Cov)
    for k in range(len(predictive_probs)):
        Cov_of_means += predictive_probs[k]*np.outer(diffs[k], diffs[k])
    Sigma_pred = E_Cov + Cov_of_means
    return mu_pred, Sigma_pred

def optimize_portfolio(mu, Sigma, target_vol_annual, allow_short=ALLOW_SHORT):
    n = len(mu)
    bounds = [(-1.0,1.0)]*n if allow_short else [(0.0,1.0)]*n
    target_var_month = target_vol_annual**2 / 12.0
    cons = [{'type':'eq','fun': lambda w: np.sum(w)-1.0},
            {'type':'ineq','fun': lambda w: target_var_month - float(w@Sigma@w)}]
    def obj(w): return -float(w@mu)
    x0 = np.array([1.0/n]*n)
    res = minimize(obj, x0, method='SLSQP', bounds=bounds, constraints=cons,
                   options={'ftol':1e-9,'maxiter':1000})
    if res.success:
        return {'weights': res.x, 'achieved_vol_annual': np.sqrt(res.x@Sigma@res.x*12), 'success': True, 'message': res.message}
    # fallback min-var
    def var_obj(w): return float(w@Sigma@w)
    res_mv = minimize(var_obj, x0, method='SLSQP', bounds=bounds, constraints=[{'type':'eq','fun': lambda w: np.sum(w)-1.0}])
    return {'weights': res_mv.x, 'achieved_vol_annual': np.sqrt(res_mv.x@Sigma@res_mv.x*12), 'success': True, 'message':'min-variance fallback'}

# -----------------------------
# Processus principal
# -----------------------------
def run_monthly_pipeline(asset_files, macro_files, target_vol_ann=TARGET_VOL_ANN,
                         n_states=N_STATES, durations=DURATIONS, rolling_fit=ROLLING_FIT):
    print("Lecture et alignement des fichiers...")
    assets_df_raw, macros_df_raw = build_common_dataframe(asset_files, macro_files)

    print("Calcul des rendements mensuels des actifs...")
    sp_ret = sp500_monthly_total_return(read_monthly_csv(asset_files['SP'])).reindex(assets_df_raw.index)
    aaa_s = try_get_series_from_df(read_monthly_csv(asset_files['AAA_yield'])).reindex(assets_df_raw.index)
    baa_s = try_get_series_from_df(read_monthly_csv(asset_files['BAA_yield'])).reindex(assets_df_raw.index)
    t10_s = try_get_series_from_df(read_monthly_csv(asset_files['T10_yield'])).reindex(assets_df_raw.index)
    r_aaa = yields_to_monthly_bond_returns(aaa_s, durations['AAA'])
    r_baa = yields_to_monthly_bond_returns(baa_s, durations['BAA'])
    r_t10 = yields_to_monthly_bond_returns(t10_s, durations['T10'])

    asset_returns = pd.DataFrame({
        'SP': sp_ret, 'AAA': r_aaa, 'BAA': r_baa, 'T10': r_t10
    }, index=assets_df_raw.index).sort_index()

    # Drop NaNs
    combined = pd.concat([asset_returns, macros_df_raw], axis=1).dropna(how='all')
    combined = combined.dropna(subset=['SP','AAA','BAA','T10'], how='any')
    combined = combined.dropna(how='any')

    features_all = macros_df_raw.reindex(combined.index)
    assets_returns_aligned = asset_returns.reindex(combined.index)

    print(f"Plage commune mensuelle (après transfo YoY): {features_all.index.min().to_timestamp().date()} -> {features_all.index.max().to_timestamp().date()}")
    print(f"Nombre de mois disponibles: {len(features_all)}")

    allocations, diagnostics, transmats = [], [], []
    idx = features_all.index
    start_pos = MIN_TRAIN_PERIOD_MONTHS if MIN_TRAIN_PERIOD_MONTHS < len(idx) else int(len(idx)*0.2)
    print(f"Début des allocations après {start_pos} mois de données (index {idx[start_pos].to_timestamp().date()})")

    for pos in range(start_pos, len(idx)-1):
        cutoff_date = idx[pos]
        train_features = features_all.loc[:cutoff_date].dropna(axis=1, how='all')
        train_assets = assets_returns_aligned.loc[:cutoff_date]

        if rolling_fit:
            model, scaler = fit_hmm(train_features, n_states=n_states)
        else:
            if pos==start_pos:
                model, scaler = fit_hmm(features_all, n_states=n_states)

        posteriors, states = get_posteriors_and_states(model, scaler, train_features)
        posterior_t = posteriors[-1,:]
        predictive_next = posterior_t @ model.transmat_
        state_means, state_covs = estimate_state_moments_from_states(states, train_assets)
        mu_pred, Sigma_pred = regime_weighted_moments(predictive_next, state_means, state_covs)
        Sigma_pred = Sigma_pred + 1e-8*np.eye(Sigma_pred.shape[0])
        
        # --- Optimisation "Long Only" (Pas de Levier) ---
        opt = optimize_portfolio(mu_pred, Sigma_pred, target_vol_ann, allow_short=ALLOW_SHORT)

        # On garde les poids bruts (somme = 1)
        final_weights = opt['weights']
        final_vol = opt['achieved_vol_annual']
        leverage = 1.0 # Toujours 1.0

        allocations.append({
            'date': cutoff_date, 
            'weights': final_weights, 
            'achieved_vol_annual': final_vol,
            'success': opt['success'], 
            'message': opt['message'], 
            'predictive_probs': predictive_next,
            'leverage_used': leverage
        })
        diagnostics.append({'date': cutoff_date, 'mu_pred': mu_pred, 'Sigma_trace': np.trace(Sigma_pred),
                            'predictive_probs': predictive_next, 'n_train_months': len(train_features)})
        transmats.append(model.transmat_)
        if (pos-start_pos)%24==0:
            print(f"Processed {pos-start_pos} months (cutoff {cutoff_date.to_timestamp().date()})")

    dates = [a['date'] for a in allocations]
    weight_matrix = np.vstack([a['weights'] for a in allocations])
    weights_df = pd.DataFrame(weight_matrix, index=dates, columns=['w_SP','w_AAA','w_BAA','w_T10'])
    
    # --- EXPORT DYNAMIQUE POUR LE PLOT (CORRECTION) ---
    # Cette boucle s'adapte automatiquement à N_STATES (p0, p1, p2, p3...)
    diag_records = []
    for d in allocations:
        rec = {
            'date': d['date'],
            'achieved_vol_annual': d['achieved_vol_annual'],
            'leverage_used': d['leverage_used'],
            'success': d['success'],
            'message': d['message']
        }
        for i, prob in enumerate(d['predictive_probs']):
            rec[f'p{i}'] = prob
        diag_records.append(rec)

    diag_df = pd.DataFrame(diag_records).set_index('date')
    mean_transmat = np.mean(np.array(transmats), axis=0)

    weights_df.to_csv(os.path.join(OUT_DIR,'monthly_allocations.csv'))
    diag_df.to_csv(os.path.join(OUT_DIR,'monthly_diagnostics.csv'))
    pd.DataFrame(mean_transmat).to_csv(os.path.join(OUT_DIR,'mean_transition_matrix.csv'))

    print("Terminé. Résultats sauvés dans:", OUT_DIR)
    return {'weights_df': weights_df,'diag_df': diag_df,'mean_transmat': mean_transmat}

# -----------------------------
# Main
# -----------------------------
if __name__ == '__main__':
    out = run_monthly_pipeline(ASSET_FILES, MACRO_FILES, target_vol_ann=TARGET_VOL_ANN,
                               n_states=N_STATES, durations=DURATIONS, rolling_fit=ROLLING_FIT)
    print("\nMatrice de transition (moyenne rolling):\n", out['mean_transmat'])
    print("\nExtrait allocations (dernier 5):\n", out['weights_df'].tail())
