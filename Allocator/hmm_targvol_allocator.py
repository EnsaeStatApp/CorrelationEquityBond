"""
portfolio_ssm_monthly_OOP.py

Refactorisation Orientée Objet avec classe Allocator.
"""

import os
import sys
import numpy as np
import pandas as pd
from sklearn.preprocessing import StandardScaler
from scipy.optimize import minimize
from datetime import datetime
from abc import ABC, abstractmethod
from tqdm.auto import tqdm 

# Import des données via Google Drive (si utilisé sur Colab)
try:
    from google.colab import drive
    drive.mount('/content/drive')
except:
    pass

sys.path.append(os.getcwd())

# --- IMPORT DU MODÈLE ---
try:
    from hmm_detector import HMMDetector
    print("Succès : HMMDetector importé.")
except ImportError:
    print("Mode local/Demo : HMMDetector simulé.")
    class HMMDetector:
        def __init__(self, **kwargs): pass
        def fit(self, X): pass
        def regime_probabilities(self, Y): return np.random.dirichlet(np.ones(4), size=len(Y))
        def get_transition_matrix(self): return np.eye(4)

# ==============================================================================
# 1. DÉFINITION DES CLASSES D'ALLOCATION
# ==============================================================================

class MeanVarianceAllocator(Allocator):
    """
    Implémentation concrète utilisant l'optimisation Moyenne-Variance sous contrainte de Volatilité.
    Gère également les fallbacks (Min-Variance, Équipondéré).
    """
    def __init__(self, target_vol_annual=0.10, allow_short=False):
        self.target_vol_annual = target_vol_annual
        self.allow_short = allow_short
        self._mu = None # Stockage temporaire pour les rendements espérés

    def set_mu(self, mu: np.ndarray):
        """Définit le vecteur des rendements espérés pour la prochaine allocation."""
        self._mu = mu

    def allocate(self, Sigma: np.ndarray) -> np.ndarray:
        n = Sigma.shape[0]
        
        # Gestion du cas où mu n'est pas défini (Fallback sur MinVar implicite ou erreur)
        if self._mu is None:
            # Si pas de mu, on considère mu = 0 (Pure Min Variance)
            mu = np.zeros(n)
        else:
            mu = self._mu

        bounds = [(-1.0, 1.0)] * n if self.allow_short else [(0.0, 1.0)] * n
        target_var_month = self.target_vol_annual**2 / 12.0
        x0 = np.array([1.0/n]*n)

        # Contraintes : Somme(w) = 1  ET  w'Sw <= TargetVar
        cons = [
            {'type': 'eq', 'fun': lambda w: np.sum(w) - 1.0},
            {'type': 'ineq', 'fun': lambda w: target_var_month - float(w @ Sigma @ w)}
        ]

        # Fonction objectif : Maximiser w'mu => Minimiser -w'mu
        def obj_return(w): return -float(w @ mu)

        # 1. Tentative Max Return s/c Target Vol
        try:
            res = minimize(obj_return, x0, method='SLSQP', bounds=bounds, constraints=cons,
                           options={'ftol': 1e-9, 'maxiter': 1000})
            if res.success:
                return res.x
        except Exception:
            pass

        # 2. Fallback: Minimum Variance (si la cible de vol est trop agressive ou échec)
        # On minimise w'Sw s/c Somme(w)=1
        def obj_var(w): return float(w @ Sigma @ w)
        try:
            res_mv = minimize(obj_var, x0, method='SLSQP', bounds=bounds, 
                              constraints=[{'type': 'eq', 'fun': lambda w: np.sum(w) - 1.0}])
            if res_mv.success:
                return res_mv.x
        except Exception:
            pass

        # 3. Fallback ultime : Équipondéré
        return x0

# ==============================================================================
# 2. PARAMÈTRES ET CONFIGURATION
# ==============================================================================

DATA_DIR = "/content/drive/MyDrive/Données Statapp/"
# DATA_DIR = "./" # Décommenter pour test local

ASSET_FILES = {
    'SP': os.path.join(DATA_DIR, 'SP500.csv'),
    'AAA_yield': os.path.join(DATA_DIR, 'AAA.csv'),
    'BAA_yield': os.path.join(DATA_DIR, 'BAA.csv'),
    'T10_yield': os.path.join(DATA_DIR, '10Year_Treasury_Yield.csv')
}

MACRO_FILES = [
    os.path.join(DATA_DIR, f) for f in [
        'AWHMAN.csv', 'BUSLOANS.csv', 'CUMFNS.csv', 'DGS10.csv', 'INDPRO.csv',
        'M2SL.csv', 'PAYEMS.csv', 'PPIACO.csv', 'TB3MS.csv', 'TCU.csv',
        'UMCSENT.csv', 'UNRATE.csv', 'USRECD.csv'
    ]
]

N_STATES = 4
RANDOM_STATE = 42
N_ITER = 100
ROLLING_FIT = True

DURATIONS = {'T10': 8.5, 'AAA': 8.0, 'BAA': 7.0}
ALLOW_SHORT = False
TARGET_VOL_ANN = 0.10
MIN_TRAIN_PERIOD_MONTHS = 120

OUT_DIR = 'hmm_monthly_output'
os.makedirs(OUT_DIR, exist_ok=True)

# ==============================================================================
# 3. FONCTIONS UTILITAIRES (IO & MATHS)
# ==============================================================================

def read_monthly_csv(path):
    if not os.path.exists(path):
        return pd.DataFrame()
    
    sep = ';' if 'SP500' in os.path.basename(path) else ','
    date_format = '%m/%d/%Y' if 'SP500' in os.path.basename(path) else '%Y-%m-%d'

    try:
        df = pd.read_csv(path, sep=sep, engine='python', header=0)
    except:
        return pd.DataFrame()

    df.columns = [c.replace('"','').strip() for c in df.columns]
    if df.empty: return df

    date_col = next((c for c in df.columns if 'date' in c.lower()), df.columns[0])
    df[date_col] = pd.to_datetime(df[date_col], format=date_format, errors='coerce')
    df = df.dropna(subset=[date_col]).set_index(date_col)
    df.index = df.index.to_period('M')
    df = df[~df.index.duplicated(keep='first')]

    num_col = next((c for c in df.columns if 'value' in c.lower() or pd.api.types.is_numeric_dtype(df[c])), df.columns[0])
    if df[num_col].dtype == object:
        df[num_col] = df[num_col].astype(str).str.replace(',', '.', regex=False).str.replace('"','', regex=False).str.strip()

    df[num_col] = pd.to_numeric(df[num_col], errors='coerce')
    return df[[num_col]].rename(columns={num_col: 'Value'}).dropna()

def try_get_series_from_df(df):
    if df.empty: return pd.Series(dtype=float)
    return pd.to_numeric(df['Value'], errors='coerce')

def yields_to_monthly_bond_returns(yields_series, duration):
    ys = yields_series.copy().astype(float)
    if ys.median() > 1.5: ys = ys / 100.0
    carry = ys / 12.0
    price_change = - duration * ys.diff()
    r = carry + price_change
    r.iloc[0] = carry.iloc[0]
    return r

def sp500_monthly_total_return(df_sp):
    s = try_get_series_from_df(df_sp).astype(float).sort_index()
    if s.empty: return s
    div_col = next((c for c in df_sp.columns if 'div' in c.lower()), None)
    if div_col:
        div = df_sp[div_col].reindex(s.index).fillna(0).astype(float)
        return (s.diff() + div)/s.shift(1)
    return s.pct_change()

def build_common_dataframe(asset_files, macro_files):
    asset_series = {}
    for k, f in asset_files.items():
        df = read_monthly_csv(f)
        if not df.empty: asset_series[k] = df

    macro_series = {}
    for fname in macro_files:
        df = read_monthly_csv(fname)
        if df.empty: continue
        s = try_get_series_from_df(df)
        name_upper = os.path.basename(fname).upper()
        is_rate = any(x in name_upper for x in ['UNRATE', 'RATE', 'YIELD', 'DGS', 'TB3', 'AAA', 'BAA'])
        macro_series[fname] = s.diff(12) if is_rate else s.pct_change(12)

    if not asset_series: return pd.DataFrame(), pd.DataFrame() # Fail gracefully

    all_indices = [v.index for v in list(asset_series.values()) + list(macro_series.values())]
    idx_intersection = all_indices[0]
    for idx in all_indices[1:]: idx_intersection = idx_intersection.intersection(idx)
    
    idx_intersection = idx_intersection.sort_values()
    assets_aligned = {k: try_get_series_from_df(df).reindex(idx_intersection) for k, df in asset_series.items()}
    macros_aligned = {k: s.reindex(idx_intersection) for k, s in macro_series.items()}

    df_macros = pd.DataFrame(macros_aligned, index=idx_intersection).replace([np.inf, -np.inf], np.nan).dropna()
    df_assets = pd.DataFrame(assets_aligned, index=idx_intersection).reindex(df_macros.index)
    
    return df_assets, df_macros

def estimate_asset_moments_from_probs(probs_vector, asset_returns):
    weights = np.maximum(probs_vector / (np.sum(probs_vector) + 1e-12), 0)
    mean_w = np.average(asset_returns, axis=0, weights=weights)
    diff = asset_returns - mean_w
    cov_w = np.cov(diff.T, aweights=weights, ddof=0)
    return mean_w, cov_w

def regime_weighted_moments(predictive_probs, state_means, state_covs):
    mus = np.array([state_means[k] for k in sorted(state_means.keys())])
    covs = np.array([state_covs[k] for k in sorted(state_covs.keys())])
    mu_pred = predictive_probs @ mus
    E_Cov = np.tensordot(predictive_probs, covs, axes=(0,0))
    diffs = mus - mu_pred[None,:]
    Cov_of_means = np.zeros_like(E_Cov)
    for k in range(len(predictive_probs)):
        Cov_of_means += predictive_probs[k]*np.outer(diffs[k], diffs[k])
    return mu_pred, E_Cov + Cov_of_means

# ==============================================================================
# 4. PIPELINE PRINCIPAL
# ==============================================================================

def run_ssm_pipeline(asset_files, macro_files, target_vol_ann=TARGET_VOL_ANN,
                     n_states=N_STATES, durations=DURATIONS, rolling_fit=ROLLING_FIT):

    print("--- Démarrage Pipeline HMM (OOP Refactor) ---")
    assets_df_raw, macros_df_raw = build_common_dataframe(asset_files, macro_files)
    if assets_df_raw.empty:
        print("Erreur: Données insuffisantes.")
        return

    # Calcul Rendements
    sp_df = read_monthly_csv(asset_files['SP'])
    sp_ret = sp500_monthly_total_return(sp_df).reindex(assets_df_raw.index)
    
    r_aaa = yields_to_monthly_bond_returns(try_get_series_from_df(read_monthly_csv(asset_files['AAA_yield'])), durations['AAA'])
    r_baa = yields_to_monthly_bond_returns(try_get_series_from_df(read_monthly_csv(asset_files['BAA_yield'])), durations['BAA'])
    r_t10 = yields_to_monthly_bond_returns(try_get_series_from_df(read_monthly_csv(asset_files['T10_yield'])), durations['T10'])

    asset_returns = pd.DataFrame({'SP': sp_ret, 'AAA': r_aaa, 'BAA': r_baa, 'T10': r_t10}, index=assets_df_raw.index).sort_index()
    combined = pd.concat([asset_returns, macros_df_raw], axis=1).replace([np.inf, -np.inf], np.nan).dropna()
    
    features_all = macros_df_raw.reindex(combined.index)
    assets_returns_aligned = asset_returns.reindex(combined.index)

    # INSTANCIATION DE L'ALLOCATEUR
    allocator = MeanVarianceAllocator(target_vol_annual=target_vol_ann, allow_short=ALLOW_SHORT)

    allocations = []
    transmats = []
    idx = features_all.index
    start_pos = MIN_TRAIN_PERIOD_MONTHS if MIN_TRAIN_PERIOD_MONTHS < len(idx) else int(len(idx)*0.2)
    model_cache = None

    iterator = tqdm(range(start_pos, len(idx)-1), desc="Calcul Allocations", unit="mois")

    for pos in iterator:
        cutoff_date = idx[pos]
        train_features = features_all.loc[:cutoff_date]
        train_assets = assets_returns_aligned.loc[:cutoff_date]

        if train_features.shape[0] < 50: continue

        X_train_scaled = StandardScaler().fit_transform(train_features.values)
        if np.isnan(X_train_scaled).any(): X_train_scaled = np.nan_to_num(X_train_scaled)

        try:
            # HMM Fitting
            if rolling_fit or model_cache is None:
                detector = HMMDetector(n_states=n_states, n_iter=N_ITER, random_state=RANDOM_STATE)
                detector.fit(X_train_scaled)
                model_cache = detector
            else:
                detector = model_cache

            probs_all = detector.regime_probabilities(Y=X_train_scaled)
            posterior_t = probs_all[-1, :]
            trans_mat = detector.get_transition_matrix()
            if np.isnan(trans_mat).any(): trans_mat = np.eye(n_states)/n_states
            transmats.append(trans_mat)

            predictive_next = posterior_t @ trans_mat

            # Moments
            state_means, state_covs = {}, {}
            for k in range(n_states):
                mu_k, sigma_k = estimate_asset_moments_from_probs(probs_all[:, k], train_assets.values)
                state_means[k] = mu_k
                state_covs[k] = sigma_k

            mu_pred, Sigma_pred = regime_weighted_moments(predictive_next, state_means, state_covs)
            Sigma_pred += 1e-8 * np.eye(Sigma_pred.shape[0])

            # --- UTILISATION DE L'ALLOCATEUR ---
            # 1. On injecte les rendements espérés (état interne)
            allocator.set_mu(mu_pred)
            # 2. On appelle la méthode du contrat (allocate) avec la matrice Sigma
            weights = allocator.allocate(Sigma_pred)

            # 3. Calcul manuel des métriques ex-post (puisque allocate ne renvoie que les poids)
            achieved_vol = np.sqrt(weights @ Sigma_pred @ weights * 12)
            
            allocations.append({
                'date': cutoff_date,
                'weights': weights,
                'achieved_vol_annual': achieved_vol,
                'predictive_probs': predictive_next,
                'success': True # On suppose le succès si des poids sont renvoyés
            })

            iterator.set_postfix({'Date': str(cutoff_date), 'Vol': f"{achieved_vol:.1%}"})

        except Exception as e:
            continue

    if not allocations:
        print("Erreur: Aucune allocation générée.")
        return {}

    # Export des résultats
    weight_matrix = np.vstack([a['weights'] for a in allocations])
    weights_df = pd.DataFrame(weight_matrix, index=[a['date'] for a in allocations], columns=['SP','AAA','BAA','T10'])

    diag_records = []
    for d in allocations:
        rec = {'date': d['date'], 'achieved_vol_annual': d['achieved_vol_annual'], 'success': d['success']}
        for i, prob in enumerate(d['predictive_probs']): rec[f'prob_regime_{i}'] = prob
        diag_records.append(rec)
    diag_df = pd.DataFrame(diag_records).set_index('date')

    if transmats:
        pd.DataFrame(np.mean(transmats, axis=0)).to_csv(os.path.join(OUT_DIR,'hmm_mean_transition_matrix.csv'))

    weights_df.to_csv(os.path.join(OUT_DIR,'hmm_allocations.csv'))
    diag_df.to_csv(os.path.join(OUT_DIR,'hmm_diagnostics.csv'))

    print("Terminé. Résultats sauvés dans:", OUT_DIR)
    return {'weights_df': weights_df, 'diag_df': diag_df}

if __name__ == '__main__':
    run_ssm_pipeline(ASSET_FILES, MACRO_FILES)
