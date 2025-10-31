import pandas as pd
import matplotlib.pyplot as plt

# Charger les données depuis la feuille Excel
df = pd.read_excel(
    'histretSP.xlsx',
    sheet_name='S&P 500 & Raw Data',
    skiprows=1,
    engine='openpyxl'
)

# Renommer les colonnes
df.columns = [
    'Year', 'S&P 500', 'Dividends', 'Dividend Yield', 'T.Bond rate',
    'Return on bond', 'Aaa Bond Rate', 'Return on Aaa', 'Baa Bond Rate',
    'Return on Baa', 'Returns on Real Estate'
]

# Conversion et mise en forme
df['Year'] = pd.to_numeric(df['Year'], errors='coerce').astype(int)
df.set_index('Year', inplace=True)
for col in ['S&P 500', 'Dividends', 'Return on bond']:
    df[col] = pd.to_numeric(df[col], errors='coerce')

# Calcul du rendement total du S&P 500
df['S&P 500 Total Return'] = (df['S&P 500'].diff() + df['Dividends']) / df['S&P 500'].shift(1)

# DataFrame des rendements
returns = df[['S&P 500 Total Return', 'Return on bond']].copy()
returns.dropna(inplace=True)
returns = returns[~returns.index.duplicated(keep='first')]
returns = returns.sort_index()

# Définir les groupes de fenêtres
windows_groups = {
    "Court terme (1-2 ans)": range(1, 3),
    "Moyen terme (3-5 ans)": range(3, 6),
    "Long terme (6-10 ans)": range(6, 11),
    "Toutes fenêtres (1-10 ans)": range(1, 11)
}

# Couleurs pour les courbes
colors = [
    '#1f77b4', '#ff7f0e', '#2ca02c', '#d62728', '#9467bd',
    '#8c564b', '#e377c2', '#7f7f7f', '#bcbd22', '#17becf'
]

# Créer la figure avec 4 subplots
fig, axes = plt.subplots(2, 2, figsize=(16, 12))
axes = axes.flatten()

# Fonction pour calculer les corrélations glissantes
def compute_rolling_corr(returns, window_range):
    rolling_corr = pd.DataFrame(index=returns.index)
    for window in window_range:
        rolling_corr[f'{window}-year corr'] = returns['S&P 500 Total Return'].rolling(window=window).corr(
            returns['Return on bond']
        )
    rolling_corr.dropna(how='all', inplace=True)
    return rolling_corr

# Tracer les sous-graphes
for ax, (title, windows) in zip(axes, windows_groups.items()):
    rolling_corr = compute_rolling_corr(returns, windows)
    for i, col in enumerate(rolling_corr.columns):
        ax.plot(
            rolling_corr.index,
            rolling_corr[col],
            label=col,
            color=colors[i],
            linewidth=2
        )
    ax.set_xlim(returns.index.min(), returns.index.max())
    ax.set_ylim(-1.05, 1.05)
    ax.set_title(title, fontsize=14)
    ax.set_xlabel("Année", fontsize=12)
    ax.set_ylabel("Corrélation", fontsize=12)
    ax.grid(True, linestyle='--', alpha=0.7)
    ax.legend(title="Fenêtre mobile", fontsize=9)

plt.tight_layout()
plt.savefig("rolling_corr_all_subplots.png")
plt.show()
