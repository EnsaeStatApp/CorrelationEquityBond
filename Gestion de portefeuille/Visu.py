"""
portfolio_plot.py

Visualisation des résultats du HMM Portfolio (Version Finale Adaptative).
Lit les fichiers générés dans 'hmm_monthly_output'.
Fonctionnalités :
- Palette de couleurs dynamique (s'adapte à 2, 3 ou 4 régimes).
- Affichage "Long Only" (100% investissement max).
- Visualisation du Plafond de Volatilité.
"""

import os
import pandas as pd
import numpy as np
import matplotlib.pyplot as plt
import matplotlib.dates as mdates
import seaborn as sns

# Configuration
OUT_DIR = 'hmm_monthly_output'
ALLOC_FILE = os.path.join(OUT_DIR, 'monthly_allocations.csv')
DIAG_FILE = os.path.join(OUT_DIR, 'monthly_diagnostics.csv')

def run_visualization():
    # 1. Vérification et Chargement
    if not os.path.exists(ALLOC_FILE):
        print(f"Erreur: {ALLOC_FILE} introuvable. Lancez d'abord le script d'optimisation (portfolio_hmm_monthly.py).")
        return

    print(f"Chargement des données depuis {OUT_DIR}...")
    weights_df = pd.read_csv(ALLOC_FILE, index_col=0, parse_dates=True)
    diag_df = pd.read_csv(DIAG_FILE, index_col=0, parse_dates=True)
    
    print("Génération du Dashboard...")
    
    # Configuration style
    sns.set(style="whitegrid")
    fig, axes = plt.subplots(3, 1, figsize=(12, 16), sharex=True)

    # =========================================================================
    # GRAPHIQUE 1 : Probabilités des Régimes (Palette Adaptative)
    # =========================================================================
    ax1 = axes[0]
    # Détection automatique des colonnes de probabilité (p0, p1, p2...)
    prob_cols = [c for c in diag_df.columns if c.startswith('p') and c[1:].isdigit()]
    n_regimes = len(prob_cols)
    
    print(f"-> {n_regimes} régimes détectés dans les données.")

    # Logique de palette adaptative
    if n_regimes == 2:
        # Binaire : Rouge (Crise) vs Vert (Expansion)
        palette = ['#e74c3c', '#2ecc71'] 
    elif n_regimes == 3:
        # Ternaire : Rouge, Vert, Jaune
        palette = ['#e74c3c', '#2ecc71', '#f1c40f']
    elif n_regimes == 4:
        # Quaternaire : Rouge, Vert, Jaune, Bleu
        palette = ['#e74c3c', '#2ecc71', '#f1c40f', '#3498db']
    else:
        # Fallback (>4) : Palette automatique
        palette = sns.color_palette("tab10", n_regimes)

    # Création du Stackplot
    ax1.stackplot(diag_df.index, [diag_df[c] for c in prob_cols], 
                  labels=[f'Régime {i}' for i in range(n_regimes)],
                  alpha=0.85, colors=palette)
    
    ax1.set_title(f'1. Régimes Économiques ({n_regimes} États)', fontsize=12, fontweight='bold')
    ax1.set_ylabel('Probabilité')
    ax1.legend(loc='upper left', frameon=True)
    ax1.set_ylim(0, 1)
    ax1.grid(True, alpha=0.3)

    # =========================================================================
    # GRAPHIQUE 2 : Allocation d'Actifs (Strictement 100%)
    # =========================================================================
    ax2 = axes[1]
    w_cols = weights_df.columns
    
    # Couleurs actifs fixes (S&P=Bleu nuit, AAA=Gris, BAA=Orange, T10=Violet)
    asset_colors = ['#34495e', '#95a5a6', '#e67e22', '#9b59b6'] 
    
    ax2.stackplot(weights_df.index, [weights_df[c] for c in w_cols],
                  labels=w_cols, colors=asset_colors, alpha=0.9)
    
    ax2.set_title('2. Allocation Dynamique (Plafond 100% Investi)', fontsize=12, fontweight='bold')
    ax2.set_ylabel('Poids du Portefeuille')
    ax2.legend(loc='upper left', frameon=True)
    
    # Échelle 0-1.05 (Long Only strict)
    ax2.set_ylim(0, 1.05)
    ax2.grid(True, alpha=0.3)

    # =========================================================================
    # GRAPHIQUE 3 : Volatilité Réalisée vs Plafond Cible
    # =========================================================================
    ax3 = axes[2]
    
    # 1. Target Vol (Ligne Rouge - Plafond)
    target_vol = 0.10 # Valeur par défaut
    ax3.axhline(target_vol, color='red', linestyle='--', linewidth=1.5, label=f'Plafond Cible ({target_vol:.0%})', alpha=0.8)
    
    # 2. Volatilité Réalisée (Ligne Noire)
    if 'achieved_vol_annual' in diag_df.columns:
        vol_realisee = diag_df['achieved_vol_annual']
        ax3.plot(diag_df.index, vol_realisee, 
                 color='black', linewidth=1.5, label='Volatilité Prévue (Modèle)')
        
        # Zone grise : montre le "Budget risque non utilisé" (quand on est défensif)
        # On ne remplit que si la vol réalisée est inférieure à la cible
        ax3.fill_between(diag_df.index, vol_realisee, target_vol, 
                         where=(vol_realisee < target_vol),
                         color='gray', alpha=0.1, label='Capacité non utilisée (Risk Off)')

    # Échelle : 0% à 20% pour centrer visuellement la cible de 10%
    ax3.set_ylim(0, 0.20)
    
    ax3.set_title('3. Contrôle du Risque (Volatility Cap)', fontsize=12, fontweight='bold')
    ax3.set_ylabel('Volatilité Annuelle')
    ax3.legend(loc='upper left', frameon=True)
    ax3.grid(True, alpha=0.3)

    # Formatage axe X (Dates)
    axes[2].xaxis.set_major_locator(mdates.YearLocator(5)) # Tous les 5 ans
    axes[2].xaxis.set_minor_locator(mdates.YearLocator(1))
    axes[2].xaxis.set_major_formatter(mdates.DateFormatter('%Y'))
    
    plt.tight_layout()
    
    # Sauvegarde
    out_path = os.path.join(OUT_DIR, 'dashboard_adaptive.png')
    plt.savefig(out_path, dpi=300)
    print(f"Graphique sauvegardé avec succès : {out_path}")
    plt.show()

if __name__ == "__main__":
    run_visualization()