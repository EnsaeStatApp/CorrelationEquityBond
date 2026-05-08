# Equity–Bond Correlation under Macroeconomic Regimes

**A regime detection, diagnostics, and allocation framework — 1970–2025**

[![Python](https://img.shields.io/badge/python-3.10+-blue.svg)](https://www.python.org/downloads/)
[![Jupyter](https://img.shields.io/badge/Jupyter-Notebook-orange.svg)](https://jupyter.org/)
[![Open In Colab](https://colab.research.google.com/assets/colab-badge.svg)](https://colab.research.google.com/github/EnsaeStatApp/CorrelationEquityBond)
[![Status](https://img.shields.io/badge/status-completed-brightgreen.svg)]()
[![License](https://img.shields.io/badge/license-academic-lightgrey.svg)]()

> The sign and magnitude of the equity–bond correlation are not constant: they regime-switch in response to macroeconomic conditions. This project builds an end-to-end framework — from Hidden Markov regime detection to risk-budgeted asset allocation — to identify, characterize, and exploit those regimes over the 1970–2025 period.

---

## Table of Contents

- [Research Question](#research-question)
- [Methodological Architecture](#methodological-architecture)
- [Three HMM Perspectives](#three-hmm-perspectives)
- [Repository Structure](#repository-structure)
- [Module Reference](#module-reference)
- [Data](#data)
- [Installation](#installation)
- [Validation Framework](#validation-framework)
- [Allocation Strategy](#allocation-strategy)
- [Historical Performance Comparison (1995–2025)](#historical-performance-comparison-19952025)
- [Authors](#authors)

---

## Research Question

> **Can we identify statistically robust latent regimes in stock–bond dynamics, and to what extent are their transitions driven by macroeconomic conditions? Finally, can this dual financial-macro understanding be exploited to outperform static benchmarks?**

The project answers this in three stages:

1. **Detection.** Fit Hidden Markov Models on financial and macro data to extract latent regimes.
2. **Diagnostics.** Validate the 4-regime structure via a four-pillar coherence framework and attribute transitions to macro drivers using **Integrated Gradients**.
3. **Allocation.** Translate regime probabilities into a risk-budgeted, proactive portfolio strategy.

## Methodological Architecture

```mermaid
flowchart LR
    A["Macro indicators<br/>(FRED)"] --> D{Detection layer}
    B["S&P 500 + 10y T-bond<br/>(log returns)"] --> D
    D --> M1["HMMMarketRegimeDetector<br/>(Pure Financial)"]
    D --> M2["HMMMacroRegimeDetector<br/>(Macro Cycle)"]
    D --> M3["HMMMarketWithMacroTransitions<br/>(Hybrid TVTP)"]
    M1 --> V[Validation suite]
    M2 --> V
    M3 --> V
    V --> V1[Emission tests]
    V --> V2[Transition inference]
    V --> V3[Coherence checker]
    V --> V4[Integrated Gradients]
    V --> WF[Walk-forward<br/>regime inference]
    WF --> ALLOC[Risk-budgeted<br/>allocator]
    ALLOC --> PERF[Performance vs Benchmarks]

```

## Three HMM Perspectives

The framework implements three complementary detectors to analyze the stock-bond relationship through different lenses:

| Detector | Observed signal | Transition logic | Purpose |
| --- | --- | --- | --- |
| **`HMMMarketRegimeDetector`** | Asset log returns | Constant (Sticky) | **Pure Financial Lens**: Identifies the 4 fundamental regimes of diversification directly from price action. Our statistical ground truth. |
| **`HMMMacroRegimeDetector`** | Macro indicators | Constant (Sticky) | **Economic Cycle Lens**: Characterizes how financial assets behave conditional on structural economic phases (Inflation, Growth). |
| **`HMMMarketWithMacroTransitions`** | Asset log returns | Input-driven (TVTP) | **The Hybrid Engine**: Uses macro covariates to anticipate regime shifts *before* they materialize in returns. |

## Repository Structure

```
CorrelationEquityBond/
│
├── data/
│   └── dataset.csv                      # Monthly macro + financial dataset (1970–2025)
│
├── notebooks/
│   ├── 01_Descriptive_Research/         # Calibration & In-Sample Analysis
│   │   ├── 01_HMM_Market_InSample.ipynb
│   │   ├── 02_HMM_Macro_InSample.ipynb
│   │   └── 03_HMMTVTP_Hybrid_InSample.ipynb
│   │
│   └── 02_WalkForward_Simulation/       # Proactive Trading Strategy
│       └── 04_WalkForward_Backtest_and_Explainability.ipynb
│
├── src/stock_bond_correlation/
│   ├── models/                          # HMM Architecture (Market, Macro, TVTP)
│   ├── diagnostics/                     # Coherence, Stats & Integrated Gradients
│   ├── backtest/                        # Walk-forward inference engine
│   ├── strategy/                        # Risk Budgeting & Performance metrics
│   └── visualization/                   # Regime Reporting & Equity curves
│
├── requirements.txt
└── README.md

```

## Module Reference

### Models (`src/.../models/`)

* `HMMMarketRegimeDetector.py`: Baseline model identifying regimes from returns alone.
* `HMMMarketWithMacroTransitionsRegimeDetector.py`: **The Core Innovation**. An Input-Driven HMM (TVTP) that bridges the gap between macro-economic cycles and financial market reactions.

### Diagnostics (`src/.../diagnostics/`)

* `HMMCoherenceChecker.py`: Orchestrates a four-pillar validation (Separability, Stability, Persistence, Confidence). Overall coherence score: **83.1%**.
* `TransitionExplainer.py`: **Integrated Gradients (IG)**. Decomposes any transition probability shift into per-covariate macro-contributions (e.g., Inflation, Slope).

### Strategy (`src/.../strategy/`)

* `RiskBudgetingAllocator.py`: Solves Spinu's log-utility convex formulation for optimal regime-aware weights.
* `StrategySimulator.py`: Applies volatility targeting, leverage caps, and stress-regime overlays (where equity risk budget shrinks during inflationary stress).

## Validation Framework

We enforce a **four-pillar validation protocol**:

1. **Separability**: Levene (vol) + Fisher-Z (corr) tests to ensure regimes are statistically distinct.
2. **Numerical stability**: Condition number of the observed Fisher information.
3. **Persistence**: Average regime duration (must be > 3 months).
4. **Confidence**: Average Viterbi posterior certainty (1 - Entropy).

## Historical Performance Comparison (1995–2025)

| **Metric** | **HMM TVTP (Strategy)** | **Rolling Risk Parity** | **Benchmark 60/40** |
| --- | --- | --- | --- |
| **Annual Return** | **9.26%** | 7.04% | 8.68% |
| **Annual Volatility** | **6.22%** | 5.64% | 7.67% |
| **Net Sharpe Ratio** | **1.08** | 0.82 | 0.81 |
| **Maximum Drawdown** | **-11.04%** | -19.13% | -27.00% |
| **Calmar Ratio** | **0.83** | 0.37 | 0.32 |

## Installation

```bash
git clone [https://github.com/EnsaeStatApp/CorrelationEquityBond.git](https://github.com/EnsaeStatApp/CorrelationEquityBond.git)
cd CorrelationEquityBond
pip install -r requirements.txt
pip install git+[https://github.com/lindermanlab/ssm.git](https://github.com/lindermanlab/ssm.git)

```

## Authors

**ENSAE — Statistical Applications Team**

* Benjamin Benisti
* Imade Haddadi
* Marie-Camille Memet
* Aurèle Thinot

## Acknowledgements

We would like to thank **Loic Brach** (HSBC) for his supervision and guidance throughout this project.
