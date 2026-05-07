import numpy as np
import pandas as pd
from sklearn.linear_model import LogisticRegression
from src.stock_bond_correlation.models.HMMMarketRegimeDetector import HMMMarketRegimeDetector
from typing import List
import ssm
from src.stock_bond_correlation.diagnostics.statistical_tests.EmissionStatisticalAnalyzer import EmissionStatisticalAnalyzer


class HMMMarketWithMacroTransitionsRegimeDetector(HMMMarketRegimeDetector):
    """
    Wrapper for the ssm library to detect market regimes where transitions are 
    impacted by macroeconomic exogenous covariates (HMM TVTP).
    
    Inherits from HMMMarketRegimeDetector to maintain common methods.

    General Operation:
    Fits an HMM on historical data up to time T (Y_{1:t}, X_{1:t}) and predicts 
    regimes for T+1.

    Important Note:
    The ssm library uses X_{t} to determine the transition from state t-1 to state t. 
    Users should ensure the input covariates are appropriately lagged.
    """
    def __init__(self, n_regimes: int = 2, n_dim: int = 2, n_input: int = 0, **kwargs: dict):
        """
        Args:
            n_regimes (int): Number of regimes (default: 2).
            n_dim (int): Number of assets/dimensions (default: 2).
            n_input (int): Number of exogenous inputs (default: 0).
            **kwargs (dict): Dynamic dictionary for specific HMM TVTP parameters passed 
                             to the parent class.
        """
        # Set transition type to input-driven; no need for the user to specify it
        kwargs['transitions'] = "inputdriven" 
        super().__init__(n_regimes=n_regimes, n_dim=n_dim, n_input=n_input, **kwargs)

    def _display_logistic_weights(self, clf: LogisticRegression, feature_names: List[str]):
        """
        Displays the weights of the logistic regression used for parameter initialization.

        Args:
            clf (LogisticRegression): The fitted logistic regression model.
            feature_names (List[str]): Names of the features used in the regression.
        """
        print("\n=== INITIAL LOGISTIC REGRESSION WEIGHTS (X_{t-1} -> Z_t) ===")
        # clf.coef_ shape is (K, M)
        weights_df = pd.DataFrame(clf.coef_, columns=feature_names) 
        weights_df.index = [f"Regime {i} Probability" for i in range(self.K)]
        # Add the intercept (the base bias for each regime transition)
        weights_df.insert(0, 'Intercept', clf.intercept_) 

        print(weights_df.round(4))
        print("Note: A positive weight increases the probability of transitioning to that regime.")
        print("-" * 62)

    def fit(self, observations: List[np.ndarray], inputs: List[np.ndarray], log_returns: List[np.ndarray] = None, 
            asset_names: List[str] = None, feature_names: List[str] = None, series_index: int = 0,
            num_iters: int = 200, tolerance: float = 1e-4, method: str = "em", warm_start: bool = True, 
            initialize: bool = True, init_method: str = "kmeans"):
        """
        Fits the HMM TVTP model.
        
        When warm_start=True:
        - Emission parameters are initialized by fitting a stationary HMM on log returns.
        - Transition parameters are initialized via a multinomial logistic regression 
          where the output is the predicted regime at time t and features are covariates 
          at time t-1.
        Otherwise:
        - Emission parameters are initialized via K-means and transitions randomly.

        Args:
            observations (List[np.ndarray]): List of log return arrays.
            inputs (List[np.ndarray]): List of covariate arrays.
            log_returns (List[np.ndarray], optional): Ignored (included for API consistency).
            asset_names (List[str], optional): Names of the assets.
            feature_names (List[str], optional): Names of the covariates.
            series_index (int): Index of the sequence to store probabilities (default: 0).
            num_iters (int): Maximum EM iterations (default: 200).
            tolerance (float): Convergence tolerance.
            method (str): Optimization method (default: 'em').
            warm_start (bool): Whether to use HMM + Logistic Regression for initialization.
            initialize (bool): Whether to perform standard initialization (default: True).
            init_method (str): Method for initial state clustering (default: 'kmeans').
        """
        # Set default names if not provided
        if feature_names is None: 
            feature_names = [f"Input_{i}" for i in range(self.M)]

        if asset_names is None: 
            asset_names = [f"Asset_{i}" for i in range(self.D)]

        if warm_start: 
            # 1. Fit stationary HMM for emission parameter initialization
            transition_kwargs_hmm = {key: value for (key, value) in self.transition_kwargs.items() if key != "l2_penalty"}
            np.random.seed(self.random_state)
            simple_hmm = ssm.HMM(self.K, self.D, M=0,
                                 observations=self.observations_type,
                                 transitions="sticky",
                                 transition_kwargs=transition_kwargs_hmm)

            print("Initialization: Fitting stationary HMM...")
            np.random.seed(self.random_state)
            simple_hmm.fit(observations, method=method, num_iters=num_iters,
                           init_method=init_method, verbose=0, tolerance=tolerance, initialize=initialize)

            # 2. Display validation statistics for visual inspection of the initialization
            initial_states = simple_hmm.most_likely_states(observations[series_index])

            print("\n" + "="*70)
            print("STATIONARY HMM INITIALIZATION AUDIT (EMISSION)")
            print("="*70)

            init_analyzer = EmissionStatisticalAnalyzer(
                log_returns=observations[series_index],
                states=initial_states,
                asset_names=asset_names
            )
            df_init = init_analyzer.get_descriptive_stats_df()
            print(df_init.to_string(index=False)) 

            # 3. Multinomial Logistic Regression for transition weights W
            print("\n" + "="*70)
            print("LOGISTIC REGRESSION INITIALIZATION AUDIT (TRANSITION)")
            print("="*70)

            # Extract most likely states for each training sequence
            labels_list = [simple_hmm.most_likely_states(y) for y in observations]
            # Concatenate into a single sequence of states
            z_stacked = np.concatenate(labels_list) 
            # Stack the list of matrices into a single feature matrix
            X_stacked = np.vstack(inputs) 

            # Fit Logistic Regression (L2 penalty by default)
            clf = LogisticRegression(multi_class='multinomial', penalty='l2', C=1.0) 
            clf.fit(X_stacked, z_stacked)

            # 4. Display initialization weights for manual verification
            self._display_logistic_weights(clf, feature_names)

            # 5. Inject parameters into the final model
            # ssm expects Ws in shape (K, M) for InputDrivenTransitions
            self.model.transitions.Ws = clf.coef_ 
            # Initialize intercepts by tiling from the regression intercepts
            self.model.transitions.log_Ps = np.tile(clf.intercept_, (self.K, 1)) 
            self.model.observations.params = simple_hmm.observations.params

            Y_to_fit = observations
            X_to_fit = inputs
            final_initialize = False
        else:
            final_initialize = initialize
            Y_to_fit = observations
            X_to_fit = inputs

        # 7. Final fit of the Macro-Driven Model (HMM TVTP)
        np.random.seed(self.random_state)
        self.model.fit(Y_to_fit, inputs=X_to_fit, method=method, num_iters=num_iters,
                       init_method=init_method, verbose=0, tolerance=tolerance,
                       initialize=final_initialize)

        self.is_fitted = True
        self.fit_observations = Y_to_fit
        self.fit_input = X_to_fit
        # Store the most likely regimes (Viterbi path) for the in-sample period
        self.viterbi_states = self.model.most_likely_states(Y_to_fit[series_index], input=X_to_fit[series_index]) 
        return self

    def get_transition_weights(self):
        """
        Returns the weights (Ws) of the transition model.
        """
        return self.model.transitions.Ws
