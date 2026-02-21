"""
Generate a structured day-by-day ML/TimeSeries/DL learning plan as an Excel workbook.
"""

import openpyxl
from openpyxl.styles import (
    Font, PatternFill, Alignment, Border, Side, numbers
)
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.datavalidation import DataValidation

# ── colour palette ──────────────────────────────────────────────
PHASE_COLORS = {
    "Phase 0: Prerequisites":                   "1B4F72",
    "Phase 1: Core Classical ML":               "117A65",
    "Phase 2: Trees, Ensembles & SVMs":         "6C3483",
    "Phase 3: Unsupervised Learning":           "B9770E",
    "Phase 4: Production ML Engineering":       "922B21",
    "Phase 5: Time Series Analysis":            "1A5276",
    "Phase 6: Deep Learning Foundations":       "0E6655",
    "Phase 7: Sequence Models & Transformers":  "7D3C98",
    "Phase 8: Generative Models":               "A04000",
    "Phase 9: Production Deep Learning":        "1F618D",
}

HEADER_FILL = PatternFill("solid", fgColor="2C3E50")
HEADER_FONT = Font(bold=True, color="FFFFFF", size=11, name="Calibri")
PHASE_FONT  = Font(bold=True, color="FFFFFF", size=10, name="Calibri")
BODY_FONT   = Font(size=10, name="Calibri")
WRAP_ALIGN   = Alignment(wrap_text=True, vertical="top")
CENTER_ALIGN = Alignment(horizontal="center", vertical="top", wrap_text=True)
THIN_BORDER  = Border(
    left=Side(style="thin", color="BDC3C7"),
    right=Side(style="thin", color="BDC3C7"),
    top=Side(style="thin", color="BDC3C7"),
    bottom=Side(style="thin", color="BDC3C7"),
)

# ── The complete curriculum ─────────────────────────────────────
# Each entry: (phase, topic, learning, implementation, deliverable, hours)
PLAN: list[tuple[str, str, str, str, str, float]] = []

def add(phase, topic, learn, impl, deliv, hrs=3.0):
    PLAN.append((phase, topic, learn, impl, deliv, hrs))

# ────────────────────────────────────────────────────────────────
# PHASE 0  –  Prerequisites  (Days 1-18)
# ────────────────────────────────────────────────────────────────
P0 = "Phase 0: Prerequisites"

add(P0, "Python OOP – Classes & Inheritance",
    "Classes, __init__, inheritance, super(), dunder methods",
    "Build a Matrix class with __add__, __mul__, __repr__",
    "matrix_class.py with unit tests", 3)

add(P0, "Python OOP – Advanced",
    "Decorators, generators, context managers, type hints",
    "Write decorators for timing & logging; a data-loading generator",
    "decorators.py, generators.py with tests", 3)

add(P0, "Project Setup & Tooling",
    "Virtual envs, pytest, project layout, Makefile, pre-commit",
    "Scaffold the ML-from-scratch repo with CI-ready structure",
    "Repo with pyproject.toml, tests/, Makefile", 2)

add(P0, "Git Workflow",
    "Branching, rebasing, conventional commits, PR workflow",
    "Create branches, practice rebase, write good commit messages",
    "Git cheat-sheet notes + practice repo history", 2)

add(P0, "NumPy – Basics",
    "Array creation, indexing, slicing, broadcasting, dtypes",
    "Implement dot product, matrix multiply, element-wise ops from scratch then with NumPy",
    "numpy_basics.py with exercises + tests", 3)

add(P0, "NumPy – Advanced",
    "Einsum, advanced indexing, strides, vectorization patterns",
    "Solve 10 vectorization problems (no for-loops)",
    "numpy_advanced.py with benchmarks vs loops", 3)

add(P0, "Linear Algebra – Vectors & Matrices",
    "Vector spaces, dot product, norms (L1/L2), matrix operations",
    "Implement vector/matrix ops in NumPy; verify with pen-and-paper",
    "linalg_basics.py with proofs as comments", 3)

add(P0, "Linear Algebra – Eigenvalues & SVD",
    "Eigendecomposition, SVD, rank, matrix inverse, pseudoinverse",
    "Implement power iteration for eigenvalues; compare with np.linalg.eig",
    "eigen_svd.py with numerical verification", 3)

add(P0, "Calculus – Derivatives & Gradients",
    "Derivatives, partial derivatives, chain rule, gradient vectors",
    "Implement numerical gradient checking function",
    "calculus.py: numerical vs analytical gradients", 3)

add(P0, "Calculus – Optimization Basics",
    "Gradient descent intuition, learning rate, convergence, convexity",
    "Implement gradient descent to minimize f(x)=x²+3x+2, visualize trajectory",
    "gradient_descent_1d.py with convergence plot", 3)

add(P0, "Probability – Fundamentals",
    "Sample spaces, conditional probability, Bayes' theorem, independence",
    "Implement Bayes' theorem solver; simulate Monty Hall problem",
    "probability.py with simulations", 3)

add(P0, "Statistics – Distributions",
    "Gaussian, Bernoulli, Binomial, Poisson, Uniform; PDF, CDF, PMF",
    "Implement PDF/CDF for Gaussian from scratch; compare with scipy",
    "distributions.py with plots", 3)

add(P0, "Statistics – Estimation & Testing",
    "MLE, MAP, confidence intervals, hypothesis testing, p-values",
    "Implement MLE for Gaussian parameters from data",
    "mle_estimation.py with synthetic data tests", 3)

add(P0, "Information Theory",
    "Entropy, cross-entropy, KL divergence, mutual information",
    "Implement entropy, cross-entropy, KL-divergence functions",
    "info_theory.py with unit tests", 3)

add(P0, "Pandas – Core Operations",
    "DataFrame, Series, indexing, filtering, groupby, merge, pivot",
    "Load a real dataset; clean, transform, aggregate, save",
    "pandas_eda.py + cleaned CSV output", 3)

add(P0, "Pandas – Advanced & Performance",
    "Window functions, apply vs vectorize, memory optimization, categorical dtype",
    "Optimize a slow Pandas pipeline (measure before/after)",
    "pandas_advanced.py with timing benchmarks", 3)

add(P0, "Data Visualization",
    "Matplotlib subplots, Seaborn statistical plots, best practices",
    "Create 8 plot types: scatter, line, bar, hist, box, heatmap, pair, violin",
    "visualization.py / notebook with 8 polished plots", 3)

add(P0, "EDA Mini-Project",
    "End-to-end exploratory data analysis on a real dataset",
    "Full EDA: missing values, distributions, correlations, feature insights",
    "eda_report.py/notebook with findings summary", 3)

# ────────────────────────────────────────────────────────────────
# PHASE 1  –  Core Classical ML  (Days 19-42)
# ────────────────────────────────────────────────────────────────
P1 = "Phase 1: Core Classical ML"

add(P1, "Linear Regression – Theory",
    "OLS, normal equation, cost function (MSE), assumptions",
    "Derive normal equation; implement closed-form solution in NumPy",
    "linear_reg_theory.py with derivation comments", 3)

add(P1, "Linear Regression – Gradient Descent",
    "Batch gradient descent for linear regression, convergence analysis",
    "Implement GD-based linear regression from scratch",
    "linear_reg_gd.py with cost curve plot", 3)

add(P1, "Linear Regression – SGD & Mini-Batch",
    "SGD, mini-batch GD, learning rate schedules",
    "Implement SGD and mini-batch GD; compare convergence",
    "linear_reg_sgd.py with comparison plots", 3)

add(P1, "Linear Regression – sklearn & Evaluation",
    "sklearn LinearRegression, R², adjusted R², MAE, RMSE",
    "Train sklearn model; compare with scratch impl on Boston-like dataset",
    "linear_reg_sklearn.py with metrics comparison table", 3)

add(P1, "Polynomial Regression & Overfitting",
    "Feature engineering (polynomial features), overfitting, underfitting",
    "Implement polynomial regression; plot degree 1-15 bias-variance",
    "poly_regression.py with overfitting visualization", 3)

add(P1, "Regularization – Ridge (L2)",
    "L2 penalty, closed form with regularization, hyperparameter λ",
    "Implement Ridge regression from scratch; compare with sklearn",
    "ridge_regression.py with regularization path plot", 3)

add(P1, "Regularization – Lasso (L1) & Elastic Net",
    "L1 penalty, sparsity, coordinate descent, Elastic Net",
    "Implement Lasso with coordinate descent; implement Elastic Net",
    "lasso_elasticnet.py with coefficient path plots", 3)

add(P1, "Bias-Variance Tradeoff Deep Dive",
    "Decomposition, model complexity, learning curves, validation curves",
    "Generate learning curves and validation curves for different models",
    "bias_variance.py with diagnostic plots", 3)

add(P1, "Cross-Validation",
    "K-fold, stratified K-fold, leave-one-out, time-series split",
    "Implement K-fold CV from scratch; compare with sklearn cross_val_score",
    "cross_validation.py with from-scratch + sklearn comparison", 3)

add(P1, "Evaluation Metrics – Regression",
    "MSE, RMSE, MAE, MAPE, R², adjusted R²; when to use which",
    "Implement all metrics from scratch; build a metrics dashboard",
    "regression_metrics.py with unit tests", 2)

add(P1, "Logistic Regression – Theory",
    "Sigmoid, log-loss / binary cross-entropy, decision boundary",
    "Derive gradient of log-loss; implement logistic regression from scratch",
    "logistic_reg.py with decision boundary plot", 3)

add(P1, "Logistic Regression – Multiclass & Regularization",
    "Softmax, one-vs-rest, one-vs-one, L2 regularization",
    "Extend to multiclass; add regularization; compare with sklearn",
    "logistic_reg_multi.py with confusion matrix plot", 3)

add(P1, "Evaluation Metrics – Classification",
    "Accuracy, precision, recall, F1, confusion matrix, ROC-AUC, PR-AUC",
    "Implement all metrics from scratch; plot ROC & PR curves",
    "classification_metrics.py with plots & tests", 3)

add(P1, "Naive Bayes",
    "Bayes theorem, Gaussian NB, Multinomial NB, Bernoulli NB, Laplace smoothing",
    "Implement Gaussian NB from scratch; test on Iris; compare with sklearn",
    "naive_bayes.py with accuracy comparison", 3)

add(P1, "K-Nearest Neighbors",
    "Distance metrics (L1/L2/Minkowski), K selection, curse of dimensionality",
    "Implement KNN from scratch (classification + regression); add KD-tree optimization",
    "knn.py with distance metric comparison", 3)

add(P1, "Feature Scaling & Encoding",
    "StandardScaler, MinMaxScaler, RobustScaler, one-hot, ordinal, target encoding",
    "Implement scalers from scratch; build custom transformer classes",
    "preprocessing.py with sklearn-compatible classes", 3)

add(P1, "Feature Selection",
    "Filter (correlation, chi²), wrapper (RFE), embedded (L1), mutual information",
    "Implement correlation-based filter + RFE from scratch",
    "feature_selection.py with comparison on real dataset", 3)

add(P1, "sklearn Pipelines & ColumnTransformer",
    "Pipeline, ColumnTransformer, custom transformers, make_pipeline",
    "Build end-to-end pipeline: preprocess → train → evaluate",
    "sklearn_pipeline.py with full pipeline code", 3)

add(P1, "Handling Missing Data & Imbalanced Data",
    "Imputation strategies, SMOTE, class weights, under/oversampling",
    "Implement simple imputer + SMOTE from scratch; compare with imblearn",
    "missing_imbalanced.py with before/after metrics", 3)

add(P1, "Classical ML Mini-Project",
    "End-to-end supervised learning project on a real dataset",
    "Full pipeline: EDA → preprocess → model selection → CV → evaluation report",
    "Jupyter notebook or .py with final report", 4)

# Catch-up / revision days
add(P1, "Phase 1 Revision & Practice",
    "Review all Phase 1 concepts; solidify weak areas",
    "Re-implement weakest algorithm; solve 5 practice problems",
    "Revised code + self-assessment notes", 3)

add(P1, "Phase 1 Revision & Practice (contd.)",
    "Practice on Kaggle-style dataset with full pipeline",
    "End-to-end prediction on a new dataset using best practices",
    "Complete submission-ready solution", 3)

# ────────────────────────────────────────────────────────────────
# PHASE 2  –  Trees, Ensembles & SVMs  (Days 43-68)
# ────────────────────────────────────────────────────────────────
P2 = "Phase 2: Trees, Ensembles & SVMs"

add(P2, "Decision Trees – Theory",
    "CART algorithm, Gini impurity, entropy, information gain, pruning",
    "Implement Gini & entropy calculations from scratch",
    "decision_tree_theory.py with metric comparisons", 3)

add(P2, "Decision Tree Classifier – From Scratch",
    "Recursive splitting, stopping criteria, tree structure (nodes/leaves)",
    "Implement full decision tree classifier from scratch",
    "decision_tree_classifier.py with predict + visualize", 4)

add(P2, "Decision Tree Regressor – From Scratch",
    "Variance reduction, MSE splitting, leaf prediction (mean)",
    "Implement decision tree regressor from scratch",
    "decision_tree_regressor.py with comparison to sklearn", 3)

add(P2, "Decision Trees – Pruning & sklearn",
    "Pre-pruning (max_depth, min_samples), post-pruning (cost-complexity)",
    "Add pruning to scratch impl; compare with sklearn DecisionTreeClassifier",
    "decision_tree_pruning.py with depth vs accuracy plot", 3)

add(P2, "Bagging & Random Subspaces",
    "Bootstrap aggregating, random subspace method, out-of-bag error",
    "Implement bagging classifier from scratch using your decision tree",
    "bagging.py with OOB error estimation", 3)

add(P2, "Random Forest – From Scratch",
    "Feature randomness, ensemble averaging, feature importance",
    "Implement Random Forest from scratch using your trees",
    "random_forest.py with feature importance plot", 4)

add(P2, "Random Forest – sklearn & Tuning",
    "sklearn RandomForest, n_estimators, max_features, tuning strategies",
    "Tune RF on real dataset; compare scratch vs sklearn accuracy/speed",
    "random_forest_sklearn.py with tuning results", 3)

add(P2, "AdaBoost – From Scratch",
    "Adaptive boosting, sample weights, weighted voting, exponential loss",
    "Implement AdaBoost with decision stumps from scratch",
    "adaboost.py with weight evolution visualization", 4)

add(P2, "Gradient Boosting – Theory",
    "Boosting as gradient descent in function space, pseudo-residuals, shrinkage",
    "Understand and document the math of gradient boosting",
    "Notes + derivation of GBM update rules", 3)

add(P2, "Gradient Boosting – From Scratch",
    "Sequential tree fitting on residuals, learning rate, n_estimators",
    "Implement GBM regressor from scratch",
    "gradient_boosting.py with staged prediction plots", 4)

add(P2, "Gradient Boosting Classifier – From Scratch",
    "Log-loss pseudo-residuals, leaf value optimization for classification",
    "Implement GBM classifier from scratch",
    "gradient_boosting_clf.py with log-loss curve", 4)

add(P2, "XGBoost – Deep Dive",
    "Regularized objective, approximate greedy, histogram-based, XGB API",
    "Train, tune XGBoost on real dataset; understand all key params",
    "xgboost_tuning.py with param importance analysis", 3)

add(P2, "LightGBM & CatBoost",
    "Leaf-wise growth (LGBM), ordered boosting (CatBoost), categorical handling",
    "Train & compare XGB vs LGBM vs CatBoost on same dataset",
    "boosting_comparison.py with speed & accuracy comparison table", 3)

add(P2, "Stacking & Blending",
    "Stacked generalization, blending, meta-learner, out-of-fold predictions",
    "Implement stacking from scratch with out-of-fold predictions",
    "stacking.py with performance vs individual models", 3)

add(P2, "Voting Classifier",
    "Hard voting, soft voting, weighted voting",
    "Implement voting classifier from scratch; compare with sklearn",
    "voting.py with ensemble vs individual comparison", 2)

add(P2, "SVM – Linear (Theory)",
    "Maximum margin, support vectors, hinge loss, primal vs dual",
    "Derive the SVM optimization problem; implement linear SVM with SGD",
    "svm_linear_theory.py with margin visualization", 3)

add(P2, "SVM – Linear (From Scratch)",
    "Sub-gradient descent for hinge loss, regularization",
    "Implement linear SVM classifier from scratch",
    "svm_linear.py with decision boundary + support vectors plot", 4)

add(P2, "SVM – Kernel Trick",
    "Kernel functions (RBF, polynomial, sigmoid), kernel matrix, dual formulation",
    "Implement kernel SVM using SMO or with sklearn; visualize non-linear boundaries",
    "svm_kernel.py with kernel comparison plots", 3)

add(P2, "SVM – Regression & Multiclass",
    "SVR (epsilon-tube), one-vs-rest, one-vs-one for multiclass SVM",
    "Train SVR and multiclass SVM with sklearn; tune C, gamma, epsilon",
    "svr_multiclass.py with tuning results", 3)

add(P2, "Hyperparameter Tuning – GridSearch & RandomSearch",
    "Exhaustive grid search, random search, search spaces, CV integration",
    "Implement grid search from scratch; compare with sklearn GridSearchCV",
    "hyperparam_grid_random.py with search visualization", 3)

add(P2, "Hyperparameter Tuning – Bayesian (Optuna)",
    "Bayesian optimization, TPE, Optuna API, pruning bad trials",
    "Set up Optuna study for XGBoost; visualize optimization history",
    "optuna_tuning.py with optimization plots", 3)

add(P2, "LDA & QDA",
    "Linear Discriminant Analysis, Quadratic DA, assumptions, Fisher's criterion",
    "Implement LDA from scratch (Fisher's); compare with sklearn",
    "lda_qda.py with decision boundary comparison", 3)

add(P2, "Ensemble Methods Mini-Project",
    "Apply all ensemble techniques to a challenging dataset",
    "Compare DT, RF, AdaBoost, GBM, XGB, LGBM, Stacking on one dataset",
    "ensemble_comparison.py with results table & analysis", 4)

add(P2, "Phase 2 Revision & Practice",
    "Review trees, ensembles, SVMs; strengthen weak areas",
    "Re-implement one algorithm; Kaggle practice problem",
    "Revised code + practice submission", 3)

add(P2, "Phase 2 Revision & Practice (contd.)",
    "Kaggle-style competition practice with full pipeline",
    "End-to-end solution using best ensemble methods",
    "Complete competition-ready notebook", 3)

# ────────────────────────────────────────────────────────────────
# PHASE 3  –  Unsupervised Learning  (Days 69-88)
# ────────────────────────────────────────────────────────────────
P3 = "Phase 3: Unsupervised Learning"

add(P3, "K-Means – Theory & From Scratch",
    "Lloyd's algorithm, centroid initialization (K-Means++), inertia, elbow method",
    "Implement K-Means from scratch with K-Means++ init",
    "kmeans.py with elbow plot + cluster visualization", 3)

add(P3, "K-Means – Advanced",
    "Mini-batch K-Means, silhouette analysis, choosing K, limitations",
    "Implement mini-batch variant; compute silhouette scores",
    "kmeans_advanced.py with silhouette analysis", 3)

add(P3, "K-Medoids & DBSCAN – Theory",
    "PAM algorithm, density-based clustering, core/border/noise points, eps & minPts",
    "Understand density-based vs centroid-based clustering",
    "Theory notes + comparison analysis", 2)

add(P3, "DBSCAN – From Scratch",
    "Region queries, density reachability, cluster expansion",
    "Implement DBSCAN from scratch",
    "dbscan.py with noise detection visualization", 4)

add(P3, "HDBSCAN & OPTICS",
    "Hierarchical density-based clustering, reachability plots",
    "Use HDBSCAN library; compare with DBSCAN on same data",
    "hdbscan_optics.py with comparison plots", 3)

add(P3, "Hierarchical Clustering",
    "Agglomerative (single/complete/average/Ward linkage), dendrogram",
    "Implement agglomerative clustering from scratch; plot dendrogram",
    "hierarchical.py with dendrogram + cluster plots", 3)

add(P3, "Gaussian Mixture Models – Theory",
    "Mixture of Gaussians, EM algorithm, E-step, M-step, convergence",
    "Derive EM update equations for GMM",
    "Theory notes + derivation", 3)

add(P3, "GMM – From Scratch",
    "Implement full EM algorithm for GMM",
    "Implement GMM from scratch; compare with sklearn GaussianMixture",
    "gmm.py with cluster probability contour plots", 4)

add(P3, "Mean Shift & Spectral Clustering",
    "Kernel density estimation, bandwidth, graph Laplacian, spectral methods",
    "Implement mean shift from scratch; use sklearn SpectralClustering",
    "meanshift_spectral.py with non-convex cluster demos", 3)

add(P3, "Clustering Evaluation & Comparison",
    "Silhouette, Calinski-Harabasz, Davies-Bouldin, ARI, NMI (external)",
    "Implement internal metrics; compare 5 clustering algorithms on 3 datasets",
    "clustering_evaluation.py with comprehensive comparison table", 3)

add(P3, "PCA – Theory & From Scratch",
    "Covariance matrix, eigendecomposition, variance explained, scree plot",
    "Implement PCA from scratch using eigen decomposition",
    "pca.py with scree plot + 2D projection", 3)

add(P3, "PCA – Advanced & Kernel PCA",
    "Incremental PCA, kernel PCA (RBF), choosing components, reconstruction",
    "Implement kernel PCA; use sklearn IncrementalPCA on large data",
    "pca_advanced.py with kernel PCA non-linear demo", 3)

add(P3, "SVD & Truncated SVD",
    "Full SVD, truncated SVD, relationship to PCA, LSA/LSI applications",
    "Implement truncated SVD for dimensionality reduction; apply to text data",
    "svd.py with text topic extraction demo", 3)

add(P3, "t-SNE & UMAP",
    "t-SNE: perplexity, KL divergence; UMAP: topological approach, parameters",
    "Apply t-SNE and UMAP to MNIST subset; compare visualizations",
    "tsne_umap.py with side-by-side comparison", 3)

add(P3, "ICA & NMF",
    "Independent Component Analysis, Non-negative Matrix Factorization, applications",
    "Apply ICA to signal separation; NMF to image decomposition",
    "ica_nmf.py with signal separation + face parts demo", 3)

add(P3, "Factor Analysis",
    "Latent factors, factor loadings, communalities, vs PCA",
    "Implement factor analysis; compare with PCA on same dataset",
    "factor_analysis.py with loading matrix visualization", 3)

add(P3, "Anomaly Detection – Isolation Forest & LOF",
    "Isolation tree, path length, local outlier factor, k-distance",
    "Implement Isolation Forest from scratch; use sklearn LOF",
    "anomaly_detection.py with anomaly visualization", 3)

add(P3, "Anomaly Detection – One-Class SVM & Elliptic Envelope",
    "One-Class SVM, Mahalanobis distance, Gaussian assumption",
    "Compare all anomaly methods on synthetic + real data",
    "anomaly_comparison.py with method comparison table", 3)

add(P3, "Unsupervised Learning Mini-Project",
    "Customer segmentation or image compression using unsupervised methods",
    "Full project: EDA → dim reduction → clustering → profiling → visualization",
    "Complete project notebook with business insights", 4)

add(P3, "Phase 3 Revision",
    "Review all unsupervised algorithms; solidify weak spots",
    "Re-implement weakest algorithm; practice clustering challenge",
    "Revised code + self-assessment", 3)

# ────────────────────────────────────────────────────────────────
# PHASE 4  –  Production ML Engineering  (Days 89-100)
# ────────────────────────────────────────────────────────────────
P4 = "Phase 4: Production ML Engineering"

add(P4, "Feature Engineering – Numerical",
    "Binning, log/power transforms, interaction features, polynomial features",
    "Build a feature engineering toolkit with reusable transformers",
    "feature_eng_numerical.py with sklearn-compatible transformers", 3)

add(P4, "Feature Engineering – Categorical & Text",
    "Target encoding, frequency encoding, TF-IDF, hashing trick",
    "Implement target encoder from scratch; build TF-IDF pipeline",
    "feature_eng_categorical.py + text pipeline", 3)

add(P4, "sklearn Custom Transformers & Pipelines",
    "BaseEstimator, TransformerMixin, custom transformers, FeatureUnion",
    "Build 3 production-quality custom transformers; integrate in pipeline",
    "custom_transformers.py that works in sklearn Pipeline", 3)

add(P4, "Model Serialization & Versioning",
    "Pickle, joblib, ONNX export, model versioning strategies",
    "Serialize trained model in all 3 formats; load & verify predictions match",
    "model_serialization.py with format comparison", 3)

add(P4, "Model Serving with FastAPI",
    "REST API design, FastAPI endpoints, request/response schemas, Pydantic",
    "Build a prediction API with FastAPI for a trained model",
    "app.py + test_api.py with /predict endpoint", 4)

add(P4, "Docker for ML",
    "Dockerfile, multi-stage builds, ML-specific Docker patterns",
    "Containerize the FastAPI model server",
    "Dockerfile + docker-compose.yml + README", 3)

add(P4, "Experiment Tracking – MLflow",
    "MLflow tracking, experiments, runs, metrics, artifacts, model registry",
    "Set up MLflow; log experiments for 3 models with metrics & artifacts",
    "mlflow_tracking.py with comparison dashboard", 3)

add(P4, "Data Versioning – DVC",
    "DVC basics, data pipelines, remote storage, reproducibility",
    "Set up DVC in project; version a dataset through pipeline stages",
    "dvc.yaml + data pipeline", 3)

add(P4, "Monitoring & Drift Detection",
    "Data drift, concept drift, PSI, KS test, performance monitoring",
    "Implement drift detector using PSI and KS test",
    "drift_detection.py with alert simulation", 3)

add(P4, "CI/CD for ML",
    "GitHub Actions for ML, automated testing, model validation gates",
    "Write GitHub Actions workflow: lint → test → train → validate → deploy",
    ".github/workflows/ml_pipeline.yml", 3)

add(P4, "End-to-End ML Project",
    "Full production pipeline: data → feature store → train → serve → monitor",
    "Build complete ML system with all production components",
    "ml_system/ project with full pipeline", 5)

add(P4, "Phase 4 Revision & Portfolio Polish",
    "Review production concepts; polish code for portfolio",
    "Clean up all code; add docstrings, type hints, README files",
    "Production-ready repo with polished code", 3)

# ────────────────────────────────────────────────────────────────
# PHASE 5  –  Time Series Analysis  (Days 101-130)
# ────────────────────────────────────────────────────────────────
P5 = "Phase 5: Time Series Analysis"

add(P5, "Time Series – Fundamentals",
    "Time series components: trend, seasonality, cyclical, noise; stationarity",
    "Load & visualize real time series; decompose into components",
    "ts_fundamentals.py with component plots", 3)

add(P5, "Stationarity & Transforms",
    "ADF test, KPSS test, differencing, log transforms, Box-Cox",
    "Implement ADF test interpretation; make non-stationary series stationary",
    "stationarity.py with before/after plots & test results", 3)

add(P5, "Autocorrelation & Partial Autocorrelation",
    "ACF, PACF, correlogram interpretation, lag selection",
    "Implement ACF from scratch; plot ACF/PACF for different processes",
    "acf_pacf.py with interpretation guide plots", 3)

add(P5, "White Noise & Random Walk",
    "Properties of white noise, random walk, drift; simulation",
    "Simulate WN and RW processes; test stationarity of each",
    "wn_random_walk.py with simulations + tests", 2)

add(P5, "Moving Average Models (MA)",
    "MA(q) process, parameter estimation, invertibility",
    "Simulate MA processes; fit MA model using statsmodels",
    "ma_model.py with simulation + fitting", 3)

add(P5, "Autoregressive Models (AR)",
    "AR(p) process, Yule-Walker equations, parameter estimation, stability",
    "Implement AR model from scratch; compare with statsmodels",
    "ar_model.py with from-scratch + comparison", 3)

add(P5, "ARMA Model",
    "ARMA(p,q), combining AR and MA, model identification",
    "Fit ARMA models; use AIC/BIC for order selection",
    "arma_model.py with order selection analysis", 3)

add(P5, "ARIMA – Theory & Implementation",
    "Integrated component, differencing, ARIMA(p,d,q), Box-Jenkins methodology",
    "Implement ARIMA pipeline: stationarity check → order selection → fit → forecast",
    "arima.py with full Box-Jenkins pipeline", 4)

add(P5, "ARIMA – Diagnostics & Forecasting",
    "Residual analysis, Ljung-Box test, confidence intervals, multi-step forecast",
    "Diagnose fitted ARIMA; generate forecasts with confidence bands",
    "arima_diagnostics.py with residual plots + forecast", 3)

add(P5, "SARIMA",
    "Seasonal ARIMA, seasonal differencing, SARIMA(p,d,q)(P,D,Q,s)",
    "Fit SARIMA to monthly data with clear seasonality",
    "sarima.py with seasonal decomposition + forecast", 3)

add(P5, "Auto ARIMA & pmdarima",
    "Automatic order selection, stepwise search, pmdarima API",
    "Use pmdarima auto_arima; compare with manual selection",
    "auto_arima.py with manual vs auto comparison", 2)

add(P5, "Exponential Smoothing – SES & DES",
    "Simple ES, double ES (Holt's method), alpha/beta parameters",
    "Implement SES and Holt's method from scratch",
    "exponential_smoothing.py with parameter tuning", 3)

add(P5, "Exponential Smoothing – Holt-Winters",
    "Triple ES, additive vs multiplicative seasonality",
    "Implement Holt-Winters from scratch; compare with statsmodels",
    "holt_winters.py with additive vs multiplicative comparison", 4)

add(P5, "ETS Framework",
    "Error-Trend-Seasonality taxonomy, state space models, ETS(A,A,A) etc.",
    "Use statsmodels ETSModel; compare different ETS specifications",
    "ets_framework.py with model comparison table", 3)

add(P5, "GARCH & Volatility Models",
    "ARCH, GARCH(1,1), volatility clustering, financial time series",
    "Fit GARCH model to stock returns; forecast volatility",
    "garch.py with volatility forecast plots", 3)

add(P5, "VAR – Vector Autoregression",
    "Multivariate time series, VAR(p), Granger causality, impulse response",
    "Fit VAR model to multivariate economic data; test Granger causality",
    "var_model.py with impulse response + causality tests", 3)

add(P5, "VECM & Cointegration",
    "Cointegration, Johansen test, VECM, error correction",
    "Test for cointegration; fit VECM to cointegrated series",
    "vecm.py with cointegration test + VECM forecast", 3)

add(P5, "Prophet",
    "Additive model, trend changepoints, seasonality, holidays, Prophet API",
    "Fit Prophet to real data; add holidays & custom seasonality",
    "prophet_model.py with component plots + forecast", 3)

add(P5, "Time Series Cross-Validation",
    "Rolling window, expanding window, TimeSeriesSplit, walk-forward validation",
    "Implement walk-forward validation from scratch; compare strategies",
    "ts_cross_validation.py with strategy comparison", 3)

add(P5, "Time Series Feature Engineering",
    "Lag features, rolling stats, Fourier features, date features, target encoding",
    "Build comprehensive TS feature engineering pipeline",
    "ts_feature_engineering.py with feature importance", 3)

add(P5, "ML for Time Series – Regression Approach",
    "XGBoost/LightGBM for time series, feature-based forecasting",
    "Build ML forecast model with lag features + gradient boosting",
    "ts_ml_forecast.py with ML vs ARIMA comparison", 3)

add(P5, "Multiple Seasonality & Complex Patterns",
    "TBATS, multiple seasonal decomposition, Fourier terms",
    "Handle data with daily + weekly + yearly seasonality",
    "multi_seasonality.py with complex pattern handling", 3)

add(P5, "Anomaly Detection in Time Series",
    "STL decomposition residuals, rolling Z-score, isolation forest for TS",
    "Build time series anomaly detector; test on real data with known anomalies",
    "ts_anomaly_detection.py with detected anomaly plots", 3)

add(P5, "Time Series Forecasting at Scale",
    "Hierarchical forecasting, forecast reconciliation, many-series patterns",
    "Forecast multiple related series with reconciliation",
    "hierarchical_forecast.py with reconciliation demo", 3)

add(P5, "Time Series Mini-Project – Classical",
    "End-to-end forecasting project using classical methods",
    "Full pipeline: EDA → stationarity → model selection → forecast → evaluate",
    "Complete classical TS forecasting project", 4)

add(P5, "Time Series Mini-Project – ML Approach",
    "End-to-end forecasting using ML/gradient boosting",
    "Full pipeline: feature engineering → model → tune → forecast → evaluate",
    "Complete ML-based TS forecasting project", 4)

add(P5, "TS Method Comparison & Best Practices",
    "When to use ARIMA vs ETS vs Prophet vs ML; production considerations",
    "Compare all methods on 3 datasets; write best-practices guide",
    "ts_comparison.py + best_practices.md with decision framework", 3)

add(P5, "Phase 5 Revision",
    "Review all time series methods; strengthen weak areas",
    "Re-implement weakest method; practice on new dataset",
    "Revised code + self-assessment", 3)

# ────────────────────────────────────────────────────────────────
# PHASE 6  –  Deep Learning Foundations  (Days 131-162)
# ────────────────────────────────────────────────────────────────
P6 = "Phase 6: Deep Learning Foundations"

add(P6, "Perceptron – From Scratch",
    "Single neuron, step function, perceptron learning rule, convergence",
    "Implement perceptron from scratch; train on linearly separable data",
    "perceptron.py with decision boundary animation", 3)

add(P6, "MLP – Forward Propagation",
    "Multi-layer architecture, weight matrices, activation functions, forward pass",
    "Implement forward pass for 3-layer MLP in NumPy",
    "mlp_forward.py with shape verification", 3)

add(P6, "Backpropagation – Theory",
    "Chain rule through computation graph, weight gradients, bias gradients",
    "Derive backprop equations for 2-layer network on paper + code",
    "backprop_derivation.py with step-by-step gradient computation", 3)

add(P6, "MLP – Full From Scratch",
    "Complete MLP with backprop, training loop, mini-batch SGD",
    "Implement full MLP from scratch in NumPy; train on MNIST",
    "mlp_from_scratch.py achieving >95% on MNIST", 4)

add(P6, "Activation Functions Deep Dive",
    "Sigmoid, tanh, ReLU, LeakyReLU, ELU, GELU, Swish; vanishing gradient problem",
    "Implement all activations + derivatives; visualize gradient flow",
    "activations.py with comparison plots + gradient analysis", 3)

add(P6, "Loss Functions Deep Dive",
    "MSE, BCE, categorical CE, hinge, Huber; when to use which",
    "Implement all loss functions + gradients; verify with numerical gradient check",
    "loss_functions.py with gradient verification", 3)

add(P6, "Optimizers – From Scratch",
    "SGD, Momentum, Nesterov, AdaGrad, RMSProp, Adam, AdamW",
    "Implement SGD, Momentum, Adam from scratch; compare convergence",
    "optimizers.py with convergence comparison on Rosenbrock", 4)

add(P6, "Weight Initialization & Batch Norm",
    "Xavier/Glorot, He, initialization impact; batch normalization forward/backward",
    "Implement He init + batch norm from scratch; show training impact",
    "init_batchnorm.py with training curves comparison", 3)

add(P6, "Regularization in DL",
    "L2 weight decay, dropout (inverted), early stopping, data augmentation",
    "Implement dropout from scratch (train vs eval mode); show regularization effect",
    "dl_regularization.py with overfitting experiment", 3)

add(P6, "Learning Rate Scheduling",
    "Step decay, exponential, cosine annealing, warmup, cyclical LR, OneCycleLR",
    "Implement 4 LR schedulers from scratch; compare training dynamics",
    "lr_schedulers.py with schedule plots + training comparison", 3)

add(P6, "PyTorch – Fundamentals",
    "Tensors, autograd, computation graph, GPU operations, requires_grad",
    "Translate NumPy MLP to PyTorch using raw autograd (no nn.Module yet)",
    "pytorch_basics.py with autograd verification", 3)

add(P6, "PyTorch – nn.Module & Training",
    "nn.Module, nn.Linear, nn.Sequential, loss functions, optimizers",
    "Build MLP with nn.Module; train on MNIST with proper training loop",
    "pytorch_mlp.py with training/validation curves", 3)

add(P6, "PyTorch – Dataset & DataLoader",
    "torch.utils.data Dataset, DataLoader, transforms, custom datasets",
    "Build custom Dataset class; implement data augmentation pipeline",
    "pytorch_data.py with custom dataset + augmentation", 3)

add(P6, "CNN – Theory",
    "Convolution operation, filters/kernels, stride, padding, pooling, receptive field",
    "Implement 2D convolution from scratch in NumPy; verify with PyTorch",
    "conv2d_scratch.py with visualization of feature maps", 3)

add(P6, "CNN – From Scratch in NumPy",
    "Full CNN forward + backward pass in NumPy (conv + pool + FC)",
    "Implement simple CNN from scratch for MNIST",
    "cnn_from_scratch.py (slow but educational)", 4)

add(P6, "CNN – PyTorch Implementation",
    "nn.Conv2d, nn.MaxPool2d, CNN architectures, training CNNs",
    "Build CNN in PyTorch; train on CIFAR-10",
    "cnn_pytorch.py achieving >80% on CIFAR-10", 3)

add(P6, "Classic CNN Architectures",
    "LeNet-5, AlexNet, VGG-16 architecture study, depth & width trade-offs",
    "Implement VGG-like network in PyTorch; analyze parameter counts",
    "classic_cnns.py with architecture visualization", 3)

add(P6, "ResNet & Skip Connections",
    "Residual learning, skip connections, identity mapping, ResNet-18/34/50",
    "Implement ResNet-18 in PyTorch; demonstrate skip connection benefit",
    "resnet.py with depth experiment (with/without skip)", 4)

add(P6, "Transfer Learning & Fine-Tuning",
    "Pre-trained models, feature extraction, fine-tuning strategies, freezing layers",
    "Fine-tune pre-trained ResNet on custom dataset (e.g., flowers)",
    "transfer_learning.py with feature extract vs fine-tune comparison", 3)

add(P6, "Data Augmentation & Advanced Training",
    "Albumentations, CutOut, MixUp, CutMix, test-time augmentation",
    "Implement MixUp from scratch; use Albumentations pipeline",
    "augmentation.py with before/after accuracy comparison", 3)

add(P6, "Model Debugging & Visualization",
    "Gradient checking, activation visualization, Grad-CAM, confusion analysis",
    "Implement Grad-CAM; visualize what CNN has learned",
    "model_debugging.py with Grad-CAM heatmaps", 3)

add(P6, "CNN Mini-Project",
    "End-to-end image classification project with production best practices",
    "Full project: data → augmentation → model → train → evaluate → export",
    "Complete image classification project (>90% accuracy)", 4)

add(P6, "DL Foundations – Revision Day 1",
    "Review MLP, backprop, optimizers, regularization",
    "Re-derive backprop; re-implement Adam optimizer",
    "Revised implementations + concept map", 3)

add(P6, "DL Foundations – Revision Day 2",
    "Review CNN architectures, transfer learning",
    "Implement a new CNN variant; practice on unseen dataset",
    "New CNN experiment + results analysis", 3)

# ────────────────────────────────────────────────────────────────
# PHASE 7  –  Sequence Models & Transformers  (Days 163-192)
# ────────────────────────────────────────────────────────────────
P7 = "Phase 7: Sequence Models & Transformers"

add(P7, "RNN – Theory",
    "Recurrent connections, hidden state, BPTT, vanishing/exploding gradients",
    "Derive BPTT equations; understand gradient flow in RNNs",
    "RNN theory notes + gradient flow diagrams", 3)

add(P7, "RNN – From Scratch",
    "Implement RNN cell, forward pass, BPTT in NumPy",
    "Implement RNN from scratch; train on character-level text generation",
    "rnn_from_scratch.py generating text", 4)

add(P7, "RNN – PyTorch",
    "nn.RNN, packed sequences, bidirectional RNN, many-to-one/many-to-many",
    "Build RNN classifier in PyTorch for sentiment analysis",
    "rnn_pytorch.py with sentiment classification", 3)

add(P7, "LSTM – Theory & From Scratch",
    "Forget/input/output gates, cell state, gradient highway",
    "Implement LSTM cell from scratch in NumPy",
    "lstm_from_scratch.py with gate visualization", 4)

add(P7, "LSTM – PyTorch Applications",
    "nn.LSTM, stacked LSTMs, bidirectional LSTM",
    "Build bidirectional LSTM for sequence classification",
    "lstm_pytorch.py with bidirectional comparison", 3)

add(P7, "GRU",
    "Reset/update gates, GRU vs LSTM, computational efficiency",
    "Implement GRU from scratch; compare with LSTM on same task",
    "gru.py with LSTM vs GRU comparison", 3)

add(P7, "Seq2Seq Models",
    "Encoder-decoder architecture, teacher forcing, beam search",
    "Implement seq2seq for simple translation or date conversion",
    "seq2seq.py with encoder-decoder training", 4)

add(P7, "Attention Mechanism – Theory",
    "Additive vs multiplicative attention, alignment scores, context vector",
    "Derive attention equations; implement Bahdanau attention from scratch",
    "attention_theory.py with attention weight visualization", 3)

add(P7, "Seq2Seq with Attention",
    "Adding attention to seq2seq, attention weights visualization",
    "Add attention to seq2seq model; visualize alignment",
    "seq2seq_attention.py with attention heatmaps", 4)

add(P7, "Self-Attention – From Scratch",
    "Query, Key, Value, scaled dot-product attention, self-attention layer",
    "Implement self-attention mechanism from scratch in NumPy + PyTorch",
    "self_attention.py with step-by-step computation", 3)

add(P7, "Multi-Head Attention",
    "Multiple attention heads, concatenation, linear projection",
    "Implement multi-head attention from scratch",
    "multi_head_attention.py with head visualization", 3)

add(P7, "Positional Encoding",
    "Sinusoidal encoding, learned encoding, why position matters",
    "Implement sinusoidal positional encoding; visualize patterns",
    "positional_encoding.py with encoding visualization", 3)

add(P7, "Transformer – Encoder Block",
    "Self-attention + Feed-forward + LayerNorm + residual connections",
    "Implement transformer encoder block from scratch",
    "transformer_encoder.py matching PyTorch nn.TransformerEncoderLayer", 4)

add(P7, "Transformer – Decoder Block",
    "Masked self-attention, cross-attention, autoregressive generation",
    "Implement transformer decoder block from scratch",
    "transformer_decoder.py with causal masking", 4)

add(P7, "Transformer – Full Architecture",
    "Putting it all together: encoder-decoder Transformer",
    "Build complete Transformer from scratch; train on toy translation",
    "transformer_full.py with training on toy task", 5)

add(P7, "Tokenization",
    "BPE, WordPiece, SentencePiece, Unigram; HuggingFace tokenizers",
    "Train BPE tokenizer on custom corpus; use HuggingFace tokenizers",
    "tokenization.py with custom tokenizer", 3)

add(P7, "BERT – Understanding",
    "Masked language model, next sentence prediction, pre-training, [CLS] token",
    "Study BERT architecture; load pre-trained BERT; extract embeddings",
    "bert_understanding.py with embedding extraction", 3)

add(P7, "BERT – Fine-Tuning",
    "Fine-tuning for classification, NER, QA with HuggingFace Transformers",
    "Fine-tune BERT for text classification using HuggingFace Trainer",
    "bert_finetuning.py with classification results", 3)

add(P7, "GPT – Understanding",
    "Decoder-only, autoregressive, causal attention, language modeling",
    "Study GPT architecture; use HuggingFace for text generation",
    "gpt_understanding.py with text generation examples", 3)

add(P7, "Vision Transformer (ViT)",
    "Image patches as tokens, position embeddings for images, cls token",
    "Implement simple ViT from scratch; compare with CNN on CIFAR-10",
    "vit.py with CNN vs ViT comparison", 4)

add(P7, "DL for Time Series – RNN/LSTM",
    "Time series forecasting with RNN/LSTM, sequence windowing",
    "Build LSTM forecaster; compare with ARIMA on same data",
    "ts_lstm.py with ARIMA vs LSTM comparison", 3)

add(P7, "DL for Time Series – Transformer",
    "Temporal attention, time series Transformer, informer concepts",
    "Build Transformer-based time series forecaster",
    "ts_transformer.py with multi-step forecast", 4)

add(P7, "NLP Mini-Project",
    "End-to-end NLP project using Transformer models",
    "Full project: data → tokenize → fine-tune BERT → evaluate → serve",
    "Complete NLP project with fine-tuned model", 4)

add(P7, "Sequence Models – Revision Day 1",
    "Review RNN, LSTM, GRU, Seq2Seq, Attention",
    "Re-implement attention mechanism; practice on new sequence task",
    "Revised implementations + concept map", 3)

add(P7, "Sequence Models – Revision Day 2",
    "Review Transformers, BERT, GPT, ViT",
    "Implement a component from scratch; test on new data",
    "New experiment + analysis", 3)

# ────────────────────────────────────────────────────────────────
# PHASE 8  –  Generative Models  (Days 193-212)
# ────────────────────────────────────────────────────────────────
P8 = "Phase 8: Generative Models"

add(P8, "Autoencoder – Theory & Vanilla AE",
    "Encoder-decoder, bottleneck, reconstruction loss, latent space",
    "Implement vanilla autoencoder in PyTorch; train on MNIST",
    "autoencoder.py with reconstruction visualization", 3)

add(P8, "Denoising & Sparse Autoencoders",
    "Denoising AE (noise injection), sparse AE (sparsity constraint)",
    "Implement denoising AE; add KL sparsity penalty for sparse AE",
    "dae_sparse_ae.py with denoising demo + feature visualization", 3)

add(P8, "Variational Autoencoder – Theory",
    "Latent distributions, reparameterization trick, ELBO, KL divergence loss",
    "Derive VAE loss function; understand reparameterization trick",
    "VAE theory notes + derivation", 3)

add(P8, "VAE – From Scratch",
    "Implement VAE in PyTorch: encoder, decoder, sampling, combined loss",
    "Implement VAE; train on MNIST; explore latent space",
    "vae.py with latent space interpolation visualization", 4)

add(P8, "GAN – Theory",
    "Generator vs discriminator, minimax game, Nash equilibrium, training dynamics",
    "Derive GAN objective; understand mode collapse, training instability",
    "GAN theory notes + training dynamics visualization", 3)

add(P8, "GAN – From Scratch",
    "Implement GAN in PyTorch: generator, discriminator, training loop",
    "Build simple GAN for MNIST digit generation",
    "gan.py generating recognizable digits", 4)

add(P8, "DCGAN",
    "Deep convolutional GAN, architectural guidelines, batch norm in GANs",
    "Implement DCGAN; train on CIFAR-10 or CelebA",
    "dcgan.py with generated image grid", 4)

add(P8, "WGAN & Training Stability",
    "Wasserstein distance, gradient penalty, spectral normalization",
    "Implement WGAN-GP; compare training stability with vanilla GAN",
    "wgan_gp.py with FID score comparison", 3)

add(P8, "Conditional GAN & Pix2Pix",
    "Class-conditional generation, paired image translation",
    "Implement conditional GAN; explore pix2pix concepts",
    "cgan.py with conditional generation demo", 3)

add(P8, "Diffusion Models – Theory",
    "Forward diffusion, reverse process, DDPM, noise schedule, score matching",
    "Study DDPM paper; understand forward/reverse process math",
    "Diffusion theory notes + noise schedule visualization", 3)

add(P8, "Diffusion Models – Implementation",
    "U-Net backbone, noise prediction, sampling, DDPM training loop",
    "Implement simplified DDPM for MNIST generation",
    "ddpm.py with generation process visualization", 4)

add(P8, "Neural Style Transfer",
    "Content/style loss, Gram matrix, VGG features, optimization-based NST",
    "Implement neural style transfer using pre-trained VGG",
    "neural_style_transfer.py with stylized images", 3)

add(P8, "Generative Models Mini-Project",
    "End-to-end generative model project",
    "Train best generative model on custom dataset; generate + evaluate",
    "Complete generative project with quality analysis", 4)

add(P8, "Phase 8 Revision",
    "Review AE, VAE, GAN, Diffusion concepts",
    "Re-implement weakest model; compare generation quality",
    "Revised implementations + model comparison", 3)

# ────────────────────────────────────────────────────────────────
# PHASE 9  –  Production Deep Learning  (Days 213-228)
# ────────────────────────────────────────────────────────────────
P9 = "Phase 9: Production Deep Learning"

add(P9, "ONNX Export & Inference",
    "ONNX format, torch.onnx.export, ONNX Runtime inference, model optimization",
    "Export PyTorch model to ONNX; benchmark inference speed",
    "onnx_export.py with speed comparison", 3)

add(P9, "Model Quantization",
    "Post-training quantization, dynamic/static quantization, INT8 inference",
    "Quantize a trained model; measure accuracy vs speed tradeoff",
    "quantization.py with accuracy/speed analysis", 3)

add(P9, "Model Pruning",
    "Unstructured/structured pruning, magnitude-based, lottery ticket hypothesis",
    "Prune a CNN; analyze sparsity vs accuracy tradeoff",
    "pruning.py with sparsity-accuracy curve", 3)

add(P9, "Knowledge Distillation",
    "Teacher-student framework, soft targets, temperature scaling",
    "Distill a large model into a smaller one; compare performance",
    "distillation.py with teacher vs student comparison", 3)

add(P9, "TorchServe / Triton",
    "Model serving at scale, TorchServe handlers, batching, model store",
    "Deploy model with TorchServe; benchmark throughput",
    "torchserve/ deployment with handler + config", 3)

add(P9, "Distributed Training – DDP",
    "DataParallel vs DistributedDataParallel, process groups, gradient sync",
    "Set up DDP training script; understand multi-GPU patterns",
    "ddp_training.py with multi-GPU support", 3)

add(P9, "Mixed Precision Training",
    "FP16/BF16, GradScaler, torch.cuda.amp, memory savings",
    "Add mixed precision to training loop; measure speedup",
    "mixed_precision.py with memory/speed comparison", 3)

add(P9, "Weights & Biases Integration",
    "W&B experiment tracking, sweeps, artifacts, model registry",
    "Set up W&B tracking for DL experiment; run hyperparameter sweep",
    "wandb_tracking.py with sweep configuration", 3)

add(P9, "DL Monitoring in Production",
    "Input validation, output monitoring, performance tracking, alerting",
    "Build DL inference pipeline with monitoring hooks",
    "dl_monitoring.py with anomaly detection on inputs/outputs", 3)

add(P9, "End-to-End DL Project – Part 1",
    "Problem definition, data collection, EDA, model architecture design",
    "Start capstone: define problem + prepare data + design architecture",
    "Project plan + data pipeline + architecture diagram", 4)

add(P9, "End-to-End DL Project – Part 2",
    "Training, evaluation, hyperparameter tuning, model analysis",
    "Train + tune + evaluate model; analyze errors",
    "Trained model + evaluation report + error analysis", 4)

add(P9, "End-to-End DL Project – Part 3",
    "Optimization, deployment, monitoring, documentation",
    "Quantize → export → deploy → monitor → document",
    "Production-deployed model + documentation + monitoring", 4)

add(P9, "Siamese Networks",
    "Contrastive loss, triplet loss, one-shot learning, similarity learning",
    "Implement Siamese network for face/signature verification",
    "siamese_net.py with similarity matching demo", 3)

add(P9, "Graph Neural Networks – Intro",
    "Graph convolution, message passing, GCN, node/graph classification",
    "Implement basic GCN with PyTorch Geometric; classify nodes",
    "gnn_intro.py with node classification on Cora", 3)

add(P9, "Final Revision & Portfolio Day 1",
    "Review entire curriculum; identify gaps; polish portfolio",
    "Fix any remaining weak areas; update all README files",
    "Polished portfolio repo", 3)

add(P9, "Final Revision & Portfolio Day 2",
    "Finalize portfolio; create learning summary; plan next steps",
    "Write comprehensive README; create skills summary",
    "Complete ML/DL portfolio with all projects", 3)


# ── Build the workbook ──────────────────────────────────────────
wb = openpyxl.Workbook()

# ────── Sheet 1: Day-by-Day Plan ────────────────────────────────
ws = wb.active
ws.title = "Day-by-Day Plan"

headers = [
    "Day #", "Week #", "Phase", "Topic",
    "Learning Objectives", "Implementation Task",
    "Deliverable", "Est. Hours", "Status", "Notes"
]

col_widths = [7, 8, 32, 38, 52, 52, 44, 10, 12, 30]

for col_idx, (hdr, w) in enumerate(zip(headers, col_widths), 1):
    cell = ws.cell(row=1, column=col_idx, value=hdr)
    cell.font = HEADER_FONT
    cell.fill = HEADER_FILL
    cell.alignment = CENTER_ALIGN
    cell.border = THIN_BORDER
    ws.column_dimensions[get_column_letter(col_idx)].width = w

# Status dropdown
status_dv = DataValidation(
    type="list",
    formula1='"Not Started,In Progress,Completed,Skipped"',
    allow_blank=True,
)
status_dv.error = "Pick a valid status"
status_dv.errorTitle = "Invalid Status"
ws.add_data_validation(status_dv)

current_phase = None
for day_idx, (phase, topic, learn, impl, deliv, hrs) in enumerate(PLAN, start=1):
    row = day_idx + 1
    week = (day_idx - 1) // 7 + 1

    ws.cell(row=row, column=1, value=day_idx).alignment = CENTER_ALIGN
    ws.cell(row=row, column=2, value=week).alignment = CENTER_ALIGN
    ws.cell(row=row, column=3, value=phase)
    ws.cell(row=row, column=4, value=topic)
    ws.cell(row=row, column=5, value=learn)
    ws.cell(row=row, column=6, value=impl)
    ws.cell(row=row, column=7, value=deliv)
    ws.cell(row=row, column=8, value=hrs).alignment = CENTER_ALIGN
    status_cell = ws.cell(row=row, column=9, value="Not Started")
    status_cell.alignment = CENTER_ALIGN
    ws.cell(row=row, column=10, value="")

    status_dv.add(status_cell)

    # Phase colour band
    phase_color = PHASE_COLORS.get(phase, "2C3E50")
    phase_fill = PatternFill("solid", fgColor=phase_color)

    for col in range(1, 11):
        c = ws.cell(row=row, column=col)
        c.border = THIN_BORDER
        c.font = BODY_FONT
        if col >= 5:
            c.alignment = WRAP_ALIGN

    # Colour the phase column
    ws.cell(row=row, column=3).fill = phase_fill
    ws.cell(row=row, column=3).font = PHASE_FONT

ws.auto_filter.ref = f"A1:J{len(PLAN)+1}"
ws.freeze_panes = "A2"
ws.sheet_properties.tabColor = "2980B9"

# ────── Sheet 2: Phase Summary ──────────────────────────────────
ws2 = wb.create_sheet("Phase Summary")

summary_headers = ["Phase", "Days", "Weeks (approx)", "Total Hours", "Key Focus Areas"]
sum_col_widths = [36, 10, 15, 12, 70]

for col_idx, (hdr, w) in enumerate(zip(summary_headers, sum_col_widths), 1):
    cell = ws2.cell(row=1, column=col_idx, value=hdr)
    cell.font = HEADER_FONT
    cell.fill = HEADER_FILL
    cell.alignment = CENTER_ALIGN
    cell.border = THIN_BORDER
    ws2.column_dimensions[get_column_letter(col_idx)].width = w

phase_data: dict[str, dict] = {}
for phase, topic, learn, impl, deliv, hrs in PLAN:
    if phase not in phase_data:
        phase_data[phase] = {"count": 0, "hours": 0.0, "topics": []}
    phase_data[phase]["count"] += 1
    phase_data[phase]["hours"] += hrs
    phase_data[phase]["topics"].append(topic.split("–")[0].split("—")[0].strip())

focus_areas = {
    "Phase 0: Prerequisites": "Python, NumPy, linear algebra, calculus, probability, statistics, Pandas, visualization",
    "Phase 1: Core Classical ML": "Linear/logistic regression, regularization, metrics, Naive Bayes, KNN, cross-validation, pipelines",
    "Phase 2: Trees, Ensembles & SVMs": "Decision trees, Random Forest, AdaBoost, Gradient Boosting, XGBoost/LightGBM, SVM, hyperparameter tuning",
    "Phase 3: Unsupervised Learning": "K-Means, DBSCAN, hierarchical, GMM, PCA, t-SNE, UMAP, anomaly detection",
    "Phase 4: Production ML Engineering": "Feature engineering, model serving (FastAPI), Docker, MLflow, DVC, CI/CD, monitoring",
    "Phase 5: Time Series Analysis": "ARIMA, SARIMA, Holt-Winters, ETS, GARCH, VAR, Prophet, ML for TS, anomaly detection in TS",
    "Phase 6: Deep Learning Foundations": "MLP, backprop, optimizers, PyTorch, CNN, ResNet, transfer learning, Grad-CAM",
    "Phase 7: Sequence Models & Transformers": "RNN, LSTM, GRU, Seq2Seq, attention, Transformer, BERT, GPT, ViT, DL for time series",
    "Phase 8: Generative Models": "Autoencoders, VAE, GAN, DCGAN, WGAN, diffusion models, neural style transfer",
    "Phase 9: Production Deep Learning": "ONNX, quantization, pruning, distillation, TorchServe, DDP, W&B, Siamese, GNN",
}

cumulative_days = 0
for row_idx, (phase, data) in enumerate(phase_data.items(), start=2):
    day_start = cumulative_days + 1
    cumulative_days += data["count"]

    ws2.cell(row=row_idx, column=1, value=phase)
    ws2.cell(row=row_idx, column=2, value=f"{day_start}-{cumulative_days}")
    ws2.cell(row=row_idx, column=3, value=round(data["count"] / 7, 1))
    ws2.cell(row=row_idx, column=4, value=data["hours"])
    ws2.cell(row=row_idx, column=5, value=focus_areas.get(phase, ""))

    phase_color = PHASE_COLORS.get(phase, "2C3E50")
    for col in range(1, 6):
        c = ws2.cell(row=row_idx, column=col)
        c.border = THIN_BORDER
        c.font = BODY_FONT
        c.alignment = WRAP_ALIGN if col == 5 else CENTER_ALIGN
    ws2.cell(row=row_idx, column=1).fill = PatternFill("solid", fgColor=phase_color)
    ws2.cell(row=row_idx, column=1).font = PHASE_FONT

# Totals row
total_row = len(phase_data) + 2
ws2.cell(row=total_row, column=1, value="TOTAL").font = Font(bold=True, size=11, name="Calibri")
ws2.cell(row=total_row, column=2, value=f"1-{cumulative_days}").font = Font(bold=True, size=11, name="Calibri")
ws2.cell(row=total_row, column=3, value=round(cumulative_days / 7, 1)).font = Font(bold=True, size=11, name="Calibri")
total_hours = sum(d["hours"] for d in phase_data.values())
ws2.cell(row=total_row, column=4, value=total_hours).font = Font(bold=True, size=11, name="Calibri")
for col in range(1, 6):
    ws2.cell(row=total_row, column=col).border = THIN_BORDER
    ws2.cell(row=total_row, column=col).alignment = CENTER_ALIGN

ws2.sheet_properties.tabColor = "27AE60"

# ────── Sheet 3: Algorithm Checklist ────────────────────────────
ws3 = wb.create_sheet("Algorithm Checklist")

algo_headers = ["#", "Category", "Algorithm", "From Scratch?", "Library Used?", "Tested?", "Notes"]
algo_widths = [5, 28, 40, 14, 14, 10, 30]

for col_idx, (hdr, w) in enumerate(zip(algo_headers, algo_widths), 1):
    cell = ws3.cell(row=1, column=col_idx, value=hdr)
    cell.font = HEADER_FONT
    cell.fill = HEADER_FILL
    cell.alignment = CENTER_ALIGN
    cell.border = THIN_BORDER
    ws3.column_dimensions[get_column_letter(col_idx)].width = w

algorithms = [
    # Regression
    ("Supervised – Regression", "Linear Regression (OLS + GD)"),
    ("Supervised – Regression", "Polynomial Regression"),
    ("Supervised – Regression", "Ridge Regression (L2)"),
    ("Supervised – Regression", "Lasso Regression (L1)"),
    ("Supervised – Regression", "Elastic Net"),
    ("Supervised – Regression", "Bayesian Linear Regression"),
    ("Supervised – Regression", "SVR"),
    ("Supervised – Regression", "Decision Tree Regressor"),
    ("Supervised – Regression", "Random Forest Regressor"),
    ("Supervised – Regression", "Gradient Boosted Regressor"),
    ("Supervised – Regression", "XGBoost Regressor"),
    ("Supervised – Regression", "LightGBM Regressor"),
    ("Supervised – Regression", "CatBoost Regressor"),
    ("Supervised – Regression", "KNN Regressor"),
    # Classification
    ("Supervised – Classification", "Logistic Regression"),
    ("Supervised – Classification", "Perceptron"),
    ("Supervised – Classification", "Naive Bayes (Gaussian/Multinomial/Bernoulli)"),
    ("Supervised – Classification", "KNN Classifier"),
    ("Supervised – Classification", "SVM (Linear)"),
    ("Supervised – Classification", "SVM (Kernel – RBF)"),
    ("Supervised – Classification", "Decision Tree Classifier"),
    ("Supervised – Classification", "Random Forest Classifier"),
    ("Supervised – Classification", "AdaBoost"),
    ("Supervised – Classification", "Gradient Boosted Classifier"),
    ("Supervised – Classification", "XGBoost Classifier"),
    ("Supervised – Classification", "LightGBM Classifier"),
    ("Supervised – Classification", "CatBoost Classifier"),
    ("Supervised – Classification", "Stacking Ensemble"),
    ("Supervised – Classification", "Voting Classifier"),
    ("Supervised – Classification", "LDA / QDA"),
    # Clustering
    ("Unsupervised – Clustering", "K-Means"),
    ("Unsupervised – Clustering", "K-Means++ / Mini-Batch K-Means"),
    ("Unsupervised – Clustering", "DBSCAN"),
    ("Unsupervised – Clustering", "HDBSCAN"),
    ("Unsupervised – Clustering", "Agglomerative Hierarchical"),
    ("Unsupervised – Clustering", "GMM (EM Algorithm)"),
    ("Unsupervised – Clustering", "Mean Shift"),
    ("Unsupervised – Clustering", "Spectral Clustering"),
    ("Unsupervised – Clustering", "OPTICS"),
    # Dim Reduction
    ("Unsupervised – Dim Reduction", "PCA"),
    ("Unsupervised – Dim Reduction", "Kernel PCA"),
    ("Unsupervised – Dim Reduction", "Truncated SVD"),
    ("Unsupervised – Dim Reduction", "t-SNE"),
    ("Unsupervised – Dim Reduction", "UMAP"),
    ("Unsupervised – Dim Reduction", "ICA"),
    ("Unsupervised – Dim Reduction", "NMF"),
    ("Unsupervised – Dim Reduction", "Factor Analysis"),
    # Anomaly
    ("Anomaly Detection", "Isolation Forest"),
    ("Anomaly Detection", "One-Class SVM"),
    ("Anomaly Detection", "Local Outlier Factor"),
    ("Anomaly Detection", "Elliptic Envelope"),
    # Time Series
    ("Time Series", "AR Model"),
    ("Time Series", "MA Model"),
    ("Time Series", "ARMA"),
    ("Time Series", "ARIMA"),
    ("Time Series", "SARIMA"),
    ("Time Series", "Auto ARIMA"),
    ("Time Series", "Simple Exponential Smoothing"),
    ("Time Series", "Holt's Linear (DES)"),
    ("Time Series", "Holt-Winters (TES)"),
    ("Time Series", "ETS Framework"),
    ("Time Series", "GARCH"),
    ("Time Series", "VAR"),
    ("Time Series", "VECM"),
    ("Time Series", "Prophet"),
    ("Time Series", "ML-based (XGBoost/LightGBM)"),
    ("Time Series", "LSTM for Time Series"),
    ("Time Series", "Transformer for Time Series"),
    # Deep Learning
    ("Deep Learning", "Perceptron"),
    ("Deep Learning", "MLP (from scratch + PyTorch)"),
    ("Deep Learning", "CNN (from scratch + PyTorch)"),
    ("Deep Learning", "LeNet / AlexNet / VGG"),
    ("Deep Learning", "ResNet"),
    ("Deep Learning", "Transfer Learning"),
    ("Deep Learning", "RNN"),
    ("Deep Learning", "LSTM"),
    ("Deep Learning", "GRU"),
    ("Deep Learning", "Seq2Seq"),
    ("Deep Learning", "Seq2Seq + Attention"),
    ("Deep Learning", "Transformer (from scratch)"),
    ("Deep Learning", "BERT (fine-tune)"),
    ("Deep Learning", "GPT (understand)"),
    ("Deep Learning", "Vision Transformer (ViT)"),
    ("Deep Learning", "Autoencoder"),
    ("Deep Learning", "VAE"),
    ("Deep Learning", "GAN"),
    ("Deep Learning", "DCGAN / WGAN"),
    ("Deep Learning", "Diffusion Model (DDPM)"),
    ("Deep Learning", "Siamese Network"),
    ("Deep Learning", "GNN (GCN)"),
]

yn_dv = DataValidation(
    type="list", formula1='"Yes,No,Partial"', allow_blank=True
)
ws3.add_data_validation(yn_dv)

for i, (cat, algo) in enumerate(algorithms, start=1):
    row = i + 1
    ws3.cell(row=row, column=1, value=i).alignment = CENTER_ALIGN
    ws3.cell(row=row, column=2, value=cat)
    ws3.cell(row=row, column=3, value=algo)
    for col in (4, 5, 6):
        c = ws3.cell(row=row, column=col, value="No")
        c.alignment = CENTER_ALIGN
        yn_dv.add(c)
    ws3.cell(row=row, column=7, value="")
    for col in range(1, 8):
        ws3.cell(row=row, column=col).border = THIN_BORDER
        ws3.cell(row=row, column=col).font = BODY_FONT

ws3.auto_filter.ref = f"A1:G{len(algorithms)+1}"
ws3.freeze_panes = "A2"
ws3.sheet_properties.tabColor = "E74C3C"

# ────── Sheet 4: Weekly Milestones ──────────────────────────────
ws4 = wb.create_sheet("Weekly Milestones")

mile_headers = ["Week #", "Phase", "Major Milestone", "Self-Test", "Achieved?"]
mile_widths = [8, 32, 55, 55, 12]

for col_idx, (hdr, w) in enumerate(zip(mile_headers, mile_widths), 1):
    cell = ws4.cell(row=1, column=col_idx, value=hdr)
    cell.font = HEADER_FONT
    cell.fill = HEADER_FILL
    cell.alignment = CENTER_ALIGN
    cell.border = THIN_BORDER
    ws4.column_dimensions[get_column_letter(col_idx)].width = w

milestones = [
    (1, P0, "Python OOP + NumPy mastery", "Can you implement matrix multiply without for-loops?"),
    (2, P0, "Linear Algebra + Calculus solid", "Can you derive gradient descent from first principles?"),
    (3, P0, "Probability, Stats, Pandas, Visualization complete", "Can you perform full EDA on a new dataset in under 1 hour?"),
    (4, P1, "Linear & Logistic Regression from scratch", "Can you explain the normal equation and derive log-loss gradient?"),
    (5, P1, "Regularization + evaluation metrics mastery", "Can you explain when to use L1 vs L2 and why?"),
    (6, P1, "Naive Bayes, KNN, feature engineering, pipelines", "Can you build an end-to-end sklearn Pipeline from memory?"),
    (7, P2, "Decision Trees from scratch", "Can you implement a decision tree that handles both Gini and entropy?"),
    (8, P2, "Random Forest + AdaBoost from scratch", "Can you explain bagging vs boosting with code?"),
    (9, P2, "Gradient Boosting from scratch", "Can you derive the GBM update rule for classification?"),
    (10, P2, "XGBoost/LightGBM + SVM mastery", "Can you tune XGBoost with Optuna and explain every key parameter?"),
    (11, P3, "K-Means, DBSCAN, hierarchical from scratch", "Can you choose the right clustering algorithm for a given data shape?"),
    (12, P3, "GMM/EM + PCA from scratch", "Can you derive the EM algorithm for GMM?"),
    (13, P3, "t-SNE/UMAP + anomaly detection", "Can you build an anomaly detection pipeline for production?"),
    (14, P4, "Feature engineering + model serving", "Can you deploy a model as a REST API with FastAPI + Docker?"),
    (15, P4, "MLflow + DVC + CI/CD complete", "Can you set up reproducible ML experiments with versioning?"),
    (16, P5, "Stationarity, ACF/PACF, AR, MA, ARMA", "Can you identify AR/MA orders from ACF/PACF plots?"),
    (17, P5, "ARIMA + SARIMA mastery", "Can you perform full Box-Jenkins methodology?"),
    (18, P5, "Exponential Smoothing + ETS + GARCH", "Can you choose between ARIMA and ETS for a given series?"),
    (19, P5, "VAR, Prophet, ML for time series", "Can you forecast multivariate time series?"),
    (20, P5, "TS cross-validation + anomaly detection + projects", "Can you build a production-grade forecasting system?"),
    (21, P6, "MLP + backprop from scratch", "Can you derive backprop for a 3-layer network on a whiteboard?"),
    (22, P6, "Optimizers + regularization + PyTorch basics", "Can you implement Adam from scratch and explain each term?"),
    (23, P6, "CNN from scratch + PyTorch", "Can you explain every operation in a conv layer including output shape?"),
    (24, P6, "ResNet + transfer learning + CNN project", "Can you fine-tune a pre-trained model on a new dataset end-to-end?"),
    (25, P7, "RNN, LSTM, GRU from scratch", "Can you implement LSTM with gate equations from memory?"),
    (26, P7, "Seq2Seq + Attention mechanism", "Can you explain and implement Bahdanau attention?"),
    (27, P7, "Transformer from scratch", "Can you implement multi-head self-attention from scratch?"),
    (28, P7, "BERT/GPT fine-tuning + ViT + DL for TS", "Can you fine-tune BERT and explain the architecture?"),
    (29, P8, "Autoencoders + VAE", "Can you derive the ELBO and reparameterization trick?"),
    (30, P8, "GAN + DCGAN + WGAN", "Can you train a stable GAN and explain mode collapse?"),
    (31, P8, "Diffusion models + generative project", "Can you explain the forward/reverse diffusion process?"),
    (32, P9, "ONNX + quantization + pruning + distillation", "Can you optimize a model for production inference?"),
    (33, P9, "DDP + serving + monitoring + capstone", "Can you deploy and monitor a DL model in production?"),
]

achieved_dv = DataValidation(
    type="list", formula1='"Yes,No,Partial"', allow_blank=True
)
ws4.add_data_validation(achieved_dv)

for i, (week, phase, milestone, self_test) in enumerate(milestones, start=2):
    ws4.cell(row=i, column=1, value=week).alignment = CENTER_ALIGN
    ws4.cell(row=i, column=2, value=phase)
    ws4.cell(row=i, column=3, value=milestone)
    ws4.cell(row=i, column=4, value=self_test)
    ac = ws4.cell(row=i, column=5, value="No")
    ac.alignment = CENTER_ALIGN
    achieved_dv.add(ac)
    phase_color = PHASE_COLORS.get(phase, "2C3E50")
    ws4.cell(row=i, column=2).fill = PatternFill("solid", fgColor=phase_color)
    ws4.cell(row=i, column=2).font = PHASE_FONT
    for col in range(1, 6):
        ws4.cell(row=i, column=col).border = THIN_BORDER
        ws4.cell(row=i, column=col).font = BODY_FONT
        if col in (3, 4):
            ws4.cell(row=i, column=col).alignment = WRAP_ALIGN

ws4.freeze_panes = "A2"
ws4.sheet_properties.tabColor = "F39C12"

# ────── Save ────────────────────────────────────────────────────
output_path = "/home/pal/Projects/MachineLearning/ML_DL_Learning_Plan.xlsx"
wb.save(output_path)
print(f"✅ Workbook saved to: {output_path}")
print(f"   Total learning days: {len(PLAN)}")
print(f"   Total hours: {sum(r[5] for r in PLAN)}")
print(f"   Sheets: {wb.sheetnames}")
