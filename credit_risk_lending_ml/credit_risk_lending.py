import numpy as np
import pandas as pd
from sklearn.model_selection import train_test_split
from sklearn.preprocessing import StandardScaler
from sklearn.tree import DecisionTreeClassifier
from sklearn.linear_model import LogisticRegression
from sklearn.metrics import ( accuracy_score, confusion_matrix, f1_score,
    precision_score,recall_score,roc_auc_score)
from sklearn.ensemble import IsolationForest
import matplotlib.pyplot as plt
from sklearn.cluster import KMeans
from sklearn.metrics import calinski_harabasz_score

df = pd.read_csv("credit_applicants.csv")

# Calculate metrics
total_applicants = len(df)
total_defaults = df["default"].sum()
default_rate = df["default"].mean() * 100
missing_bureau_pct = df["credit_bureau_score"].isnull().mean() * 100
missing_bureau_count = df["credit_bureau_score"].isnull().sum()

# Engineer is_thin_file flag
df["is_thin_file"] = df["credit_bureau_score"].isnull().astype(int)

X = df.drop(columns=["applicant_id", "default"])
y = df["default"]

# Stratified 75/25 Train/Test Split
X_train, X_test, y_train, y_test = train_test_split(
    X, y, test_size=0.25, random_state=42, stratify=y)

train_size = len(X_train)
test_size = len(X_test)
train_defaults = y_train.sum()
test_defaults = y_test.sum()
train_default_pct = y_train.mean() * 100
test_default_pct = y_test.mean() * 100

# Training-derived Median Imputation for credit_bureau_score
train_bureau_median = X_train["credit_bureau_score"].median()

X_train["credit_bureau_score"] = X_train["credit_bureau_score"].fillna(
    train_bureau_median)
X_test["credit_bureau_score"] = X_test["credit_bureau_score"].fillna(
    train_bureau_median)

# One-Hot Encoding for employment_type
X_train = pd.get_dummies(X_train, columns=["employment_type"], drop_first=True, dtype=int)
X_test = pd.get_dummies(X_test, columns=["employment_type"], drop_first=True, dtype=int)

# Feature Scaling with StandardScaler
scaler = StandardScaler()
numeric_cols = [
    "age",
    "monthly_income_inr",
    "existing_loans_count",
    "credit_utilization_ratio",
    "upi_monthly_inflow_inr",
    "bounced_payments_count",
    "credit_bureau_score",
]

X_train[numeric_cols] = scaler.fit_transform(X_train[numeric_cols])
X_test[numeric_cols] = scaler.transform(X_test[numeric_cols])

print("----------------------------------PART A---------------------------------")
report = f""" --- EDA AND PREPROCESSING SUMMARY REPORT ---
================================================================
Total Applicant Records            : {total_applicants}
Exact Measured Default Rate        : {default_rate:.2f}% ({total_defaults} / {total_applicants} applicants)
Missing Credit Bureau Score Rate   : {missing_bureau_pct:.2f}% ({missing_bureau_count} / {total_applicants} applicants)
================================================================
Train/Test Split Ratio             : 75 / 25 (Stratified, random_state=42)
Training Set Size                  : {train_size} applicants ({train_defaults} defaults | {train_default_pct:.2f}%)
Testing Set Size                   : {test_size} applicants ({test_defaults} defaults | {test_default_pct:.2f}%)
================================================================
Training-Derived Bureau Median     : {train_bureau_median:.2f}
Categorical Encoding Strategy      : One-Hot Encoding (drop_first=True)
Numerical Scaling Strategy         : StandardScaler (fit on train only)
================================================================
"""

print(report)

# Train Logistic Regression
log_reg = LogisticRegression(random_state=42)
log_reg.fit(X_train, y_train)

# Train Decision Tree Classifier
tree_clf = DecisionTreeClassifier(random_state=42)
tree_clf.fit(X_train, y_train)

# Predict on Test Set
y_pred_log = log_reg.predict(X_test)
y_prob_log = log_reg.predict_proba(X_test)[:, 1]

y_pred_tree = tree_clf.predict(X_test)
y_prob_tree = tree_clf.predict_proba(X_test)[:, 1]


# Computing evaluation metrics
def get_metrics(y_true, y_pred, y_prob):
    cm = confusion_matrix(y_true, y_pred)
    tn, fp, fn, tp = cm.ravel()
    return {
        "Confusion Matrix (TN, FP, FN, TP)": f"[[{tn}, {fp}], [{fn}, {tp}]]",
        "Accuracy": accuracy_score(y_true, y_pred),
        "Precision": precision_score(y_true, y_pred, zero_division=0),
        "Recall": recall_score(y_true, y_pred, zero_division=0),
        "F1-Score": f1_score(y_true, y_pred, zero_division=0),
        "ROC-AUC": roc_auc_score(y_true, y_prob),
    }

metrics_log = get_metrics(y_test, y_pred_log, y_prob_log)
metrics_tree = get_metrics(y_test, y_pred_tree, y_prob_tree)

comparison_df = pd.DataFrame([metrics_log, metrics_tree], index=["Logistic Regression", "Decision Tree Classifier"])

# Display Metrics Table
print("----------------------------------PART B---------------------------------")
print("=========================================================================")
print("             MODEL EVALUATION COMPARISON TABLE (TEST SET)                ")
print("=========================================================================")
print(comparison_df.to_string())
print("=========================================================================")

# Obtain predicted probabilities on test set
y_prob_test = log_reg.predict_proba(X_test)[:, 1]
test_results = pd.DataFrame({"actual_default": y_test.values, "predicted_prob": y_prob_test})

# Bucket applicants into 4 Risk Tiers (Quartiles of predicted default probability)
test_results["risk_tier"] = pd.qcut(
    test_results["predicted_prob"], q=4,
    labels=[ "Tier 1 (Low Risk)", "Tier 2 (Medium-Low Risk)",
        "Tier 3 (Medium-High Risk)", "Tier 4 (High Risk)",],)

# Aggregate metrics by Risk Tier
pricing_table = (test_results.groupby("risk_tier", observed=False).agg(
        total_applicants=("actual_default", "count"),
        observed_defaults=("actual_default", "sum"),
        min_predicted_prob=("predicted_prob", "min"),
        max_predicted_prob=("predicted_prob", "max"),
        observed_default_rate=("actual_default", lambda x: (x.sum() / len(x)) * 100,),).reset_index())

# Assign Illustrative Interest Rate Ranges (Lower Risk → Lower Rate)
interest_rate_map = {
    "Tier 1 (Low Risk)": "9.5% - 12.0%",
    "Tier 2 (Medium-Low Risk)": "13.0% - 16.5%",
    "Tier 3 (Medium-High Risk)": "17.0% - 21.0%",
    "Tier 4 (High Risk)": "22.0% - 28.0% (or Decline)"}

pricing_table["interest_rate_range"] = pricing_table["risk_tier"].map(interest_rate_map)

# Format Probability Ranges as Percentages
pricing_table["pred_prob_range"] = pricing_table.apply(lambda row: ( f"{row['min_predicted_prob']*100:.1f}% -"f" {row['max_predicted_prob']*100:.1f}%"), axis=1,)

# Reorder columns for final display
final_pricing_table = pricing_table[[
    "risk_tier",
    "total_applicants",
    "pred_prob_range",
    "observed_defaults",
    "observed_default_rate",
    "interest_rate_range",
]]

# Print Table
print("==========================================================================")
print("                   RISK-BASED PRICING TABLE (TEST SET)                    ")
print("==========================================================================")
print(final_pricing_table.to_string(index=False))
print("==========================================================================")

behaviour_df = pd.read_csv("txn_behaviour.csv")
features = ["txn_hour", "is_new_device", "txn_amount_inr"]
X = behaviour_df[features]

scaler = StandardScaler()
X_scaled = scaler.fit_transform(X)

contamination_rate = 15 / len(behaviour_df)

# Fit Isolation Forest
iso_forest = IsolationForest(random_state=42, contamination=contamination_rate)
behaviour_df["anomaly_pred"] = iso_forest.fit_predict(X_scaled)
behaviour_df["is_seeded_anomaly"] = behaviour_df["txn_id"].str.startswith("BTXNA")

# Calculation of recall against ground truth
seeded_anomalies = behaviour_df[behaviour_df["is_seeded_anomaly"]]
flagged_seeded_anomalies = seeded_anomalies[
    seeded_anomalies["anomaly_pred"] == -1]

detected_count = len(flagged_seeded_anomalies)
total_seeded = len(seeded_anomalies)
recall_pct = (detected_count / total_seeded) * 100
missed_seeded_anomalies = total_seeded - detected_count

print("----------------------------------PART C---------------------------------")
print("==========================================================================")
print("                   ISOLATION FOREST ANOMALY DETECTION REPORT                    ")
print("==========================================================================")
print(f"Total Behavioral Transactions Evaluated : {len(behaviour_df)}")
print(f"Total Injected Seeded Anomalies (BTXNA) : {total_seeded}")
print(f"Contamination Parameter                 : {contamination_rate:.6f} ({15} / {len(behaviour_df)})")
print("================================================================")
print(f"Flagged Seeded Anomalies                : {detected_count}")
print(f"Missed Seeded Anomalies                 : {missed_seeded_anomalies}")
print(f"Isolation Forest Anomaly Recall         : {recall_pct:.2f}%")

# Combine transformed train and test features back to cluster the full dataset
X_all = pd.concat([X_train, X_test], axis=0).sort_index()
y_all = pd.concat([y_train, y_test], axis=0).sort_index()

# Evaluate K using Calinski-Harabasz Index & Inertia (Elbow Method)
k_range = range(2, 7)
ch_scores = []
inertias = []

for k in k_range:
    kmeans = KMeans(n_clusters=k, random_state=42, n_init=10)
    labels = kmeans.fit_predict(X_all)
    inertias.append(kmeans.inertia_)
    ch_scores.append(calinski_harabasz_score(X_all, labels))

# Automatically select optimal k (highest Calinski-Harabasz score)
optimal_k = k_range[np.argmax(ch_scores)]

# Fit final K-Means model with optimal k
kmeans = KMeans(n_clusters=optimal_k, random_state=42, n_init=10)
df["cluster"] = kmeans.fit_predict(X_all)

# Summarize Cluster Characteristics & Default Over-Indexing
cluster_summary = (
    df.groupby("cluster").agg(
        total_applicants=("default", "count"),
        observed_defaults=("default", "sum"),
        default_rate_pct=("default", lambda x: (x.mean()) * 100),
        avg_income=("monthly_income_inr", "mean"),
        avg_utilization=("credit_utilization_ratio", "mean"),
        avg_bounced_payments=("bounced_payments_count", "mean"),
        thin_file_pct=("is_thin_file", lambda x: (x.mean()) * 100),).reset_index())

print("==========================================================================================")
print("                   K-MEANS CLUSTERING REPORT                                              ")
print("==========================================================================================")
print(f"Calinski-Harabasz Scores for k={list(k_range)}: {[f'{score:.2f}' for score in ch_scores]}")
print(f"Optimal K selected via Calinski-Harabasz Index: {optimal_k}")
print("==========================================================================================")
print(cluster_summary.to_string(index=False))
print("==========================================================================================")