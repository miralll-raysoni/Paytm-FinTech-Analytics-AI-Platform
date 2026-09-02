import pandas as pd
import os
import matplotlib.pyplot as plt

ledger_df = pd.read_csv("ledger.csv")
gateway_df = pd.read_csv("gateway_export.csv")

def reconcile_payments(ledger_df, gateway_df):

    # 1. Identify missing transactions using Set Operations on transaction_id
    ledger_ids = set(ledger_df["transaction_id"])
    gateway_ids = set(gateway_df["transaction_id"])

    missing_in_gateway_ids = ledger_ids - gateway_ids
    missing_in_ledger_ids = gateway_ids - ledger_ids

    missing_in_gateway = ledger_df[ledger_df["transaction_id"].isin(missing_in_gateway_ids)].copy()
    missing_in_ledger = gateway_df[gateway_df["transaction_id"].isin(missing_in_ledger_ids)].copy()

    # 2. Comparison for common transactions using pd.merge
    common_ids = ledger_ids.intersection(gateway_ids)

    common_ledger = ledger_df[ledger_df["transaction_id"].isin(common_ids)]
    common_gateway = gateway_df[gateway_df["transaction_id"].isin(common_ids)]

    merged = pd.merge(common_ledger, common_gateway, on="transaction_id", suffixes=("_ledger", "_gateway"),)

    # 3. Identify Amount Mismatches
    amount_mismatches = merged[
        merged["amount_inr_ledger"] != merged["amount_inr_gateway"]].copy()
    amount_mismatches["amount_difference"] = (amount_mismatches["amount_inr_gateway"]- amount_mismatches["amount_inr_ledger"])

    # 4. Identify Status Mismatches
    status_mismatches = merged[
        merged["status_ledger"] != merged["status_gateway"]].copy()

    return (missing_in_gateway, missing_in_ledger, amount_mismatches, status_mismatches)


if __name__ == "__main__":

    # Run reconciliation
    missing_gateway, missing_ledger, amount_diff, status_diff = reconcile_payments(ledger_df, gateway_df)

    # Financial At-Risk Calculations
    val_missing_gateway = missing_gateway["amount_inr"].sum()
    val_missing_ledger = missing_ledger["amount_inr"].sum()
    val_status_diff = status_diff["amount_inr_ledger"].sum()
    val_amount_diff_abs = amount_diff["amount_difference"].abs().sum()

    total_mismatches = (len(missing_gateway) + len(missing_ledger) + len(amount_diff) + len(status_diff))
    total_at_risk_inr = (val_missing_gateway + val_missing_ledger + val_status_diff + val_amount_diff_abs)

    # Report Discrepancy Counts
    print("--- PAYMENT RECONCILIATION SUMMARY REPORT ---")
    print("================================================================\n")
    print(f"Total Ledger Transactions: {len(ledger_df)}")
    print(f"Total Gateway Export Transactions: {len(gateway_df)}\n")

    print(f"1. Missing in Gateway Export     : {len(missing_gateway)} txns  | ₹{val_missing_gateway:,.2f}")
    print(f"2. Missing in Ledger             : {len(missing_ledger)} txns  | ₹{val_missing_ledger:,.2f}")
    print(f"3. Status Mismatches             : {len(status_diff)} txns   | ₹{val_status_diff:,.2f}")
    print(f"4. Amount Mismatches (Abs Delta) : {len(amount_diff)} txns  | ₹{val_amount_diff_abs:,.2f}")

    print("================================================================\n")

    print(f"TOTAL MISMATCH TRANSACTIONS     : {total_mismatches} txns")
    print(f"TOTAL CAPITAL AT-RISK           : ₹{total_at_risk_inr:,.2f}")
    print("================================================================\n")


# PART D: Four-layer analytics dashboard (code-generated, not a live BI tool)
os.makedirs("dashboard_output", exist_ok=True)

# Data Loading & Merging
ledger_df = pd.read_csv("ledger.csv")
merchants_df = pd.read_csv("merchants.csv")
gateway_df = pd.read_csv("gateway_export.csv")

ledger_merged = pd.merge(ledger_df, merchants_df, on="merchant_id", how="left")
ledger_merged["transaction_time"] = pd.to_datetime(
    ledger_merged["transaction_time"]
)
ledger_merged["date"] = ledger_merged["transaction_time"].dt.date

# Layer 1: Headline Scorecards
total_gmv = ledger_df["amount_inr"].sum()
overall_success_rate = (
    (ledger_df["status"] == "captured").sum() / len(ledger_df)
) * 100

merged_rec = pd.merge(
    ledger_df,
    gateway_df,
    on="transaction_id",
    suffixes=("_ledger", "_gateway"),
)
matched_count = (
    (merged_rec["amount_inr_ledger"] == merged_rec["amount_inr_gateway"])
    & (merged_rec["status_ledger"] == merged_rec["status_gateway"])
).sum()
match_rate = (matched_count / len(ledger_df)) * 100

overall_cb_ratio = (
    (ledger_df["status"] == "chargeback").sum() / len(ledger_df)
) * 100

fig, ax = plt.subplots(figsize=(12, 3))
ax.axis("off")

scorecards = [
    ("Total GMV", f"₹{total_gmv:,.0f}"),
    ("Success Rate", f"{overall_success_rate:.1f}%"),
    ("Reconciled Match Rate", f"{match_rate:.1f}%"),
    ("Chargeback Ratio", f"{overall_cb_ratio:.2f}%"),
]

for idx, (title, val) in enumerate(scorecards):
    rect = plt.Rectangle(
        (idx * 0.25 + 0.02, 0.15),
        0.21,
        0.7,
        facecolor="#F4F6F9",
        edgecolor="#2C3E50",
    )
    ax.add_patch(rect)
    ax.text(
        idx * 0.25 + 0.125,
        0.60,
        title,
        fontsize=12,
        ha="center",
        color="#7F8C8D",
        weight="bold",
    )
    ax.text(
        idx * 0.25 + 0.125,
        0.32,
        val,
        fontsize=16,
        ha="center",
        color="#2C3E50",
        weight="bold",
    )

plt.title(
    "Paytm Payments Analytics — Headline Scorecards",
    fontsize=14,
    weight="bold",
    pad=15,
)
plt.tight_layout()
plt.savefig(
    "dashboard_output/layer1_headline_scorecards.png", dpi=300, bbox_inches="tight"
)
plt.close()

# Layer 2: Trends Layer (Time Series)
daily_stats = (
    ledger_merged.groupby("date")
    .agg(
        daily_gmv=("amount_inr", "sum"),
        cb_count=("status", lambda s: (s == "chargeback").sum()),
    )
    .reset_index()
)

fig, ax1 = plt.subplots(figsize=(12, 5))

color1 = "#1f77b4"
ax1.set_xlabel("Date", fontsize=11, labelpad=10)
ax1.set_ylabel("Daily GMV (INR)", color=color1, fontsize=11, weight="bold")
ax1.plot(
    daily_stats["date"],
    daily_stats["daily_gmv"],
    color=color1,
    linewidth=2.5,
    marker="o",
)
ax1.tick_params(axis="y", labelcolor=color1)
ax1.grid(True, linestyle="--", alpha=0.5)

ax2 = ax1.twinx()
color2 = "#d62728"
ax2.set_ylabel(
    "Chargeback Count", color=color2, fontsize=11, weight="bold"
)
ax2.bar(
    daily_stats["date"],
    daily_stats["cb_count"],
    color=color2,
    alpha=0.4,
    width=0.6,
)
ax2.tick_params(axis="y", labelcolor=color2)

plt.title(
    "30-Day Trend: Daily GMV vs. Daily Chargeback Count",
    fontsize=14,
    weight="bold",
    pad=15,
)
fig.autofmt_xdate()
plt.tight_layout()
plt.savefig(
    "dashboard_output/layer2_trends_timeseries.png", dpi=300, bbox_inches="tight"
)
plt.close()

# Layer 3: Breakdown Layer (Bar Charts)
fig, (ax1, ax2) = plt.subplots(1, 2, figsize=(14, 5))

method_gmv = (
    ledger_merged.groupby("payment_method")["amount_inr"]
    .sum()
    .sort_values(ascending=False)
)
ax1.bar(
    method_gmv.index,
    method_gmv.values / 1000,
    color=["#2B5B84", "#3A86C8", "#59A5E8", "#8EC5FC"],
)
ax1.set_title("GMV by Payment Method (in ₹'000)", fontsize=12, weight="bold")
ax1.set_ylabel("GMV (Thousands INR)")
for i, v in enumerate(method_gmv.values):
    ax1.text(i, (v / 1000) + 2, f"₹{v/1000:.1f}k", ha="center", weight="bold")

cat_gmv = (
    ledger_merged.groupby("category")["amount_inr"]
    .sum()
    .sort_values(ascending=False)
)
ax2.bar(cat_gmv.index, cat_gmv.values / 1000, color="#2E8B57")
ax2.set_title("GMV by Merchant Category (in ₹'000)", fontsize=12, weight="bold")
ax2.set_ylabel("GMV (Thousands INR)")
ax2.tick_params(axis="x", rotation=30)
for i, v in enumerate(cat_gmv.values):
    ax2.text(i, (v / 1000) + 2, f"₹{v/1000:.1f}k", ha="center", weight="bold")

plt.tight_layout()
plt.savefig(
    "dashboard_output/layer3_breakdown_gmv.png", dpi=300, bbox_inches="tight"
)
plt.close()

# Layer 4: Details Layer (Formatted Table Image)
merchant_summary = (
    ledger_merged.groupby(["merchant_id", "merchant_name"])
    .agg(
        total_txns=("transaction_id", "count"),
        cb_txns=("status", lambda s: (s == "chargeback").sum()),
        total_gmv=("amount_inr", "sum"),
    )
    .reset_index()
)

merchant_summary["cb_ratio_pct"] = (
    merchant_summary["cb_txns"] / merchant_summary["total_txns"]
) * 100

top10_merchants = merchant_summary.sort_values(
    by="total_txns", ascending=False
).head(10)
top10_merchants["High-Risk Flag"] = top10_merchants["cb_ratio_pct"].apply(
    lambda r: "HIGH RISK (>1%)" if r > 1.0 else "Normal"
)

fig, ax = plt.subplots(figsize=(10, 4))
ax.axis("off")

table_data = []
headers = [
    "Merchant Name",
    "Total Txns",
    "Total GMV (₹)",
    "Chargebacks",
    "CB Ratio (%)",
    "Risk Flag",
]

for _, row in top10_merchants.iterrows():
    table_data.append([
        row["merchant_name"],
        f"{row['total_txns']}",
        f"₹{row['total_gmv']:,}",
        f"{row['cb_txns']}",
        f"{row['cb_ratio_pct']:.2f}%",
        row["High-Risk Flag"],
    ])

tbl = ax.table(
    cellText=table_data,
    colLabels=headers,
    cellLoc="center",
    loc="center",
    colColours=["#2C3E50"] * len(headers),
)
tbl.auto_set_font_size(False)
tbl.set_fontsize(9)
tbl.scale(1.2, 1.4)

for (row_idx, col_idx), cell in tbl.get_celld().items():
    if row_idx == 0:
        cell.get_text().set_color("white")
        cell.get_text().set_weight("bold")
    else:
        flag_val = table_data[row_idx - 1][5]
        if flag_val == "HIGH RISK (>1%)":
            if col_idx == 5:
                cell.set_facecolor("#FFCCCC")
                cell.get_text().set_color("#900C3F")
                cell.get_text().set_weight("bold")

plt.title(
    "Top 10 Merchants by Transaction Count — Performance & Risk Detail",
    fontsize=12,
    weight="bold",
    pad=10,
)
plt.savefig(
    "dashboard_output/layer4_details_top10_merchants.png",
    dpi=300,
    bbox_inches="tight",
)
plt.close()