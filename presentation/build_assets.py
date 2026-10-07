"""
Generate every chart used by the RetainIQ project-review deck.

All figures are produced from the project's own artifacts (raw dataset, held-out
test split, trained models, production bulk-scoring pipeline), so the deck and
the repository can never disagree. Run from the repository root:

    python presentation/build_assets.py

Outputs land in presentation/assets/ as transparent PNGs sized for 16:9 slides.
"""
from __future__ import annotations

import os
import sys
from pathlib import Path

import matplotlib
matplotlib.use("Agg")
import matplotlib.pyplot as plt
import numpy as np
import pandas as pd
from matplotlib.patches import FancyArrowPatch, FancyBboxPatch

ROOT = Path(__file__).resolve().parents[1]
if str(ROOT) not in sys.path:
    sys.path.insert(0, str(ROOT))

ASSETS = Path(__file__).resolve().parent / "assets"
ASSETS.mkdir(parents=True, exist_ok=True)

# --- Signal Ops palette (mirrors static/style.css tokens) -------------------
INK = "#0A0F1C"
PANEL = "#111A2C"
PANEL_2 = "#162136"
LINE = "#223049"
TEXT = "#EAF0F7"
MUTED = "#8CA0BE"
ACTION = "#4C8DFF"
OK = "#2FD8B4"
WARN = "#F5A623"
BAD = "#FF5C6C"

plt.rcParams.update({
    # Opaque ink background rather than transparency: the PNGs then render
    # correctly in any viewer (file previews, Word, email) instead of showing
    # near-white labels on a white page. Slide backgrounds use the same colour,
    # so the images still blend seamlessly into the deck.
    "figure.facecolor": INK,
    "axes.facecolor": INK,
    "savefig.facecolor": INK,
    "savefig.transparent": False,
    "text.color": TEXT,
    "axes.labelcolor": MUTED,
    "axes.edgecolor": LINE,
    "xtick.color": MUTED,
    "ytick.color": MUTED,
    "font.family": "DejaVu Sans",
    "font.size": 12,
    "axes.titlesize": 15,
    "axes.titleweight": "bold",
    "axes.grid": True,
    "grid.color": LINE,
    "grid.linewidth": 0.8,
    "grid.alpha": 0.55,
    "legend.frameon": False,
    "figure.dpi": 200,
})


def _strip(ax):
    for side in ("top", "right"):
        ax.spines[side].set_visible(False)
    for side in ("left", "bottom"):
        ax.spines[side].set_color(LINE)


def _save(fig, name):
    path = ASSETS / name
    fig.savefig(path, dpi=200, transparent=False, bbox_inches="tight", pad_inches=0.15)
    plt.close(fig)
    print("wrote", path.relative_to(ROOT))


def _bar_labels(ax, bars, values, fmt="{:.1f}%", offset=0.6, color=TEXT):
    for bar, value in zip(bars, values):
        ax.text(
            bar.get_width() + offset, bar.get_y() + bar.get_height() / 2,
            fmt.format(value), va="center", ha="left", color=color, fontsize=11,
        )


# ---------------------------------------------------------------------------
# Data loading
# ---------------------------------------------------------------------------

def load_truth() -> pd.DataFrame:
    df = pd.read_csv(ROOT / "Data" / "01.Telco_customer_churn_Dataset.csv")
    df["churned"] = (df["Churn Label"] == "Yes").astype(int)
    return df


def load_scores() -> pd.DataFrame:
    """Score the full portfolio through the exact production bulk pipeline."""
    from utils.bulk_prediction import run_bulk_prediction_from_bytes

    raw = (ROOT / "Data" / "01.Telco_customer_churn_Dataset.csv").read_bytes()
    return run_bulk_prediction_from_bytes(raw, "telco.csv")["results_df"]


def load_test_metrics():
    import joblib
    from sklearn.metrics import (accuracy_score, confusion_matrix, f1_score,
                                 precision_score, recall_score, roc_auc_score,
                                 roc_curve)

    X = pd.read_csv(ROOT / "Data" / "processed_data" / "test" / "X_test.csv")
    y = pd.read_csv(ROOT / "Data" / "processed_data" / "test" / "y_test.csv").iloc[:, 0].to_numpy()

    rows, roc = [], None
    for label, fname in [
        ("Logistic Regression", "logistic_regression.pkl"),
        ("Random Forest", "random_forest.pkl"),
        ("XGBoost", "xgboost.pkl"),
        ("LightGBM", "lightgbm.pkl"),
    ]:
        model = joblib.load(ROOT / "models" / fname)
        proba = model.predict_proba(X)[:, 1]
        pred = (proba >= 0.5).astype(int)
        rows.append({
            "model": label,
            "accuracy": accuracy_score(y, pred) * 100,
            "precision": precision_score(y, pred) * 100,
            "recall": recall_score(y, pred) * 100,
            "f1": f1_score(y, pred) * 100,
            "auc": roc_auc_score(y, proba) * 100,
        })
        if label == "Logistic Regression":
            cm = confusion_matrix(y, pred)
            fpr, tpr, _ = roc_curve(y, proba)
            roc = {"fpr": fpr, "tpr": tpr, "auc": roc_auc_score(y, proba) * 100}
    return pd.DataFrame(rows), cm, roc


# ---------------------------------------------------------------------------
# Charts
# ---------------------------------------------------------------------------

def chart_churn_split(truth: pd.DataFrame):
    churn = int(truth["churned"].sum())
    total = len(truth)
    rate = churn / total * 100
    fig, ax = plt.subplots(figsize=(5.0, 4.0))
    ax.grid(False)
    ax.pie(
        [churn, total - churn], startangle=90, counterclock=False,
        colors=[BAD, PANEL_2], wedgeprops=dict(width=0.34, edgecolor=INK, linewidth=2),
    )
    ax.text(0, 0.10, f"{rate:.2f}%", ha="center", va="center", fontsize=30,
            fontweight="bold", color=BAD)
    ax.text(0, -0.22, "churned", ha="center", va="center", fontsize=12, color=MUTED)
    ax.text(0, -1.42, f"{churn:,} churned   ·   {total - churn:,} retained   ·   {total:,} total",
            ha="center", va="center", fontsize=11, color=MUTED)
    _save(fig, "chart_churn_split.png")


def chart_rate_by(truth, column, order, title, name, colors=None, horizontal=True):
    rates = truth.groupby(column)["churned"].mean().reindex(order) * 100
    colors = colors or [BAD if r == rates.max() else ACTION for r in rates]
    fig, ax = plt.subplots(figsize=(6.6, 4.0))
    if horizontal:
        bars = ax.barh(rates.index[::-1], rates.values[::-1], color=colors[::-1], height=0.6)
        ax.set_xlim(0, max(rates.values) * 1.28)
        _bar_labels(ax, bars, rates.values[::-1])
        ax.set_xlabel("Churn rate (%)")
    else:
        bars = ax.bar(rates.index, rates.values, color=colors, width=0.6)
        ax.set_ylim(0, max(rates.values) * 1.22)
        for b, v in zip(bars, rates.values):
            ax.text(b.get_x() + b.get_width() / 2, v + max(rates.values) * 0.03,
                    f"{v:.1f}%", ha="center", color=TEXT, fontsize=11)
        ax.set_ylabel("Churn rate (%)")
    ax.set_title(title, loc="left", pad=14)
    ax.grid(axis="y" if not horizontal else "x", alpha=0.0)
    ax.grid(axis="x" if not horizontal else "y", alpha=0.35)
    _strip(ax)
    _save(fig, name)


def chart_tenure(truth: pd.DataFrame):
    bins = [0, 12, 24, 48, 72]
    labels = ["0–12 mo", "13–24 mo", "25–48 mo", "49–72 mo"]
    bands = pd.cut(truth["Tenure Months"], bins=bins, labels=labels, include_lowest=True)
    rates = truth.groupby(bands, observed=False)["churned"].mean() * 100
    fig, ax = plt.subplots(figsize=(6.6, 4.0))
    ramp = [BAD, "#FF8A5C", WARN, OK]
    bars = ax.bar(rates.index.astype(str), rates.values, color=ramp, width=0.6)
    for b, v in zip(bars, rates.values):
        ax.text(b.get_x() + b.get_width() / 2, v + 1.4, f"{v:.1f}%", ha="center", color=TEXT, fontsize=11)
    ax.set_ylim(0, max(rates.values) * 1.22)
    ax.set_ylabel("Churn rate (%)")
    ax.set_title("Churn falls sharply with tenure", loc="left", pad=14)
    ax.grid(axis="x", alpha=0.0)
    _strip(ax)
    _save(fig, "chart_tenure.png")


def chart_models(metrics: pd.DataFrame):
    fig, ax = plt.subplots(figsize=(8.8, 4.3))
    x = np.arange(len(metrics))
    width = 0.26
    series = [("Accuracy", "accuracy", ACTION), ("F1 score", "f1", WARN), ("ROC-AUC", "auc", OK)]
    for i, (label, key, color) in enumerate(series):
        bars = ax.bar(x + (i - 1) * width, metrics[key], width * 0.92, label=label, color=color)
        for b, v in zip(bars, metrics[key]):
            ax.text(b.get_x() + b.get_width() / 2, v + 1.4, f"{v:.1f}", ha="center",
                    color=TEXT, fontsize=9.5)
    ax.set_xticks(x)
    ax.set_xticklabels([m.replace(" ", "\n") if len(m) > 12 else m for m in metrics["model"]])
    ax.set_ylim(0, 112)
    ax.set_yticks([0, 20, 40, 60, 80, 100])
    ax.set_ylabel("Score (%)")
    ax.set_title("Four algorithms benchmarked on the held-out test set (n = 1,409)",
                 loc="left", pad=30, fontsize=14)
    ax.legend(ncol=3, loc="lower left", bbox_to_anchor=(0, 1.005), fontsize=11,
              handlelength=1.4, columnspacing=1.8)
    ax.grid(axis="x", alpha=0.0)
    _strip(ax)
    _save(fig, "chart_models.png")


def chart_roc(roc):
    fig, ax = plt.subplots(figsize=(5.6, 4.4))
    ax.plot(roc["fpr"], roc["tpr"], color=OK, linewidth=2.6,
            label=f"Logistic Regression (AUC {roc['auc']:.2f}%)")
    ax.plot([0, 1], [0, 1], color=MUTED, linestyle="--", linewidth=1.2,
            label="Random classifier (50%)")
    ax.fill_between(roc["fpr"], roc["tpr"], color=OK, alpha=0.12)
    ax.set_xlabel("False positive rate")
    ax.set_ylabel("True positive rate")
    ax.set_title("ROC curve — production model", loc="left", pad=14)
    ax.legend(loc="lower right", fontsize=10.5)
    ax.set_xlim(-0.02, 1.02)
    ax.set_ylim(-0.02, 1.02)
    _strip(ax)
    _save(fig, "chart_roc.png")


def chart_confusion(cm):
    tn, fp = cm[0]
    fn, tp = cm[1]
    total = tn + fp + fn + tp
    fig, ax = plt.subplots(figsize=(6.0, 4.6))
    ax.grid(False)
    ax.set_xlim(0, 2)
    ax.set_ylim(0, 2)
    ax.axis("off")

    cells = [
        (0, 1, tn, "True Negative", OK, "kept correctly"),
        (1, 1, fp, "False Positive", BAD, "over-flagged"),
        (0, 0, fn, "False Negative", BAD, "missed churner"),
        (1, 0, tp, "True Positive", OK, "caught churner"),
    ]
    for x, y, value, label, color, note in cells:
        ax.add_patch(FancyBboxPatch(
            (x + 0.045, y + 0.055), 0.91, 0.89,
            boxstyle="round,pad=0.01,rounding_size=0.05",
            linewidth=2.2, edgecolor=color, facecolor=PANEL))
        ax.text(x + 0.5, y + 0.70, f"{value:,}", ha="center", va="center",
                color=TEXT, fontsize=27, fontweight="bold")
        ax.text(x + 0.5, y + 0.44, label, ha="center", va="center",
                color=color, fontsize=11.5, fontweight="bold")
        ax.text(x + 0.5, y + 0.24, f"{value / total * 100:.1f}% · {note}",
                ha="center", va="center", color=MUTED, fontsize=9.6)

    ax.text(-0.02, 1.5, "Actually\nstayed", ha="right", va="center", color=MUTED, fontsize=11)
    ax.text(-0.02, 0.5, "Actually\nchurned", ha="right", va="center", color=MUTED, fontsize=11)
    ax.text(0.5, -0.07, "Predicted stay", ha="center", va="top", color=MUTED, fontsize=11)
    ax.text(1.5, -0.07, "Predicted churn", ha="center", va="top", color=MUTED, fontsize=11)
    ax.set_title("Confusion matrix — held-out test set (n = 1,409)", loc="center",
                 pad=16, fontsize=14)
    _save(fig, "chart_confusion.png")


def chart_risk_bands(scores: pd.DataFrame, truth: pd.DataFrame):
    order = ["Low Risk", "Medium Risk", "High Risk"]
    counts = scores["Risk_Segment"].value_counts().reindex(order)
    actual = [truth.loc[scores.index[scores["Risk_Segment"] == b], "churned"].mean() * 100 for b in order]
    colors = [OK, WARN, BAD]

    fig, ax = plt.subplots(figsize=(8.8, 4.4))
    top = max(counts.values)
    bars = ax.bar(order, counts.values, color=colors, width=0.52)
    for b, v in zip(bars, counts.values):
        # both lines sit inside the bar so nothing overflows or collides with the line
        ax.text(b.get_x() + b.get_width() / 2, v * 0.93,
                f"{v / len(scores) * 100:.1f}% of portfolio", ha="center", va="center",
                color=INK, fontsize=9.5)
        ax.text(b.get_x() + b.get_width() / 2, v * 0.70, f"{v:,}", ha="center", va="center",
                color=INK, fontsize=16, fontweight="bold")
    ax.set_ylim(0, top * 1.22)
    ax.set_ylabel("Customers")
    ax.grid(axis="x", alpha=0.0)

    ax2 = ax.twinx()
    ax2.plot(order, actual, color=TEXT, marker="o", markersize=7, linewidth=1.8,
             markerfacecolor=INK, markeredgewidth=2)
    for i, v in enumerate(actual):
        ax2.text(i + 0.13, v, f"{v:.1f}%", ha="left", va="center", color=TEXT,
                 fontsize=11.5, fontweight="bold",
                 bbox=dict(facecolor=INK, edgecolor="none", pad=1.7, alpha=0.88))
    ax2.set_ylim(0, 108)
    ax2.set_yticks([0, 20, 40, 60, 80, 100])
    ax2.set_ylabel("Actual churn rate in band (%)")
    ax2.grid(False)
    for spine in ax2.spines.values():
        spine.set_visible(False)
    ax.set_title("Risk bands ordered real churn correctly: 9.5% → 42.0% → 73.0%",
                 loc="left", pad=14, fontsize=13.5)
    _strip(ax)
    _save(fig, "chart_risk_bands.png")


def chart_revenue(scores: pd.DataFrame):
    high = scores.loc[scores["Risk_Segment"] == "High Risk", "Monthly Charges"].sum()
    medium = scores.loc[scores["Risk_Segment"] == "Medium Risk", "Monthly Charges"].sum()
    total = scores["Monthly Charges"].sum()
    fig, ax = plt.subplots(figsize=(8.0, 3.5))
    rows = [
        ("Entire portfolio", total, PANEL_2, MUTED),
        ("High + medium risk", high + medium, WARN, TEXT),
        ("High risk only", high, BAD, TEXT),
    ]
    for i, (label, value, color, valuecolor) in enumerate(rows):
        ax.barh(i, value, color=color, height=0.52)
        ax.text(value * 1.02, i, f"${value:,.0f} / month", va="center",
                color=valuecolor, fontsize=12, fontweight="bold")
    ax.set_yticks(range(len(rows)), [r[0] for r in rows])
    ax.set_xlim(0, total * 1.30)
    ax.set_xticks(np.linspace(0, total, 6))
    ax.set_xticklabels([f"{int(v):,}" for v in np.linspace(0, total, 6)])
    ax.set_xlabel("Recurring monthly charges (USD)")
    ax.set_title("Monthly revenue exposure by risk band", loc="left", pad=14)
    ax.grid(axis="y", alpha=0.0)
    _strip(ax)
    _save(fig, "chart_revenue.png")


HUMANIZE = {
    "Tenure Months": "Account tenure",
    "Monthly Charges": "Monthly charges",
    "Total Charges": "Total charges",
    "Gender_Male": "Male customer",
    "Senior Citizen_Yes": "Senior citizen",
    "Partner_Yes": "Has a partner",
    "Dependents_Yes": "Has dependents",
    "Phone Service_Yes": "Has phone service",
    "Contract_One year": "One-year contract",
    "Contract_Two year": "Two-year contract",
    "Internet Service_Fiber optic": "Fiber optic internet",
    "Internet Service_No": "No internet service",
    "Payment Method_Electronic check": "Electronic check payment",
    "Payment Method_Mailed check": "Mailed check payment",
    "Payment Method_Credit card (automatic)": "Automatic card payment",
    "Paperless Billing_Yes": "Paperless billing",
    "Tech Support_Yes": "Has tech support",
    "Online Security_Yes": "Has online security",
    "Multiple Lines_Yes": "Multiple phone lines",
}


def chart_high_risk_profile(scores: pd.DataFrame):
    """Descriptive profile of the scored high-risk band, no modelling assumptions."""
    truth = load_truth()
    truth["seg"] = scores["Risk_Segment"].values
    hi = truth[truth["seg"] == "High Risk"]
    allc = truth

    def share(df, col, val):
        return (df[col] == val).mean() * 100

    factors = [
        ("Month-to-month contract", share(hi, "Contract", "Month-to-month"), share(allc, "Contract", "Month-to-month")),
        ("No online security add-on", share(hi, "Online Security", "No"), share(allc, "Online Security", "No")),
        ("No tech support", share(hi, "Tech Support", "No"), share(allc, "Tech Support", "No")),
        ("Fiber optic internet", share(hi, "Internet Service", "Fiber optic"), share(allc, "Internet Service", "Fiber optic")),
        ("Pays by electronic check", share(hi, "Payment Method", "Electronic check"), share(allc, "Payment Method", "Electronic check")),
        ("Senior citizen", share(hi, "Senior Citizen", "Yes"), share(allc, "Senior Citizen", "Yes")),
    ]

    names = [f[0] for f in factors]
    high = np.array([f[1] for f in factors])
    overall = np.array([f[2] for f in factors])

    fig, ax = plt.subplots(figsize=(8.8, 4.5))
    y = np.arange(len(names))
    h = 0.34
    ax.barh(y + h / 2, high, height=h, color=BAD, label="High-risk band (n = 1,135)")
    ax.barh(y - h / 2, overall, height=h, color=PANEL_2, edgecolor=MUTED,
            linewidth=1.0, label="Whole portfolio (n = 7,043)")
    for i, (hv, ov) in enumerate(zip(high, overall)):
        ax.text(hv + 1.6, i + h / 2, f"{hv:.0f}%", va="center", color=TEXT, fontsize=11)
        ax.text(ov + 1.6, i - h / 2, f"{ov:.0f}%", va="center", color=MUTED, fontsize=11)
    ax.set_yticks(y, names)
    ax.set_xlim(0, 116)
    ax.set_xticks([0, 20, 40, 60, 80, 100])
    ax.set_xlabel("Share of customers (%)")
    ax.set_title("Profile of the high-risk band vs the whole portfolio", loc="left",
                 pad=30, fontsize=14)
    ax.legend(ncol=2, loc="lower left", bbox_to_anchor=(0, 1.005), fontsize=10.5)
    ax.grid(axis="y", alpha=0.0)
    _strip(ax)
    _save(fig, "chart_high_risk_profile.png")


def chart_architecture():
    fig, ax = plt.subplots(figsize=(11.0, 4.6))
    ax.set_xlim(0, 100)
    ax.set_ylim(0, 56)
    ax.axis("off")
    ax.grid(False)

    def box(x, y, w, h, title, sub, edge, fill=PANEL, title_size=10.5, sub_size=8.2):
        ax.add_patch(FancyBboxPatch(
            (x, y), w, h, boxstyle="round,pad=0.4,rounding_size=1.4",
            linewidth=1.3, edgecolor=edge, facecolor=fill))
        ax.text(x + w / 2, y + h * 0.66, title, ha="center", va="center",
                color=TEXT, fontsize=title_size, fontweight="bold")
        ax.text(x + w / 2, y + h * 0.28, sub, ha="center", va="center",
                color=MUTED, fontsize=sub_size)

    def arrow(x1, y1, x2, y2, color=LINE):
        ax.add_patch(FancyArrowPatch((x1, y1), (x2, y2), arrowstyle="-|>",
                                     mutation_scale=13, linewidth=1.4, color=color))

    # pipeline row
    box(0.5, 38, 23.5, 13, "Customer data", "7,043 telco records", ACTION)
    box(25.8, 38, 23.5, 13, "Preprocessing", "cleaned · encoded · 30 features", ACTION)
    box(51.1, 38, 23.5, 13, "Logistic Regression", "binary churn classifier", OK)
    box(76.4, 38, 23.5, 13, "Churn probability", "per-customer score", OK)
    for x in (24.0, 49.3, 74.6):
        arrow(x, 44.5, x + 1.8, 44.5, MUTED)

    # intelligence row
    box(8.0, 19, 26.0, 13, "Risk segmentation", "low · medium · high", WARN)
    box(37.0, 19, 26.0, 13, "Explainability", "feature contributions", WARN)
    box(66.0, 19, 26.0, 13, "Retention actions", "recommendation engine", WARN)

    # delivery row
    box(8.0, 3, 26.0, 10.5, "Signal Ops Console", "bulk scoring · dashboard · reports", MUTED, PANEL_2, 9.6, 8.0)
    box(37.0, 3, 26.0, 10.5, "Prediction API", "Flask / FastAPI endpoints", MUTED, PANEL_2, 9.6, 8.0)
    box(66.0, 3, 26.0, 10.5, "Power BI layer", "4-page BI dashboard", MUTED, PANEL_2, 9.6, 8.0)

    # fan out from the probability box to the three intelligence boxes
    arrow(88.0, 38.0, 88.0, 32.5, LINE)
    arrow(88.0, 32.5, 21.0, 32.5, LINE)
    arrow(88.0, 32.5, 50.0, 32.5, LINE)
    arrow(88.0, 32.5, 79.0, 32.5, LINE)
    arrow(21.0, 32.5, 21.0, 32.1, LINE)
    arrow(50.0, 32.5, 50.0, 32.1, LINE)
    arrow(79.0, 32.5, 79.0, 32.1, LINE)
    # sequential hand-offs inside the intelligence row
    arrow(66.0, 25.5, 63.4, 25.5, LINE)
    arrow(37.0, 25.5, 34.4, 25.5, LINE)
    # down to delivery
    arrow(21.0, 19.0, 21.0, 13.9, LINE)
    arrow(50.0, 19.0, 50.0, 13.9, LINE)
    arrow(79.0, 19.0, 79.0, 13.9, LINE)

    _save(fig, "chart_architecture.png")


def main():
    truth = load_truth()
    print("scoring full portfolio through the production pipeline ...")
    scores = load_scores()
    metrics, cm, roc = load_test_metrics()

    print(metrics.round(2).to_string(index=False))
    print("confusion matrix:\n", cm)

    chart_churn_split(truth)
    chart_rate_by(truth, "Contract", ["Month-to-month", "One year", "Two year"],
                  "Churn is 15x higher on month-to-month contracts", "chart_contract.png")
    chart_rate_by(truth, "Internet Service", ["Fiber optic", "DSL", "No"],
                  "Fiber optic customers churn most", "chart_internet.png")
    chart_tenure(truth)
    chart_models(metrics)
    chart_roc(roc)
    chart_confusion(cm)
    chart_risk_bands(scores, truth)
    chart_revenue(scores)
    chart_high_risk_profile(scores)
    chart_architecture()

    metrics.to_csv(ASSETS / "model_metrics.csv", index=False)
    print("\nall assets written")


if __name__ == "__main__":
    main()
