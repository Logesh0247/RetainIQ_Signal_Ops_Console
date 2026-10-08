# PPT Content Checklist — RetainIQ Project Review

Every detail the deck must contain, with the verified value for each.

Status: ✅ = already in `RetainIQ_Project_Review.pptx` (11 slides: title, roadmap,
and project phases through operating-point selection; slides 12–16 were removed
at the user's request).
Every retained content slide is laid out on the vertical fill engine in
`presentation/build_deck.py` — kicker/heading → content band (1.66in–6.80in) →
rail — so no slide ends early. Verify with `audit_deck.py` (geometry),
`audit_layout.py` (dead space) and `preview_deck.py` (rasterised renders).

---

## 1. Title & identity (slide 1)

- ✅ Project name: **RetainIQ — Signal Ops Console**
- ✅ One-line description: explainable ML for churn prediction, risk intelligence, retention strategy
- ✅ Presenter: Logesh S., B.Sc. Data Science
- ✅ Live URL: retainiq-predictive-customer-retention-zq6x.onrender.com
- ✅ Source: github.com/Logesh0247/RetainIQ_Signal_Ops_Console
- ✅ Headline numbers: 80.34% accuracy · 84.91% ROC-AUC · 7,043 scored · $92,539 revenue at risk

## 2. Agenda / roadmap (slide 2)

- ✅ The one-line thesis: probability → action, not just a classification score
- ✅ Four-part structure of the talk

## 3. Business problem (slide 3)

- ✅ What churn costs; why dashboards describe but don't prioritise
- ✅ Why blanket campaigns waste budget
- ✅ **The four questions:** Who will churn? → Why? → What action? → What's at stake?
- ✅ Base rate framing: predicting "nobody churns" is already 73.46% accurate

## 4. Dataset (slide 4)

- ✅ Source: IBM Telco Customer Churn
- ✅ 7,043 customers · 34 raw features · 30 final ML features
- ✅ Target: Churn Label (binary)
- ✅ Class balance: 1,869 churned (26.54%) / 5,174 stayed (73.46%)
- ✅ Split: 80/20 stratified → 5,634 train / 1,409 test (both exactly 26.54%)
- ✅ Feature groups: profile, services, contract & billing, value
- ✅ **Leakage control:** dropped Churn Score, Churn Reason, Churn Category, Customer Status
- ✅ Train/serve parity: 30 columns frozen in `feature_columns.pkl`, unit-tested

## 5. Preprocessing (slide 5)

- ✅ Cleaning: dtype fixes, blanks, duplicates, category normalisation
- ✅ 11 blank TotalCharges, all zero-tenure new customers
- ✅ Encoding: 3 numeric features + 27 binary flags = 30
- ✅ Removal of ID and geographic columns
- ✅ Why one shared pipeline matters (notebook and app use the same code)

## 6. EDA findings (slide 6)

- ✅ Contract: month-to-month 42.7% vs one-year 11.3% vs two-year 2.8% (15× gap)
- ✅ Tenure: 47.4% churn in year one; 55.5% of all churners are in year one
- ✅ Avg tenure: churned 18.0 months vs stayed 37.6 months
- ✅ Avg monthly charges: churned $74.44 vs stayed $61.27
- ✅ Internet service: fiber optic 41.9% vs DSL 19.0% vs none 7.4%
- ✅ Payment method: electronic check 45.3% vs auto-pay ~15–17%
- ✅ Senior citizens: 41.7% vs 23.6%
- ✅ Missing add-ons: no online security 41.8%, no tech support 41.6%

## 7. Feature engineering (slide 7) and model development (slide 8)

- ✅ Four algorithms: Logistic Regression, Random Forest, XGBoost, LightGBM
- ✅ Rationale for each (linear baseline / bagging / two boosting implementations)
- ✅ Same features, same split, same scoring code — only the algorithm differs
- ✅ Why more than one model (result is a property of the data, not one config)

## 8. Evaluation (slide 9), model selection (slide 10) & operating point (slide 11)

- ✅ Full benchmark table (held-out 1,409 customers):

| Model | Accuracy | Precision | Recall | F1 | ROC-AUC |
|---|---|---|---|---|---|
| **Logistic Regression ← selected** | **80.34%** | **64.74%** | **56.95%** | **60.60%** | **84.91%** |
| LightGBM | 80.06% | 64.49% | 55.35% | 59.57% | 84.71% |
| XGBoost | 79.06% | 61.72% | 55.61% | 58.51% | 83.34% |
| Random Forest | 78.99% | 62.34% | 52.67% | 57.10% | 83.48% |

- ✅ Confusion matrix: TN 919 · FP 116 · FN 161 · TP 213
- ✅ ROC curve with AUC
- ✅ Selection reason: chosen on F1 + ROC-AUC, not accuracy alone
- ✅ Honest read on recall (56.95% — two in five churners missed)

## 9–18. Later project phases and closing material

- ⬜ Not in this 11-slide copy. Explainability, risk segmentation, retention
  intelligence, Power BI, web application, deployment, technology stack,
  responsible AI, thank-you/QR and appendix pages were removed when slides 12–16
  were cut. The cover retains clickable live-app and repository links.

---

## Presentation mechanics (not slide content, but required)

- ✅ Speaker notes on every slide with timings totalling ~10 minutes
- ⬜ No appendix in this shortened 16-slide version; all slides after 16 are omitted.
- ✅ Section numbering and slide numbers
- ✅ Consistent dark "Signal Ops" theme matching the live product

## Assets that must exist in the repo

- ✅ 15 generated charts (`presentation/assets/`) — all reproducible
- ✅ `model_metrics.csv` — exported benchmark table
- ✅ 5 console screenshots (`templates/images/`)
- ✅ 4 Power BI screenshots (`Power_BI_dashboard/Screenshots/`)
- ⬜ **Action for you:** re-capture the prediction-dashboard screenshot — the committed one says "Random Forest predictions" while the deployed model is Logistic Regression

## Rules to keep the deck defensible

1. Never say "revenue saved" — the internal CSV column is misnamed; the deck says "revenue at risk".
2. Never claim SHAP is in production — coefficients explain the deployed model.
3. Never compare the 26.54% historical churn rate with the 16.1% high-risk share as if they were the same metric.
4. Never select or praise a model on accuracy alone.
5. Never omit recall when quoting performance.
