# PPT Content Checklist — RetainIQ Project Review

Every detail the deck must contain, with the verified value for each.

Status: ✅ = already in `RetainIQ_Project_Review.pptx` (20 presentation slides: the
original review deck's final slide and final five appendix pages are omitted, and
one Power BI vs RetainIQ comparison slide has been added). The business problem
remains slide 3; the comparison is slide 4, before the dataset on slide 5. Every
slide is laid out on the vertical fill engine in
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
- ✅ Base-rate framing: predicting "nobody churns" is already 73.46% accurate

## 4. Existing project layer vs RetainIQ (slide 4)

- ✅ Placement: after the cycle map and business problem; before dataset collection
- ✅ Power BI: portfolio-level historical churn KPIs and segment patterns
- ✅ RetainIQ: uploaded customer rows scored with deployed Logistic Regression;
  churn probability, risk tier, coefficient-based drivers, rule-based suggestions,
  bulk review and private report downloads
- ✅ Difference: portfolio monitoring vs customer-level prioritisation
- ✅ Complementary layers; no campaign uplift measured; revenue remains at-risk exposure

## 5. Dataset (slide 5)

- ✅ Source: IBM Telco Customer Churn
- ✅ 7,043 customers · 34 raw features · 30 final ML features
- ✅ Target: Churn Label (binary)
- ✅ Class balance: 1,869 churned (26.54%) / 5,174 stayed (73.46%)
- ✅ Split: 80/20 stratified → 5,634 train / 1,409 test (both exactly 26.54%)
- ✅ Feature groups: profile, services, contract & billing, value
- ✅ **Leakage control:** dropped Churn Score, Churn Reason, Churn Category, Customer Status
- ✅ Train/serve parity: 30 columns frozen in `feature_columns.pkl`, unit-tested

## 6. Preprocessing & feature engineering (slides 6 and 8)

- ✅ Cleaning: dtype fixes, blanks, duplicates, category normalisation
- ✅ 11 blank TotalCharges, all zero-tenure new customers
- ✅ Encoding: 3 numeric features + 27 binary flags = 30
- ✅ Removal of ID and geographic columns
- ✅ Why one shared pipeline matters (notebook and app use the same code)

## 7. EDA findings (slide 7) — pick 3, hold the rest as backup

- ✅ Contract: month-to-month 42.7% vs one-year 11.3% vs two-year 2.8% (15× gap)
- ✅ Tenure: 47.4% churn in year one; 55.5% of all churners are in year one
- ✅ Avg tenure: churned 18.0 months vs stayed 37.6 months
- ✅ Avg monthly charges: churned $74.44 vs stayed $61.27
- ✅ Internet service: fiber optic 41.9% vs DSL 19.0% vs none 7.4%
- ✅ Payment method: electronic check 45.3% vs auto-pay ~15–17%
- ✅ Senior citizens: 41.7% vs 23.6%
- ✅ Missing add-ons: no online security 41.8%, no tech support 41.6%

## 8. Model development (slide 9)

- ✅ Four algorithms: Logistic Regression, Random Forest, XGBoost, LightGBM
- ✅ Rationale for each (linear baseline / bagging / two boosting implementations)
- ✅ Same features, same split, same scoring code — only the algorithm differs
- ✅ Why more than one model (result is a property of the data, not one config)

## 9. Evaluation & model selection (slides 10–12) — the credibility section

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

## 10. Risk segmentation (slide 14)

- ✅ Thresholds: < 30% low · 30–60% medium · ≥ 60% high
- ✅ Band sizes: 4,435 low · 1,473 medium · 1,135 high
- ✅ **Validation against real labels:** actual churn 9.5% / 42.0% / 73.0%
- ✅ Operational payoff: 1,135 to prioritise instead of 7,043

## 11. Explainability (slide 13)

- ✅ Per-customer drivers = coefficient × value (honest for a linear model)
- ✅ High-risk band profile: 100% month-to-month · 95% no online security · 94% no tech support · 91% fiber optic · 80% electronic check · 32% senior · avg tenure 9.6 mo
- ✅ SHAP status stated precisely: explored in research, production uses coefficients

## 12. Retention intelligence (slide 15)

- ✅ 2,608 at-risk customers (high + medium) each got a recommendation
- ✅ Recommendation breakdown: Long-Term Contract 792 · 5% Discount 663 · 15% Discount 658 · Welcome Package 430 · Free Online Security 45 · Autopay 9 · Check-in 9 · Premium Support 2
- ✅ Rules driven by the customer's own risk factors
- ✅ Decision framework: Data → Insight → Prediction → Explanation → Action → Value

## 13. Business value / revenue (slide 15)

- ✅ High-risk monthly revenue at risk: **$92,539**
- ✅ High + medium: **$200,302**
- ✅ Entire portfolio: $456,117
- ✅ Framing: "revenue at risk / exposure" — never "revenue saved"

## 14. Product & demo (slide 17)

- ✅ System architecture diagram (data → model → risk/explain/action → console/API/BI)
- ✅ Flask console modules: Bulk Prediction, Dashboard, Reports, About
- ✅ Workflow: upload → validate → score → KPIs → drill-down → download
- ✅ Screenshots: home, bulk upload, dashboard, reports
- ✅ Note: single-prediction page was retired (route redirects home)

## 15. Engineering & deployment (slide 18)

- ✅ Power BI: 4 pages — executive overview, customer insights, risk intelligence, retention strategy
- ✅ Clickable `.pbit` template link on slide 16 (GitHub file page)
- ✅ API: Flask `/api/health`, `/api/predict` + FastAPI wrapper with Pydantic validation
- ✅ Tests: 9 unit tests (encoding, model load, scoring, risk, recommendations, CSV validation, report privacy)
- ✅ Deployment: Gunicorn (gthread, 300s timeout, preload, worker recycling), Docker, Procfile, Render
- ✅ Reproducibility: one script regenerates every chart and metric

## 16. Technology stack (slide 19)

- ✅ Built: tool logos in a left-side rail; spacious grouped stack details on the right

- **Language & data:** Python 3.11 · Pandas · NumPy · SQL Server · SQLAlchemy
- **Machine learning:** Scikit-learn · Logistic Regression (deployed) · Random Forest · XGBoost · LightGBM · Joblib · SHAP
- **Web application:** Flask · Jinja2 · HTML5 · CSS3 · vanilla JavaScript · Gunicorn
- **APIs & delivery:** FastAPI + Pydantic · Docker · Render · Git/GitHub · Power BI (.pbix/.pbit)
- The reason the stack hangs together: one shared Python package for preprocessing + scoring
- ⚠️ **Do not claim Plotly** — it appears in an older README draft but is not imported anywhere in the code

## 17. Limitations (slide 20) — reviewers probe here

- ✅ Historical data from one telecom
- ✅ 56.95% recall — two in five churners missed
- ✅ No live behavioural feed, no drift monitoring, no auto-retraining
- ✅ Probabilities are not guarantees
- ✅ Revenue is exposure, not realised savings

## 18. Future work (slide 20)

- ✅ Threshold tuning and probability calibration
- ✅ Cost-sensitive learning
- ✅ Drift detection + scheduled retraining (MLOps)
- ✅ Counterfactual / what-if explanations
- ✅ CRM integration and A/B testing of offers

---

## Presentation mechanics (not slide content, but required)

- ✅ Speaker notes on every slide with timings totalling about 11 minutes 40 seconds
- ✅ 20 presentation slides: the comparison is added as slide 4; the final closing page and original final-five appendix pages are omitted.
- ✅ Section numbering and slide numbers
- ✅ Consistent dark "Signal Ops" theme matching the live product

## Assets that must exist in the repo

- ✅ 15 generated charts (`presentation/assets/`) — all reproducible
- ✅ Tool marks (Python, Pandas, Scikit-learn, Flask, FastAPI, Docker, Power BI,
  GitHub) in `presentation/assets/tool_logos/`; source attribution is included there
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
