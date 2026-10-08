# Master Prompt — RetainIQ Storytelling Deck

> Copy everything between the horizontal rules into the AI tool you want to build the
> deck with. It is self-contained: all facts, colours, fonts and structure are included,
> so it works with or without access to the repository.

---

## ROLE

You are a senior presentation designer and data-storytelling specialist. You build
investor-grade decks for technical products. You care about three things above all:
narrative momentum, pixel-accurate alignment, and factual integrity.

## TASK

Create a **22-slide 16:9 presentation** about **RetainIQ — Signal Ops Console**, an
explainable machine learning platform for customer churn prediction, risk intelligence
and retention strategy.

The deck must be **a story, not a report**. It must be **extremely attractive**, and its
visual theme must **exactly match the product's own design system**, described below.

---

# 1. THE STORY (this is the spine — every slide serves it)

**Logline:** A telecom is losing a quarter of its customers. Its dashboards can say what
already happened, but nobody can say who leaves next or what to do about it. RetainIQ
follows one customer's weakening signal from raw data to a predicted risk, an explained
reason, a recommended action, and the money it protects.

**Three-act structure:**

- **ACT I — THE SIGNAL FADES (slides 1–8).** The business pain. A customer base bleeding
  26.54% of its people. Reports that describe but cannot predict. The data that holds the
  pattern, and the first discoveries that reveal it.
- **ACT II — READING THE SIGNAL (slides 9–15).** Building and interrogating the model.
  Four algorithms, a measurement of how stable the result is, a deliberate choice of
  operating point, then the two questions that make predictions usable: *why* this
  customer, and *how urgent* is it.
- **ACT III — THE SIGNAL RECOVERED (slides 16–22).** What the intelligence is worth, the
  product that delivers it, the stack that carries it, the honest limits, and a closing
  return to the opening image — the signal, now read in time.

**Recurring storytelling motif — "signal strength."** The product models churn as a
customer's signal to the business weakening. Use a **1–5 bar signal-strength meter** as
the deck's recurring visual: strongest on the title slide, decaying in Act I, being read
in Act II, restored in Act III. This is the one motif that ties the story together.

**Tone:** confident, precise, engineering-honest. No hype, no exclamation marks, no
"revolutionary." Let the numbers carry the weight. Where there is a weakness, name it and
then show what was measured about it.

---

# 2. VISUAL THEME — match the application exactly

The app is a dark "network operations console." Recreate it precisely.

## 2.1 Colour tokens (use these exact hex values)

| Token | Hex | Use |
|---|---|---|
| Ink (page background) | `#0A0F1C` | every slide background |
| Panel | `#111A2C` | cards, tiles, containers |
| Panel raised | `#162136` | nested/emphasis panels |
| Hairline | `#223049` | 1px borders, dividers, grid rules |
| Line soft | `#1A2540` | secondary dividers |
| Text | `#EAF0F7` | body and headings |
| Muted | `#8CA0BE` | captions, axis labels, footnotes, eyebrows |
| Action blue | `#4C8DFF` | pipeline/data/flow, primary actions |
| Signal green | `#2FD8B4` | safe, confirmed, healthy, selected, delivered |
| Warn amber | `#F5A623` | medium risk, caution, experiments |
| Bad red | `#FF5C6C` | churn, high risk, the problem |

**Colour discipline:** red always means customer loss or high risk. Green always means
retained, safe, or verified. Amber means medium/experimental. Blue means data movement or
the product itself. Never use a colour decoratively outside these meanings.

**No gradients. No drop shadows on text. No stock photography. No clip art. No emoji.**
The app is flat, engineered and precise — the deck must be too.

## 2.2 Typography

- **Display / headings:** Space Grotesk (fallback: Poppins, Montserrat). Headings are 500–700 weight.
- **Body:** Inter (fallback: Helvetica Neue, Arial). 400–600 weight.
- **Data / labels / eyebrows:** IBM Plex Mono (fallback: Consolas, Courier New).

Rules:
- **Eyebrow labels** above every slide title: uppercase, IBM Plex Mono, 10pt, letter-spaced,
  in signal green — e.g. `PHASE 07 · MODEL EVALUATION`.
- **Slide titles:** Space Grotesk 26–28pt, white, sentence case, one line where possible.
- **Every number is set in IBM Plex Mono.** Numbers are the hero of this deck; they should
  look like machine output, not prose.
- **Subtitles** in muted `#8CA0BE`, 12–13pt, one line, positioned directly under the title.

## 2.3 Layout and alignment (be strict)

- Slide size 16:9 (13.333in × 7.5in).
- **One left margin at 0.62in.** Titles, body, cards, charts and the footer all align to it.
  The right margin mirrors it. Nothing floats at an arbitrary x-position.
- **A single horizontal rule** (1px, `#223049`) sits under the header block and above the
  footer on every content slide — it is what makes the deck feel engineered.
- **Header block:** eyebrow at y=0.40in, title at y=0.63in, subtitle at y=1.16in, rule at
  ~1.50in. Keep this identical on every slide.
- **Footer:** rule at y=7.06in; left side reads `RetainIQ · Signal Ops Console`, right side
  carries the slide number in IBM Plex Mono.
- **Grid:** content cards use an even column grid — 2, 3, 4 or 5 columns, with equal gutters
  (0.16–0.28in). Cards in a row must be **exactly the same height**, top-aligned.
- **Card style:** fill `#111A2C`, 1px border `#223049`, corner radius ~9px. Emphasis cards
  swap the border to green/amber/red at 1.2–1.4px. Never fill a card with a bright colour.
- **KPI tiles:** a row of equal-width tiles, each with a large mono number (20–26pt, in an
  accent colour) above a small uppercase muted label.
- **Progress rail:** a thin bar of **14 ticks** along the bottom of every slide, one per
  development phase. Ticks for completed phases are filled green; the current phase is
  bright green and slightly taller; future phases are hairline grey. This makes the
  development cycle visible at all times.
- **Whitespace:** leave real breathing room. If a slide feels full, cut content — never
  shrink type below 9.5pt or crowd margins.

## 2.4 Imagery and charts

- **Charts** must be generated in the same palette: ink background, hairline gridlines,
  green/amber/red encoding, mono axis labels in muted grey, no chart junk, no 3D, no
  drop shadows. Direct-label data where possible instead of using legends.
- **Product screenshots** are dark and match the theme — place them directly on the ink
  background. **Power BI screenshots are light** — place them inside a white card with a
  hairline border so they read as a deliberate inset, never bare on the dark page.
- Phases 1–8 use data visuals; phases 12–14 use product visuals. Never mix a chart and a
  screenshot on the same slide unless one is clearly secondary and smaller.

---

# 3. SLIDE-BY-SLIDE STRUCTURE (22 slides)

**Slide 1 — Title.** `RetainIQ` in large Space Grotesk, `Signal Ops Console` beneath it in
signal green. Subtitle: "An explainable machine learning platform for customer churn
prediction, risk intelligence and retention strategy." A row of four KPI tiles: `80.34%`
test accuracy · `84.91%` ROC-AUC · `7,043` customers scored · `$92,539` monthly revenue at
risk. Presenter: Logesh S., B.Sc. Data Science. Live app URL and repository URL at the
foot. Signal-strength meter motif, strong.

**Slide 2 — The development cycle.** The map for everything that follows. Draw the
fourteen-phase cycle as a clean horizontal flow in three arcs — Data (1–5), Modelling
(6–8), Intelligence & Delivery (9–14): Business Understanding → Data Collection → Data
Cleaning → EDA → Feature Engineering → Model Development → Model Evaluation → Model
Selection → Explainability → Risk Segmentation → Retention Intelligence → Business
Intelligence → Web Application → Deployment. Colour the arcs blue → amber → green.

**Slide 3 — PHASE 01 · Business Understanding.** What churn costs; why dashboards describe
but cannot predict; why blanket campaigns waste budget. The four questions the platform
answers: *Who will churn? Why? What should we do? What is at stake?* Close with the base
rate trap: predicting "nobody churns" is already 73.46% accurate.

**Slide 4 — What RetainIQ adds to the existing project.** Give this its own slide,
separate from Business Understanding. **Power BI** provides portfolio-level historical
churn, KPIs and segment patterns. **RetainIQ** scores uploaded customer rows with the
deployed Logistic Regression, returns churn probabilities and risk tiers, shows
coefficient-based drivers with rule-based suggestions, and supports bulk review/export.
Make clear the layers are complementary. Campaign uplift has not been measured; revenue
figures are exposure, not confirmed savings.

**Slide 5 — PHASE 02 · Data Collection.** IBM Telco Customer Churn dataset. KPI tiles:
`7,043` customers · `34` raw features · `30` engineered features · `26.54%` churn rate ·
`80/20` stratified split (5,634 train / 1,409 test, both exactly 26.54% churn). A donut
showing 1,869 churned vs 5,174 retained. Feature groups: profile, services, contract &
billing, value.

**Slide 6 — PHASE 03 · Data Cleaning & Preprocessing.** Cleaning steps as a horizontal
flow: raw upload → clean → encode → 30 features. Two callouts: **leakage prevention** —
Churn Score, Churn Reason, Churn Category and Customer Status were removed because they
are only known after a customer has already churned; and **train/serve parity** — the 30
columns are frozen and rebuilt by the same code for every upload.

**Slide 7 — PHASE 04 · Exploratory Data Analysis.** The three findings that shaped
everything: contract type (month-to-month **42.7%** vs one-year 11.3% vs two-year **2.8%** —
a 15× gap), tenure (**47.4%** of first-year customers churn; 55.5% of all churners are in
year one), and service experience (fiber optic 41.9% vs 7.4% with no internet; no tech
support 41.6%; no online security 41.8%). Two charts plus three takeaway cards.

**Slide 8 — PHASE 05 · Feature Engineering.** 3 numeric features (tenure, monthly charges,
total charges) + 27 binary flags = 30. Why flags rather than invented ordering. The frozen
column contract and the unit test that guards it.

**Slide 9 — PHASE 06 · Model Development.** Four algorithms benchmarked: Logistic
Regression, Random Forest, XGBoost, LightGBM. Same features, same split, same scoring code
— only the algorithm changes. One line on why each was chosen. Grouped bar chart of the
four models.

**Slide 10 — PHASE 07 · Model Evaluation.** The measurement slide. Confusion matrix — TN
919, FP 116, FN 161, TP 213 — plus an ROC curve at 84.91% AUC. Then **stability**: 5-fold
stratified cross-validation gives F1 **62.11% ± 2.96%** and accuracy **81.19% ± 1.18%**,
which contains the deployed result — so the single split is not carrying the outcome. State
recall honestly: 56.95%.

**Slide 11 — PHASE 08 · Model Selection — which algorithm.** The benchmark table:

| Model | Accuracy | Precision | Recall | F1 | ROC-AUC |
|---|---|---|---|---|---|
| **Logistic Regression ← selected** | 80.34% | 64.74% | 56.95% | 60.60% | 84.91% |
| LightGBM | 80.06% | 64.49% | 55.35% | 59.57% | 84.71% |
| XGBoost | 79.06% | 61.72% | 55.61% | 58.51% | 83.34% |
| Random Forest | 78.99% | 62.34% | 52.67% | 57.10% | 83.48% |

The selected row is highlighted with a green border. Note that selection was made on F1 and
ROC-AUC, never accuracy alone.

**Slide 12 — PHASE 08 · Model Selection — the operating point.** The strongest slide in the
deck. A threshold sweep table:

| Threshold | Precision | Recall | F1 | Customers flagged |
|---|---|---|---|---|
| 0.30 | 53.04% | 74.60% | 62.00% | 526 |
| **0.40** | 58.06% | **67.38%** | **62.38%** | 434 |
| **0.50 (deployed)** | 64.74% | 56.95% | 60.60% | 329 |
| 0.70 | 76.27% | 24.06% | 36.59% | 118 |

Plus the imbalance experiment as a callout: training with balanced class weights lifts
recall to **77.81%** while ROC-AUC stays essentially unchanged (**84.89** vs **84.90**) —
evidence the model already ranks customers correctly and only the decision point was
discarding churners.

**Slide 13 — PHASE 09 · Explainability.** The profile of the high-risk band versus the
whole portfolio: **100%** month-to-month, **95%** no online security, **94%** no tech
support, **91%** fiber optic, **80%** electronic check, **32%** senior citizens; average
tenure **9.6 months** against 32.4 portfolio-wide. Per-customer explanation is computed as
coefficient × feature value. State SHAP precisely: explored during research, production
explanation uses the linear model's coefficients.

**Slide 14 — PHASE 10 · Risk Segmentation.** Bands: `< 30%` Low · `30–60%` Medium ·
`≥ 60%` High. Sizes: 4,435 / 1,473 / 1,135. Validation against real outcomes: actual churn
**9.5% → 42.0% → 73.0%**, so the bands genuinely order risk. Operational payoff: prioritise
1,135 customers instead of 7,043.

**Slide 15 — PHASE 11 · Retention Intelligence.** 2,608 at-risk customers each received a
specific action: Promote Long-Term Contract 792 · 5% Discount 663 · 15% Discount 658 ·
Welcome Package 430 · Free Online Security 45 · Autopay 9 · Check-in 9 · Premium Support 2.
Revenue framing: **$92,539/month** high risk, **$200,302** including medium. Say "revenue at
risk / exposure" — never "saved."

**Slide 16 — PHASE 12 · Business Intelligence.** The Power BI layer — four pages: executive
overview, customer insights, risk intelligence, retention strategy. Frame it as the
descriptive and diagnostic layer beneath the predictive ML application. Show two light
Power BI screenshots inside white cards.

**Slide 17 — PHASE 13 · Web Application.** The Flask Signal Ops Console. Architecture:
data → preprocessing → Logistic Regression → probability → {risk segmentation,
explainability, retention actions} → {console, API, Power BI}. Console flow: upload → validate
→ score → dashboard KPIs → drill into a customer → download a report. Reports are private
per visitor. Show two dark console screenshots.

**Slide 18 — PHASE 14 · Deployment.** Gunicorn with a gthread worker, 300-second timeout,
preload so the model loads once, periodic worker recycling to avoid memory restarts; Docker
and Procfile provided; hosted on Render. The live URL, stated plainly. Mention cold starts
on the free tier — the instance sleeps when idle.

**Slide 19 — Technology Stack.** Four columns, grouped by role: *Language & Data* (Python
3.11, Pandas, NumPy, SQL Server, SQLAlchemy) · *Machine Learning* (Scikit-learn, Logistic
Regression deployed, Random Forest, XGBoost, LightGBM, Joblib, SHAP) · *Web Application*
(Flask, Jinja2, HTML5, CSS3, vanilla JavaScript, Gunicorn) · *APIs & Delivery* (FastAPI +
Pydantic, Docker, Render, Git/GitHub, Power BI). A callout explains why the stack holds
together: one shared Python package handles preprocessing and scoring for the notebooks, the
app and the API.

**Slide 20 — Responsible AI, limitations & the next cycle.** Three panels. *Limitations:*
historical data from one telecom; 56.95% recall at the deployed threshold; no live
behavioural feed, drift monitoring or automatic retraining; probabilities are not
guarantees. *Fairness note:* senior citizens churn at 41.7%, so the model may systematically
target them — no bias audit has been run. *Cost of errors:* a wasted offer is cheap, a lost
customer is expensive — which is the economic reason the operating threshold is a business
decision, not a technical default. Close by looping back to Phase 02: drift detection and
scheduled retraining begin the next cycle.

**Slide 21 — See it running: demo and project links.** Keep the live application URL in
full (`retainiq-predictive-customer-retention-zq6x.onrender.com`), the live-app QR, the
repository link and its secondary QR. Invite the audience to explore the running app and
source. Keep the project takeaway, presenter credit and headline metrics on this
resources/demo slide.

**Slide 22 — Thank You.** A separate final slide with a clear **Thank You** heading and
**Questions & discussion** invitation, the closing project message, presenter credit,
headline metrics and a full-strength signal meter. Do not repeat the QR/resource panel.

---

# 4. REQUIRED SPEAKER NOTES

For every slide, write speaker notes of roughly 45–60 seconds of natural spoken delivery
(total ~11 minutes 55 seconds), plus a timing cue. Notes must:
- open with a spoken transition from the previous slide, not a topic label;
- name the single most important number and say what it means;
- where relevant, include a prepared answer to the obvious question.

---

# 5. HARD CONSTRAINTS — factual integrity

This deck will be presented to examiners who may open the repository. Violating any of
these is worse than a bland slide.

1. **Use only the numbers given above.** Do not invent, round differently, or embellish.
2. **Never say "revenue saved."** Always "revenue at risk" or "revenue exposure."
3. **Never claim SHAP is the production explainability method.** It was research;
   coefficients explain the deployed model.
4. **Never praise or select a model on accuracy alone.**
5. **Never quote performance without recall.**
6. **Do not list Plotly** in the technology stack — it is not used in the codebase.
7. **Do not conflate** the 26.54% historical churn rate with the 16.1% share of customers in
   the high-risk band. They are different measurements.
8. **XGBoost hyperparameters** should be read from the project's training notebook before
   being stated anywhere; if unavailable, omit them rather than guess.

---

# 6. DELIVERABLE

A **22-slide 16:9 presentation**, ending on a separate Thank You slide. Do not add appendix slides to this presentation.

Deliver the deck as an editable file, with speaker notes embedded, and confirm for each
slide that the alignment grid, the margin, and the 14-phase progress rail are consistent.
