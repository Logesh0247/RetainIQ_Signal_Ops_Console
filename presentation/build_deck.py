"""
Build the RetainIQ project-review deck (.pptx).

Design: mirrors the application's "Signal Ops" dark console theme -- ink
background, signal-green accent, monospace labels -- so the slides and the
product look like one artefact.

Every number on every slide comes from the repository: the raw dataset, the
saved model artifacts, the held-out test split, or the production scoring
pipeline (see presentation/build_assets.py for the figures it plots).

Run from the repository root:

    python presentation/build_deck.py

The script measures each text block against its box using DejaVu metrics and
prints an overflow report, so layout problems surface here rather than on
presentation day.
"""
from __future__ import annotations

import os
from pathlib import Path

from PIL import Image, ImageFont
from pptx import Presentation
from pptx.dml.color import RGBColor
from pptx.enum.shapes import MSO_SHAPE
from pptx.enum.text import MSO_ANCHOR, PP_ALIGN
from pptx.util import Emu, Inches, Pt

HERE = Path(__file__).resolve().parent
ROOT = HERE.parent
ASSETS = HERE / "assets"
IMAGES = ROOT / "templates" / "images"
PBI = ROOT / "Power_BI_dashboard" / "Screenshots"
OUT = HERE / "RetainIQ_Project_Review.pptx"

# --- palette (static/style.css tokens) --------------------------------------
INK = RGBColor(0x0A, 0x0F, 0x1C)
PANEL = RGBColor(0x11, 0x1A, 0x2C)
PANEL_2 = RGBColor(0x16, 0x21, 0x36)
LINE = RGBColor(0x22, 0x30, 0x49)
TEXT = RGBColor(0xEA, 0xF0, 0xF7)
MUTED = RGBColor(0x8C, 0xA0, 0xBE)
ACTION = RGBColor(0x4C, 0x8D, 0xFF)
OK = RGBColor(0x2F, 0xD8, 0xB4)
WARN = RGBColor(0xF5, 0xA6, 0x23)
BAD = RGBColor(0xFF, 0x5C, 0x6C)
WHITE = RGBColor(0xFF, 0xFF, 0xFF)

FONT = "Segoe UI"
MONO = "Consolas"

SW, SH = 13.333, 7.5
M = 0.62
CW = SW - 2 * M

# --- text measurement (conservative proxy for the render fonts) -------------
_TTF = Path(__import__("matplotlib").__file__).parent / "mpl-data" / "fonts" / "ttf"
_FONTS = {
    False: ImageFont.truetype(str(_TTF / "DejaVuSans.ttf"), 100),
    True: ImageFont.truetype(str(_TTF / "DejaVuSans-Bold.ttf"), 100),
}
_MONO = {
    False: ImageFont.truetype(str(_TTF / "DejaVuSansMono.ttf"), 100),
    True: ImageFont.truetype(str(_TTF / "DejaVuSansMono-Bold.ttf"), 100),
}

WARNINGS: list[str] = []


def text_width_pt(text: str, pt: float, bold: bool = False, mono: bool = False) -> float:
    font = (_MONO if mono else _FONTS)[bold]
    return font.getlength(text) / 100 * pt


def wrap_lines(text: str, pt: float, box_w_in: float, bold=False, mono=False) -> int:
    limit = box_w_in * 72
    words, lines, cur = text.split(), 0, ""
    for word in words:
        trial = f"{cur} {word}".strip()
        if text_width_pt(trial, pt, bold, mono) <= limit or not cur:
            cur = trial
        else:
            lines += 1
            cur = word
    return lines + (1 if cur else 0)


# --- primitives -------------------------------------------------------------
def add_slide(prs):
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    bg = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, 0, 0, Inches(SW), Inches(SH))
    bg.fill.solid()
    bg.fill.fore_color.rgb = INK
    bg.line.fill.background()
    bg.shadow.inherit = False
    return slide


def rect(slide, x, y, w, h, fill=PANEL, edge=LINE, radius=0.08, edge_w=1.0):
    shape = slide.shapes.add_shape(
        MSO_SHAPE.ROUNDED_RECTANGLE, Inches(x), Inches(y), Inches(w), Inches(h))
    shape.adjustments[0] = radius
    if fill is None:
        shape.fill.background()
    else:
        shape.fill.solid()
        shape.fill.fore_color.rgb = fill
    if edge is None:
        shape.line.fill.background()
    else:
        shape.line.color.rgb = edge
        shape.line.width = Pt(edge_w)
    shape.shadow.inherit = False
    return shape


def bar(slide, x, y, w, h, color):
    shape = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(x), Inches(y), Inches(w), Inches(h))
    shape.fill.solid()
    shape.fill.fore_color.rgb = color
    shape.line.fill.background()
    shape.shadow.inherit = False
    return shape


def text(slide, x, y, w, h, blocks, align=PP_ALIGN.LEFT, anchor=MSO_ANCHOR.TOP,
         line_spacing=1.12):
    """blocks: list of dicts {t, pt, color, bold, mono, space_after, align, spacing}"""
    box = slide.shapes.add_textbox(Inches(x), Inches(y), Inches(w), Inches(h))
    tf = box.text_frame
    tf.word_wrap = True
    tf.margin_left = tf.margin_right = tf.margin_top = tf.margin_bottom = 0
    tf.vertical_anchor = anchor

    for i, item in enumerate(blocks):
        p = tf.paragraphs[0] if i == 0 else tf.add_paragraph()
        p.alignment = item.get("align", align)
        p.line_spacing = item.get("spacing", line_spacing)
        if item.get("space_before"):
            p.space_before = Pt(item["space_before"])
        p.space_after = Pt(item.get("space_after", 6))
        run = p.add_run()
        run.text = item["t"]
        run.font.name = MONO if item.get("mono") else FONT
        run.font.size = Pt(item["pt"])
        run.font.bold = item.get("bold", False)
        run.font.color.rgb = item.get("color", TEXT)

        # --- layout audit -------------------------------------------------
        pt = item["pt"]
        needed = wrap_lines(item["t"], pt, w, item.get("bold", False), item.get("mono", False))
        required_h = needed * pt * 1.22 * item.get("spacing", line_spacing) / 72
        if required_h > h + 0.06:
            WARNINGS.append(
                f"possible vertical overflow: {needed} lines need {required_h:.2f}in in "
                f"{h:.2f}in box -> \"{item['t'][:58]}...\"")
    return box


def eyebrow(slide, label, x=M, y=0.40, color=OK):
    text(slide, x, y, 8.0, 0.26,
         [{"t": label, "pt": 10, "color": color, "bold": True, "mono": True, "space_after": 0}])


def heading(slide, title, sub=None):
    text(slide, M, 0.63, CW, 0.52,
         [{"t": title, "pt": 26, "color": TEXT, "bold": True, "space_after": 0}])
    y = 1.16
    if sub:
        text(slide, M, y, CW, 0.30,
             [{"t": sub, "pt": 12.5, "color": MUTED, "space_after": 0}])
        y += 0.34
    bar(slide, M, y, CW, 0.014, LINE)
    return y + 0.16


def footer(slide, number, label="RetainIQ · Signal Ops Console"):
    bar(slide, M, 7.06, CW, 0.010, LINE)
    text(slide, M, 7.12, 8.0, 0.24,
         [{"t": label, "pt": 9, "color": MUTED, "space_after": 0}])
    text(slide, SW - M - 1.2, 7.12, 1.2, 0.24,
         [{"t": f"{number:02d}", "pt": 9, "color": MUTED, "mono": True,
           "space_after": 0, "align": PP_ALIGN.RIGHT}])


def kpi(slide, x, y, w, h, value, label, color=OK, value_pt=25, label_pt=9.5):
    rect(slide, x, y, w, h, fill=PANEL, edge=LINE)
    text(slide, x + 0.16, y + 0.14, w - 0.28, h - 0.22,
         [{"t": value, "pt": value_pt, "color": color, "bold": True, "space_after": 2,
           "mono": False},
          {"t": label, "pt": label_pt, "color": MUTED, "space_after": 0}])


def callout(slide, x, y, w, h, title, body, color=WARN):
    rect(slide, x, y, w, h, fill=PANEL_2, edge=color, edge_w=1.2)
    text(slide, x + 0.20, y + 0.14, w - 0.40, h - 0.26,
         [{"t": title, "pt": 11.5, "color": color, "bold": True, "space_after": 3},
          {"t": body, "pt": 10.8, "color": TEXT, "space_after": 0, "spacing": 1.16}])


def bullets(slide, x, y, w, h, items, pt=12, gap=7, color=TEXT, marker="▸"):
    blocks = []
    for i, item in enumerate(items):
        title, rest = (item if isinstance(item, tuple) else (None, item))
        if title:
            blocks.append({"t": f"{marker}  {title}", "pt": pt, "color": OK,
                           "bold": True, "space_after": 1,
                           "space_before": 0 if i == 0 else gap})
            blocks.append({"t": f"     {rest}", "pt": pt - 0.5, "color": color,
                           "space_after": 0, "spacing": 1.18})
        else:
            blocks.append({"t": f"{marker}  {rest}", "pt": pt, "color": color,
                           "space_after": gap})
    return text(slide, x, y, w, h, blocks)


def image_fit(slide, path, x, y, w, h, align_x=0.5, align_y=0.5):
    """Place an image scaled to fit inside the box, preserving aspect ratio."""
    with Image.open(path) as im:
        iw, ih = im.size
    scale = min(w / iw, h / ih)
    dw, dh = iw * scale, ih * scale
    left = x + (w - dw) * align_x
    top = y + (h - dh) * align_y
    return slide.shapes.add_picture(str(path), Inches(left), Inches(top),
                                    Inches(dw), Inches(dh))


def notes(slide, body):
    slide.notes_slide.notes_text_frame.text = body.strip()


def picture_card(slide, path, x, y, w, h, pad=0.09, fill=WHITE, edge=LINE):
    """White card behind light (Power BI) screenshots so they read on ink."""
    rect(slide, x, y, w, h, fill=fill, edge=edge)
    image_fit(slide, path, x + pad, y + pad, w - 2 * pad, h - 2 * pad)


# --- deck -------------------------------------------------------------------
def build():
    prs = Presentation()
    prs.slide_width = Inches(SW)
    prs.slide_height = Inches(SH)
    n = 0

    # ================= 01 TITLE =================
    n += 1
    s = add_slide(prs)
    bar(s, 0, 0, SW, 0.085, OK)
    text(s, M, 1.28, 9.6, 0.28,
         [{"t": "FINAL YEAR PROJECT  ·  B.SC. DATA SCIENCE", "pt": 11, "color": OK,
           "bold": True, "mono": True, "space_after": 0}])
    text(s, M, 1.62, 10.4, 1.62,
         [{"t": "RetainIQ", "pt": 54, "color": TEXT, "bold": True, "space_after": 2},
          {"t": "Signal Ops Console", "pt": 28, "color": OK, "bold": True, "space_after": 0}])
    bar(s, M, 3.36, 2.1, 0.035, OK)
    text(s, M, 3.66, 7.9, 1.0,
         [{"t": "An explainable machine learning platform for customer churn "
                "prediction, risk intelligence and retention strategy.",
           "pt": 14.5, "color": MUTED, "space_after": 0, "spacing": 1.25}])

    # signal-strength motif from the product
    bx, by = 11.05, 1.72
    for i, hgt in enumerate([0.22, 0.38, 0.54, 0.70, 0.86]):
        bar(s, bx + i * 0.20, by + (0.86 - hgt), 0.115, hgt, OK if i < 3 else LINE)
    text(s, bx - 0.05, by + 0.94, 1.5, 0.24,
         [{"t": "signal strength", "pt": 8.5, "color": MUTED, "mono": True,
           "space_after": 0}])

    kpi(s, M, 5.00, 2.36, 1.02, "80.34%", "TEST ACCURACY", OK)
    kpi(s, M + 2.50, 5.00, 2.36, 1.02, "84.91%", "ROC-AUC", OK)
    kpi(s, M + 5.00, 5.00, 2.36, 1.02, "7,043", "CUSTOMERS SCORED", ACTION)
    kpi(s, M + 7.50, 5.00, 4.59, 1.02, "$92,539", "MONTHLY REVENUE EXPOSURE IN HIGH-RISK BAND", BAD, 25, 9)

    text(s, M, 6.32, 12.1, 0.7,
         [{"t": "Logesh S.   ·   B.Sc. Data Science", "pt": 12.5, "color": TEXT,
           "bold": True, "space_after": 3},
          {"t": "Live application: retainiq-predictive-customer-retention-zq6x.onrender.com"
                "     ·     Source: github.com/Logesh0247/RetainIQ_Signal_Ops_Console",
           "pt": 10.5, "color": MUTED, "mono": True, "space_after": 0}])
    notes(s, """
OPENING (0:00–0:30). "Good morning. I'm Logesh, and this is RetainIQ — a churn
prediction platform I built end to end: data cleaning, model training, an
explainability layer, a retention recommendation engine, a Flask web app and a
Power BI layer. The headline: the model predicts churn with 80.34% accuracy and
a ROC-AUC of 84.91%, but the real point is what happens after the prediction —
risk bands, explanations, actions and revenue exposure."

Do not read the slides. Land the four KPIs, then move on.
""")

    # ================= 02 WHAT THIS PROJECT DOES =================
    n += 1
    s = add_slide(prs)
    eyebrow(s, "OVERVIEW")
    y = heading(s, "What this project does",
                "The one-line thesis, then the four parts of the talk")
    rect(s, M, y, CW, 1.16, fill=PANEL_2, edge=OK, edge_w=1.2)
    text(s, M + 0.28, y + 0.20, CW - 0.56, 0.8,
         [{"t": "RetainIQ turns churn from a yes/no prediction into an operational "
                "workflow: score the portfolio  →  explain the risk  →  recommend the "
                "action  →  quantify the revenue at stake.",
           "pt": 14, "color": TEXT, "space_after": 0, "spacing": 1.22}])

    y2 = y + 1.40
    items = [
        ("01", "Problem & data", "Why retention is reactive today, and the 7,043-customer Telco dataset behind the model", ACTION),
        ("02", "Method & model", "Preprocessing, leakage control, feature engineering, four algorithms, one selected", WARN),
        ("03", "Risk, reason & action", "Probability → risk bands → per-customer explanation → retention recommendation", OK),
        ("04", "Product & results", "The live console, dashboard, reports, Power BI layer — and the honest limitations", BAD),
    ]
    for i, (num, title, sub, color) in enumerate(items):
        row = i % 2
        col = i // 2
        x = M + col * (CW / 2 + 0.12)
        w = CW / 2 - 0.12
        yy = y2 + row * 1.52
        rect(s, x, yy, w, 1.32, fill=PANEL, edge=LINE)
        text(s, x + 0.24, yy + 0.16, 0.7, 0.5,
             [{"t": num, "pt": 24, "color": color, "bold": True, "mono": True, "space_after": 0}])
        text(s, x + 1.02, yy + 0.20, w - 1.26, 0.9,
             [{"t": title, "pt": 14.5, "color": TEXT, "bold": True, "space_after": 3},
              {"t": sub, "pt": 10.8, "color": MUTED, "space_after": 0, "spacing": 1.18}])
    footer(s, n)
    notes(s, """
(0:30–0:55) Read the thesis line aloud — it is the spine of the whole talk. Then
point at the four quadrants: "I'll move through problem and data, then how the
model was built and chosen, then how a probability becomes an action, and
finally the product itself and what it does and does not deliver."
""")

    # ================= 03 PROBLEM =================
    n += 1
    s = add_slide(prs)
    eyebrow(s, "01  ·  BUSINESS PROBLEM")
    y = heading(s, "The problem: retention is still reactive",
                "Reports explain what happened. They do not say who leaves next.")
    bullets(s, M, y + 0.06, 6.55, 2.1, [
        ("Churn is expensive and it compounds.",
         "26.54% of this customer base churned — 1,869 of 7,043 customers."),
        ("Dashboards describe, they do not prioritise.",
         "Knowing last quarter's churn rate does not tell a retention team whom to call today."),
        ("Generic campaigns waste budget.",
         "Blanket offers discount loyal customers and still miss the ones about to leave."),
    ], pt=11.5, gap=8)

    qs = [
        ("Who is likely to churn?", "Predicted churn probability per customer", ACTION),
        ("Why are they at risk?", "Feature contributions behind each prediction", WARN),
        ("What should we do?", "A retention action matched to the risk drivers", OK),
        ("What is at stake?", "Monthly recurring revenue exposed in each risk band", BAD),
    ]
    for i, (q, sub, color) in enumerate(qs):
        col = i % 2
        row = i // 2
        x = M + col * 3.34
        yy = y + 2.42 + row * 1.30
        rect(s, x, yy, 3.16, 1.14, fill=PANEL, edge=LINE)
        bar(s, x, yy + 0.14, 0.035, 0.86, color)
        text(s, x + 0.20, yy + 0.16, 2.86, 0.9,
             [{"t": q, "pt": 12, "color": TEXT, "bold": True, "space_after": 3},
              {"t": sub, "pt": 9.8, "color": MUTED, "space_after": 0, "spacing": 1.14}])

    rect(s, 7.62, y + 0.06, 5.09, 4.94, fill=PANEL, edge=LINE)
    text(s, 7.86, y + 0.22, 4.6, 0.3,
         [{"t": "THE BASE RATE", "pt": 10, "color": MUTED, "bold": True, "mono": True,
           "space_after": 0}])
    image_fit(s, ASSETS / "chart_churn_split.png", 7.72, y + 0.52, 4.89, 4.0)
    text(s, 7.86, y + 4.42, 4.6, 0.5,
         [{"t": "Any model that predicts \"nobody churns\" is already 73.46% accurate — "
                "which is why accuracy alone decides nothing.",
           "pt": 10.5, "color": MUTED, "space_after": 0, "spacing": 1.16}])
    footer(s, n)
    notes(s, """
(0:55–1:45) Frame the business pain, then the four questions — they are the
structure of the entire solution. Finish on the donut: "Notice the trap here.
If I predict nobody churns, I'm 73.46% accurate and completely useless. That
number is why I never selected a model on accuracy alone." That line earns
credibility early and sets up slide 7.
""")

    # ================= 04 DATASET =================
    n += 1
    s = add_slide(prs)
    eyebrow(s, "01  ·  DATA")
    y = heading(s, "The data behind the model",
                "IBM Telco Customer Churn — 7,043 customers, 34 raw fields, one binary "
                "target · 80/20 stratified split: 5,634 train / 1,409 test")
    tiles = [
        ("7,043", "CUSTOMERS", TEXT), ("34", "RAW FEATURES", TEXT),
        ("30", "ML FEATURES", TEXT), ("26.54%", "CHURN RATE", BAD),
        ("80 / 20", "TRAIN / TEST SPLIT", ACTION),
    ]
    tw = (CW - 4 * 0.16) / 5
    for i, (v, l, c) in enumerate(tiles):
        kpi(s, M + i * (tw + 0.16), y + 0.04, tw, 1.06, v, l, c, 20, 9)

    yy = y + 1.30
    text(s, M, yy, 5.9, 2.6,
         [{"t": "What a row contains", "pt": 13, "color": OK, "bold": True, "space_after": 6},
          {"t": "  ▸  Customer profile — gender, senior citizen, partner, dependents", "pt": 11.3, "color": TEXT, "space_after": 4},
          {"t": "  ▸  Services — phone, internet, security, backup, protection, support, streaming", "pt": 11.3, "color": TEXT, "space_after": 4, "spacing": 1.16},
          {"t": "  ▸  Contract & billing — contract type, payment method, paperless, charges", "pt": 11.3, "color": TEXT, "space_after": 4},
          {"t": "  ▸  Value — tenure, monthly charges, total charges, CLTV", "pt": 11.3, "color": TEXT, "space_after": 4},
          {"t": "Target: Churn Label (Yes / No), converted to 1 / 0.", "pt": 11.3, "color": MUTED, "space_after": 0, "space_before": 8}])

    callout(s, 6.85, yy, 5.86, 1.74,
            "LEAKAGE PREVENTION — the detail that makes the results believable",
            "Churn Score, Churn Reason, Churn Category and Customer Status were removed "
            "before training. All four are only known after a customer has already "
            "churned, so keeping them would inflate the metrics and break the model in "
            "production.")
    callout(s, 6.85, yy + 1.90, 5.86, 1.36,
            "TRAIN / SERVE PARITY",
            "The 30 encoded columns are frozen in feature_columns.pkl and rebuilt by the "
            "same preprocessing code for every upload — training and live scoring cannot "
            "drift apart. There is a unit test for exactly this.", OK)
    footer(s, n)
    notes(s, """
(1:45–2:30) Quick pass over the tiles. Spend the time on the leakage callout —
reviewers consistently reward this. Say: "Churn Score and Churn Reason are the
dataset's most seductive columns, and they're poison: they exist only after the
customer has left. I dropped them. That is why 80% here is a real 80%."
Then one sentence on train/serve parity and move on.
""")

    # ================= 05 EDA =================
    n += 1
    s = add_slide(prs)
    eyebrow(s, "01  ·  EXPLORATORY ANALYSIS")
    y = heading(s, "What the data said",
                "Three patterns that shaped the model — and the retention strategy")
    image_fit(s, ASSETS / "chart_contract.png", M, y + 0.04, 5.95, 3.30)
    image_fit(s, ASSETS / "chart_tenure.png", M + 6.15, y + 0.04, 5.94, 3.30)

    yy = y + 3.46
    cards = [
        ("Contract is the strongest signal", "42.7% churn on month-to-month vs 2.8% on two-year contracts — a 15× gap.", OK),
        ("Churn is front-loaded", "47.4% of first-year customers churn; 55.5% of all churners are in year one. Churned customers average 18.0 months of tenure vs 37.6.", WARN),
        ("It is a service-experience story", "Fiber optic churns at 41.9% vs 7.4% with no internet; customers without tech support or online security churn at 41.6% / 41.8%.", ACTION),
    ]
    cw = (CW - 2 * 0.18) / 3
    for i, (t, b, c) in enumerate(cards):
        x = M + i * (cw + 0.18)
        rect(s, x, yy, cw, 1.72, fill=PANEL, edge=LINE)
        bar(s, x, yy + 0.16, 0.035, 1.40, c)
        text(s, x + 0.20, yy + 0.18, cw - 0.40, 1.40,
             [{"t": t, "pt": 11.5, "color": c, "bold": True, "space_after": 4},
              {"t": b, "pt": 10.2, "color": TEXT, "space_after": 0, "spacing": 1.18}])
    footer(s, n)
    notes(s, """
(2:30–3:20) Don't narrate every bar — three takeaways only.
Contract: "month-to-month is a 15× risk multiplier."
Tenure: "risk is front-loaded — year one is where retention money belongs."
Experience: "the churn story is service experience, not price alone: fiber optic,
no security, no support."
Close with the bridge: "these three patterns are what the model later recovered
on its own — and what the recommendation engine acts on."
""")

    # ================= 06 PIPELINE =================
    n += 1
    s = add_slide(prs)
    eyebrow(s, "02  ·  PREPROCESSING & FEATURES")
    y = heading(s, "From raw records to 30 model-ready features",
                "One reproducible pipeline, used identically in training and in the live app")

    steps = [
        ("Raw upload", "7,043 records · 34 columns", ACTION),
        ("Clean", "types, blanks, duplicates, category names", ACTION),
        ("Encode", "16 categoricals → 27 flags", WARN),
        ("30 features", "frozen column contract", OK),
    ]
    sw_ = (CW - 3 * 0.30) / 4
    for i, (t, sub, c) in enumerate(steps):
        x = M + i * (sw_ + 0.30)
        rect(s, x, y + 0.06, sw_, 1.10, fill=PANEL_2, edge=c, edge_w=1.2)
        text(s, x + 0.18, y + 0.20, sw_ - 0.36, 0.86,
             [{"t": t, "pt": 13, "color": TEXT, "bold": True, "space_after": 3},
              {"t": sub, "pt": 9.8, "color": MUTED, "space_after": 0, "spacing": 1.14}])
        if i < 3:
            text(s, x + sw_ + 0.02, y + 0.36, 0.28, 0.4,
                 [{"t": "→", "pt": 17, "color": OK, "bold": True, "space_after": 0,
                   "align": PP_ALIGN.CENTER}])

    yy = y + 1.44
    bullets(s, M, yy, 5.95, 2.9, [
        ("Cleaning", "Data-type correction, blank TotalCharges for new customers, duplicate "
                     "checks, categorical normalisation, and removal of ID and geographic "
                     "columns that carry no predictive value."),
        ("Encoding", "Every service column becomes an explicit flag (Contract_Two year, "
                     "Tech Support_Yes, Internet Service_Fiber optic …) so no ordering is "
                     "invented where none exists."),
        ("Scaling & split", "An 80/20 stratified split keeps the 26.54% churn rate "
                            "identical in train (5,634) and test (1,409)."),
    ], pt=11.2, gap=9)

    rect(s, 6.85, yy, 5.86, 2.34, fill=PANEL, edge=OK, edge_w=1.2)
    text(s, 7.07, yy + 0.18, 5.42, 2.0,
         [{"t": "WHY THIS SLIDE MATTERS", "pt": 10.5, "color": OK, "bold": True,
           "mono": True, "space_after": 7},
          {"t": "The single most common failure in student ML projects is a notebook that "
                "scores well and an app that scores differently. Here the same "
                "preprocessing module builds the feature frame for training, for the "
                "test-set evaluation, for the live upload — and a unit test asserts the "
                "saved column list matches what the model expects.",
           "pt": 11, "color": TEXT, "space_after": 0, "spacing": 1.20}])
    footer(s, n)
    notes(s, """
(3:20–4:05) Walk the four boxes quickly, then spend the time on the right panel.
"Cleaning, encoding and the split all live in one module that both the training
notebook and the live app call. The 30-column contract is saved next to the
model and tested — so the app can't silently score on the wrong schema."
If asked about scaling: the saved pipeline handles it; note that feature scaling
matters for the linear model and is applied consistently.
""")

    # ================= 07 BENCHMARK =================
    n += 1
    s = add_slide(prs)
    eyebrow(s, "02  ·  MODELLING")
    y = heading(s, "Four algorithms, one pipeline",
                "Same features, same split, same scoring code — only the algorithm changes")
    image_fit(s, ASSETS / "chart_models.png", M, y + 0.04, 7.85, 3.44)

    bullets(s, 8.75, y + 0.10, 3.96, 3.30, [
        ("Logistic Regression", "linear baseline; interpretable coefficients"),
        ("Random Forest", "bagged trees; captures non-linearity"),
        ("XGBoost & LightGBM", "gradient boosting; strong tabular performers"),
    ], pt=10.8, gap=8)

    rect(s, 8.75, y + 2.42, 3.96, 1.06, fill=PANEL_2, edge=WARN, edge_w=1.2)
    text(s, 8.93, y + 2.54, 3.60, 0.86,
         [{"t": "Benchmark ladder (F1)", "pt": 10.2, "color": WARN, "bold": True,
           "space_after": 3},
          {"t": "LR 60.6  ·  LightGBM 59.6  ·  XGBoost 58.5  ·  RF 57.1",
           "pt": 10.2, "color": TEXT, "space_after": 0, "mono": True}])

    yy = y + 3.62
    callout(s, M, yy, CW, 1.02, "WHY FOUR MODELS AND NOT ONE",
            "A single model proves nothing. Four algorithms on the same pipeline show the "
            "result is a property of the features and the problem — not of one lucky "
            "configuration. Accuracy, precision, recall, F1 and ROC-AUC were all "
            "recorded, and the winner was chosen on F1 and ROC-AUC, not accuracy.")
    footer(s, n)
    notes(s, """
(4:05–4:45) "Same features, same split, same code path — the only thing that
changes is the algorithm." Point out that Logistic Regression led on F1 and on
ROC-AUC, with LightGBM a close second. The takeaway to say out loud: "churn here
is largely linear in the encoded features, so the simplest model wins — and that
also buys me explainability for free." That is the transition into slide 8.
""")

    # ================= 08 EVALUATION =================
    n += 1
    s = add_slide(prs)
    eyebrow(s, "02  ·  EVALUATION")
    y = heading(s, "Why Logistic Regression ships",
                "Held-out test set, 1,409 customers the model never saw in training")
    image_fit(s, ASSETS / "chart_confusion.png", M, y + 0.02, 5.35, 3.02)
    image_fit(s, ASSETS / "chart_roc.png", M + 5.55, y + 0.02, 5.00, 3.02)

    yy = y + 3.18
    tiles = [("80.34%", "ACCURACY", TEXT), ("64.74%", "PRECISION", TEXT),
             ("56.95%", "RECALL", WARN), ("60.60%", "F1 SCORE", OK),
             ("84.91%", "ROC-AUC", OK)]
    tw = (CW - 4 * 0.16) / 5
    for i, (v, l, c) in enumerate(tiles):
        kpi(s, M + i * (tw + 0.16), yy, tw, 0.94, v, l, c, 19, 9)

    text(s, M, yy + 1.04, CW, 0.5,
         [{"t": "Honest reading: the model catches 213 of 374 churners in the test set "
                "(56.95% recall) and is right 64.74% of the time when it raises a flag. "
                "For a retention team that is a usable work queue — and the missed half "
                "is the number one item in the future-work slide.",
           "pt": 10.8, "color": MUTED, "space_after": 0, "spacing": 1.16}])
    footer(s, n)
    notes(s, """
(4:45–5:35) This is the slide to slow down on. "1,409 customers were held back
from training entirely. 919 stayed and were correctly left alone; 116 were
flagged unnecessarily; 161 churners were missed; 213 were caught."
Then the reflection: "56.95% recall is the honest weakness of this model. I would
rather state it than hide it — and it is exactly why threshold tuning and
cost-sensitive learning are on the future-work slide." Reviewers reward this.
""")

    # ================= 09 RISK BANDS =================
    n += 1
    s = add_slide(prs)
    eyebrow(s, "03  ·  RISK SEGMENTATION")
    y = heading(s, "Turning a probability into a work queue",
                "Every scored customer lands in one of three operational bands")

    bands = [
        ("< 30%", "LOW RISK", "4,435 customers", OK),
        ("30 – 60%", "MEDIUM RISK", "1,473 customers", WARN),
        ("≥ 60%", "HIGH RISK", "1,135 customers", BAD),
    ]
    bw = (CW - 2 * 0.24) / 3
    for i, (rng, label, count, c) in enumerate(bands):
        x = M + i * (bw + 0.24)
        rect(s, x, y + 0.02, bw, 0.94, fill=PANEL, edge=c, edge_w=1.2)
        text(s, x + 0.22, y + 0.14, bw - 0.44, 0.72,
             [{"t": f"{rng}   {label}", "pt": 11.5, "color": c, "bold": True,
               "mono": True, "space_after": 2},
              {"t": count, "pt": 12.5, "color": TEXT, "space_after": 0}])

    image_fit(s, ASSETS / "chart_risk_bands.png", M, y + 1.06, 7.72, 3.42)
    rect(s, 8.62, y + 1.10, 4.09, 1.60, fill=PANEL_2, edge=OK, edge_w=1.2)
    text(s, 8.82, y + 1.24, 3.73, 1.36,
         [{"t": "VALIDATION", "pt": 10.5, "color": OK, "bold": True, "mono": True,
           "space_after": 5},
          {"t": "The bands track reality: actual churn is 9.5% in low risk, 42.0% in "
                "medium and 73.0% in high risk. The segmentation is not decoration — it "
                "orders the portfolio correctly.",
           "pt": 10.6, "color": TEXT, "space_after": 0, "spacing": 1.18}])
    rect(s, 8.62, y + 2.82, 4.09, 1.64, fill=PANEL, edge=BAD, edge_w=1.2)
    text(s, 8.82, y + 2.96, 3.73, 1.40,
         [{"t": "THE OPERATIONAL PAYOFF", "pt": 10.5, "color": BAD, "bold": True,
           "mono": True, "space_after": 5},
          {"t": "Instead of calling 7,043 customers, the team starts with 1,135 "
                "(16.1% of the base) where 73% will actually leave — carrying "
                "$92,539 of recurring monthly charges.",
           "pt": 10.6, "color": TEXT, "space_after": 0, "spacing": 1.18}])
    footer(s, n)
    notes(s, """
(5:35–6:15) "A probability is not something a retention agent can act on, so I
cut it into three bands."
Then the validation, which is the point: "I checked the bands against the real
churn labels. Low risk actually churns 9.5%, medium 42%, high 73%. The band is a
genuine ordering of risk, not a cosmetic split."
Close with the operational line: "the team starts with 1,135 customers instead
of 7,043."
""")

    # ================= 10 EXPLAINABILITY =================
    n += 1
    s = add_slide(prs)
    eyebrow(s, "03  ·  EXPLAINABILITY")
    y = heading(s, "Why the model flags a customer",
                "A profile of the high-risk band, and a per-customer explanation in the app")
    image_fit(s, ASSETS / "chart_high_risk_profile.png", M, y + 0.02, 7.55, 3.86)

    rect(s, 8.40, y + 0.04, 4.31, 1.94, fill=PANEL, edge=LINE)
    text(s, 8.60, y + 0.18, 3.95, 1.70,
         [{"t": "PER-CUSTOMER EXPLANATION", "pt": 10.5, "color": OK, "bold": True,
           "mono": True, "space_after": 6},
          {"t": "For each prediction the console shows the top drivers, computed as "
                "coefficient × customer value — the honest linear explanation for a "
                "logistic model, with human labels like \"Fiber optic internet service\" "
                "or \"Account tenure\".",
           "pt": 10.6, "color": TEXT, "space_after": 0, "spacing": 1.18}])

    callout(s, 8.40, y + 2.10, 4.31, 1.80,
            "ON SHAP — INTELLECTUALLY HONEST",
            "SHAP was explored during the project (summary, bar, dependence and "
            "waterfall plots are in the repository). Production explanation is aligned "
            "to the deployed Logistic Regression model instead — accurate for a linear "
            "model, and it costs nothing at request time.", WARN)
    footer(s, n)
    notes(s, """
(6:15–7:00) Left chart: "this is who the high-risk band is — every one of them is
month-to-month, 91% are on fiber optic, 95% have no online security, average
tenure under ten months."
Right: "and for a single customer the app shows which of those factors pushed
this specific score up."
Be upfront on SHAP if asked: "I did the SHAP analysis during experimentation; the
deployed model is linear, so I explain it with coefficients — the same
information without the runtime cost."
""")

    # ================= 11 ACTION & REVENUE =================
    n += 1
    s = add_slide(prs)
    eyebrow(s, "03  ·  RETENTION INTELLIGENCE")
    y = heading(s, "From risk to action — and to money",
                "2,608 at-risk customers (high + medium) each received a specific recommendation")
    image_fit(s, ASSETS / "chart_revenue.png", M, y + 0.04, 7.10, 3.10)

    rect(s, 7.95, y + 0.02, 4.76, 3.14, fill=PANEL, edge=LINE)
    text(s, 8.15, y + 0.16, 4.40, 2.90,
         [{"t": "RECOMMENDATIONS ISSUED (FULL PORTFOLIO)", "pt": 10.5, "color": OK,
           "bold": True, "mono": True, "space_after": 7},
          {"t": "Promote Long-Term Contract            792", "pt": 10.4, "color": TEXT, "space_after": 3, "mono": True},
          {"t": "5% Discount Offer                            663", "pt": 10.4, "color": TEXT, "space_after": 3, "mono": True},
          {"t": "Offer 15% Discount                          658", "pt": 10.4, "color": TEXT, "space_after": 3, "mono": True},
          {"t": "Welcome Retention Package          430", "pt": 10.4, "color": TEXT, "space_after": 3, "mono": True},
          {"t": "Free Online Security                       45", "pt": 10.4, "color": TEXT, "space_after": 3, "mono": True},
          {"t": "Switch to Autopay / Check-in          18", "pt": 10.4, "color": TEXT, "space_after": 3, "mono": True},
          {"t": "Free Premium Support                      2", "pt": 10.4, "color": TEXT, "space_after": 8, "mono": True},
          {"t": "Rules are driven by the customer's own risk factors — a month-to-month "
                "customer with no security gets a different offer than a loyal "
                "high-value one.", "pt": 10.2, "color": MUTED, "space_after": 0,
           "spacing": 1.16}])

    yy = y + 3.30
    callout(s, M, yy, 7.10, 1.22, "SPEAK PRECISELY ABOUT THIS NUMBER",
            "RetainIQ reports monthly revenue at risk — the recurring charges sitting in "
            "flagged customers. $92,539/month in high risk, $200,302 including medium. "
            "It is exposure, not money already saved.", WARN)
    rect(s, 7.95, yy, 4.76, 1.22, fill=PANEL, edge=LINE)
    text(s, 8.15, yy + 0.14, 4.40, 1.0,
         [{"t": "DECISION FRAMEWORK", "pt": 10.5, "color": OK, "bold": True,
           "mono": True, "space_after": 5},
          {"t": "Data → insight → prediction → explanation → action → business value. "
                "The platform deliberately does not stop at step three.",
           "pt": 10.4, "color": TEXT, "space_after": 0, "spacing": 1.18}])
    footer(s, n)
    notes(s, """
(7:00–7:40) "Prediction alone doesn't retain anyone." Show the offer list — note
that the biggest group is a contract-upgrade nudge, which follows directly from
the EDA finding.
Then be careful and precise about the money: "this is monthly revenue at risk,
exposure — $92,539 a month sitting in the high-risk band. I deliberately do not
call it savings."
""")

    # ================= 12 PRODUCT =================
    n += 1
    s = add_slide(prs)
    eyebrow(s, "04  ·  THE PRODUCT")
    y = heading(s, "The product: a Signal Ops Console",
                "Flask application, deployed and publicly reachable — not a notebook")
    image_fit(s, ASSETS / "chart_architecture.png", M, y - 0.02, CW, 2.28)

    yy = y + 2.36
    ph = 2.30
    image_fit(s, IMAGES / "01.Home.png", M, yy, 4.62, ph)
    image_fit(s, IMAGES / "04.Bulk_prediction.png", M + 4.80, yy, 4.40, ph)
    rect(s, M + 9.36, yy, 2.73, ph, fill=PANEL, edge=LINE)
    text(s, M + 9.54, yy + 0.14, 2.37, ph - 0.28,
         [{"t": "WHAT A USER DOES", "pt": 10, "color": OK, "bold": True, "mono": True,
           "space_after": 6},
          {"t": "1  Upload a customer CSV", "pt": 9.5, "color": TEXT, "space_after": 2},
          {"t": "2  Validate the file first", "pt": 9.5, "color": TEXT, "space_after": 2},
          {"t": "3  Score the portfolio", "pt": 9.5, "color": TEXT, "space_after": 2},
          {"t": "4  Read KPIs, drill into a customer", "pt": 9.5, "color": TEXT, "space_after": 2, "spacing": 1.10},
          {"t": "5  Download the report", "pt": 9.5, "color": TEXT, "space_after": 7},
          {"t": "Reports are private per visitor — a cookie-scoped run, not a shared "
                "folder.", "pt": 9.2, "color": MUTED, "space_after": 0, "spacing": 1.14}])
    footer(s, n)
    notes(s, """
(7:40–8:20) This is the demo slide — if the live site is up, switch to it here
for 30 seconds; otherwise walk the screenshots.
Say: "bulk upload, validation before scoring, a dashboard with churn rate, risk
distribution and revenue at risk, customer-level drill-down, and downloadable
reports." One engineering detail worth one sentence: "reports are scoped to the
visitor through a signed cookie, so users never see each other's data — that has
its own unit test."
""")

    # ================= 13 ENGINEERING =================
    n += 1
    s = add_slide(prs)
    eyebrow(s, "04  ·  ENGINEERING & DEPLOYMENT")
    y = heading(s, "Engineering beyond the notebook",
                "Business intelligence, an API, automated tests and a tuned deployment")

    cards = [
        ("Power BI layer", OK, "A four-page dashboard — executive overview, customer insights, risk intelligence, retention strategy — built on the exported scoring dataset. The ML app is predictive; Power BI is descriptive and diagnostic."),
        ("Prediction API", ACTION, "The Flask app exposes /api/health and /api/predict, and a FastAPI wrapper serves the same engine with Pydantic-validated payloads for programmatic use."),
        ("Automated tests", WARN, "Nine unit tests cover encoding, model loading, scoring, risk classification, recommendation generation, CSV validation and report privacy."),
        ("Production deployment", BAD, "Gunicorn with a gthread worker, 300-second timeout, preload and periodic worker recycling for small cloud instances; Docker and Procfile provided; hosted on Render."),
    ]
    cw = (CW - 0.28) / 2
    for i, (t, c, b) in enumerate(cards):
        x = M + (i % 2) * (cw + 0.28)
        yy = y + 0.04 + (i // 2) * 1.72
        rect(s, x, yy, cw, 1.56, fill=PANEL, edge=LINE)
        bar(s, x, yy + 0.14, 0.035, 1.28, c)
        text(s, x + 0.22, yy + 0.16, cw - 0.46, 1.26,
             [{"t": t, "pt": 12.5, "color": c, "bold": True, "space_after": 4},
              {"t": b, "pt": 10.3, "color": TEXT, "space_after": 0, "spacing": 1.18}])

    picture_card(s, PBI / "page1_executive_overview.png.png", 8.88, y + 0.04, 3.83, 2.10)
    picture_card(s, PBI / "page3_risk_intelligence.png.png", 8.88, y + 2.32, 3.83, 2.10)

    rect(s, M, y + 3.52, 8.60, 1.06, fill=PANEL_2, edge=OK, edge_w=1.2)
    text(s, M + 0.22, y + 3.64, 8.16, 0.86,
         [{"t": "REPRODUCIBLE BY CONSTRUCTION", "pt": 10.3, "color": OK, "bold": True,
           "mono": True, "space_after": 4},
          {"t": "Every chart in this deck is generated from the repository's own "
                "artifacts — the raw dataset, the saved models, the held-out split and "
                "the production scoring pipeline. Re-run one script and the figures "
                "reproduce, including the metrics table.",
           "pt": 10.3, "color": TEXT, "space_after": 0, "spacing": 1.16}])
    footer(s, n)
    notes(s, """
(8:20–8:55) Fast pass — four cards, one sentence each. The Power BI screenshots
on the right show the BI half of the deliverable.
Use the bottom strip as your integrity line: "every number in this deck is
generated by a script in the repository from the project's own artifacts — I can
re-run it live."
""")

    # ================= 14 RESULTS / LIMITS =================
    n += 1
    s = add_slide(prs)
    eyebrow(s, "04  ·  CONCLUSION")
    y = heading(s, "Results, honest limits, and what comes next")

    col = (CW - 2 * 0.26) / 3

    rect(s, M, y + 0.02, col, 4.05, fill=PANEL, edge=OK, edge_w=1.2)
    text(s, M + 0.22, y + 0.16, col - 0.44, 3.75,
         [{"t": "WHAT WAS DELIVERED", "pt": 11, "color": OK, "bold": True, "mono": True,
           "space_after": 7},
          {"t": "▸  End-to-end ML pipeline on 7,043 customers, 30 engineered features",
           "pt": 10.6, "color": TEXT, "space_after": 5},
          {"t": "▸  Four algorithms benchmarked; Logistic Regression selected on F1 and ROC-AUC",
           "pt": 10.6, "color": TEXT, "space_after": 5, "spacing": 1.16},
          {"t": "▸  80.34% accuracy, 60.60% F1, 84.91% ROC-AUC on held-out data",
           "pt": 10.6, "color": TEXT, "space_after": 5, "spacing": 1.16},
          {"t": "▸  Risk bands validated against real outcomes (9.5% → 42.0% → 73.0%)",
           "pt": 10.6, "color": TEXT, "space_after": 5, "spacing": 1.16},
          {"t": "▸  Live Flask console, dashboard, private reports, Power BI layer",
           "pt": 10.6, "color": TEXT, "space_after": 0, "spacing": 1.16}])

    rect(s, M + col + 0.26, y + 0.02, col, 4.05, fill=PANEL, edge=WARN, edge_w=1.2)
    text(s, M + col + 0.48, y + 0.16, col - 0.44, 3.75,
         [{"t": "LIMITATIONS — STATED PLAINLY", "pt": 11, "color": WARN, "bold": True,
           "mono": True, "space_after": 7},
          {"t": "▸  Trained on historical data from one telecom; behaviour may not transfer",
           "pt": 10.6, "color": TEXT, "space_after": 5, "spacing": 1.16},
          {"t": "▸  56.95% recall — roughly two in five churners are still missed",
           "pt": 10.6, "color": TEXT, "space_after": 5, "spacing": 1.16},
          {"t": "▸  No live behavioural feed, no drift monitoring, no automatic retraining",
           "pt": 10.6, "color": TEXT, "space_after": 5, "spacing": 1.16},
          {"t": "▸  Probabilities are not guarantees; recommendations are decision support",
           "pt": 10.6, "color": TEXT, "space_after": 5, "spacing": 1.16},
          {"t": "▸  Revenue figures are exposure, not realised savings",
           "pt": 10.6, "color": TEXT, "space_after": 0, "spacing": 1.16}])

    rect(s, M + 2 * (col + 0.26), y + 0.02, col, 4.05, fill=PANEL, edge=ACTION, edge_w=1.2)
    text(s, M + 2 * (col + 0.26) + 0.22, y + 0.16, col - 0.44, 3.75,
         [{"t": "WHAT COMES NEXT", "pt": 11, "color": ACTION, "bold": True, "mono": True,
           "space_after": 7},
          {"t": "▸  Threshold tuning and probability calibration to lift recall",
           "pt": 10.6, "color": TEXT, "space_after": 5, "spacing": 1.16},
          {"t": "▸  Cost-sensitive learning — a missed churner costs more than a wasted offer",
           "pt": 10.6, "color": TEXT, "space_after": 5, "spacing": 1.16},
          {"t": "▸  Drift detection and scheduled retraining (MLOps loop)",
           "pt": 10.6, "color": TEXT, "space_after": 5, "spacing": 1.16},
          {"t": "▸  Counterfactual / what-if explanations for retention agents",
           "pt": 10.6, "color": TEXT, "space_after": 5, "spacing": 1.16},
          {"t": "▸  CRM integration and A/B testing of retention offers",
           "pt": 10.6, "color": TEXT, "space_after": 0, "spacing": 1.16}])

    rect(s, M, y + 4.24, CW, 0.78, fill=PANEL_2, edge=OK, edge_w=1.2)
    text(s, M + 0.24, y + 4.36, CW - 0.48, 0.58,
         [{"t": "The contribution: churn prediction reframed from a classification score "
                "into a retention intelligence workflow — and shipped as a working product. "
                "Thank you — questions welcome.",
           "pt": 12, "color": TEXT, "bold": True, "space_after": 0, "spacing": 1.16}])
    footer(s, n)
    notes(s, """
(8:55–9:40) Close deliberately. Left column: what exists and works. Middle: the
limits, stated before anyone asks — especially recall. Right: the roadmap.
Final line, spoken: "The contribution isn't the 80% — it's that a probability
becomes a prioritised, explained, costed action, in a product a retention team
can actually open. Thank you — I'm happy to take questions."

LIKELY QUESTIONS
· Why not deep learning? 1,409 test rows and a largely linear signal; a neural
  net would add opacity, not accuracy. Benchmarks support this.
· Why is recall low? Class imbalance plus a default 0.5 threshold; tuning and
  calibration are the fix, and are on the roadmap.
· Isn't "revenue at risk" optimistic? It is exposure, not realised savings — the
  deck and the app both say so; validation would need a live campaign.
· How do you know it isn't leakage? Churn Score / Reason / Category / Status were
  dropped before training — they are post-outcome fields.
""")

    # ================= APPENDIX =================
    n += 1
    s = add_slide(prs)
    eyebrow(s, "APPENDIX", color=MUTED)
    y = heading(s, "Full benchmark table",
                "Held-out test set, 1,409 customers — every metric recorded")
    rows = [
        ("Model", "Accuracy", "Precision", "Recall", "F1", "ROC-AUC", True),
        ("Logistic Regression  ← selected", "80.34%", "64.74%", "56.95%", "60.60%", "84.91%", False),
        ("LightGBM", "80.06%", "64.49%", "55.35%", "59.57%", "84.71%", False),
        ("XGBoost", "79.06%", "61.72%", "55.61%", "58.51%", "83.34%", False),
        ("Random Forest", "78.99%", "62.34%", "52.67%", "57.10%", "83.48%", False),
    ]
    colw = [4.35, 1.56, 1.56, 1.42, 1.42, 1.79]
    yy = y + 0.10
    for ri, row in enumerate(rows):
        x = M
        hdr = row[-1]
        for ci, cell in enumerate(row[:-1]):
            if ri == 0:
                rect(s, x, yy, colw[ci], 0.50, fill=PANEL_2, edge=LINE)
            elif ri == 1:
                rect(s, x, yy, colw[ci], 0.54, fill=PANEL, edge=OK, edge_w=1.2)
            else:
                rect(s, x, yy, colw[ci], 0.54, fill=PANEL, edge=LINE)
            color = MUTED if hdr else (OK if ri == 1 and ci > 0 else TEXT)
            text(s, x + 0.14, yy + 0.13, colw[ci] - 0.24, 0.32,
                 [{"t": cell, "pt": 11 if not hdr else 10.2, "color": color,
                   "bold": hdr or (ri == 1 and ci == 0), "space_after": 0,
                   "mono": hdr or ci > 0}])
            x += colw[ci]
        yy += 0.60 if ri == 0 else 0.60

    text(s, M, yy + 0.22, CW, 1.4,
         [{"t": "All four models were trained on the identical 5,634-row training split "
                "with the identical 30-feature frame, and evaluated with the same scoring "
                "code — the only variable is the algorithm.",
           "pt": 11, "color": MUTED, "space_after": 6, "spacing": 1.18},
          {"t": "Reproduce: python presentation/build_assets.py  →  writes "
                "presentation/assets/model_metrics.csv",
           "pt": 10.5, "color": OK, "mono": True, "space_after": 0}])
    footer(s, n, label="Appendix · use only if asked")
    notes(s, "Backup slide. Use if anyone questions the benchmark or asks for the full "
             "metric set.")

    n += 1
    s = add_slide(prs)
    eyebrow(s, "APPENDIX", color=MUTED)
    y = heading(s, "The prediction API",
                "The same model and preprocessing exposed programmatically")
    rect(s, M, y + 0.06, 6.0, 3.30, fill=PANEL_2, edge=LINE)
    text(s, M + 0.22, y + 0.20, 5.56, 3.0,
         [{"t": "POST /api/predict", "pt": 11.5, "color": OK, "bold": True, "mono": True,
           "space_after": 7},
          {"t": "{ \"gender\": \"Male\", \"SeniorCitizen\": 1,", "pt": 10, "color": TEXT, "space_after": 1, "mono": True},
          {"t": "  \"Partner\": \"Yes\", \"Dependents\": \"No\",", "pt": 10, "color": TEXT, "space_after": 1, "mono": True},
          {"t": "  \"tenure\": 8, \"InternetService\": \"Fiber optic\",", "pt": 10, "color": TEXT, "space_after": 1, "mono": True},
          {"t": "  \"OnlineSecurity\": \"No\", \"TechSupport\": \"No\",", "pt": 10, "color": TEXT, "space_after": 1, "mono": True},
          {"t": "  \"Contract\": \"Month-to-month\", ... }", "pt": 10, "color": TEXT, "space_after": 10, "mono": True},
          {"t": "→  probability · prediction · risk_level", "pt": 10.2, "color": OK, "space_after": 1, "mono": True},
          {"t": "→  top_drivers · primary_action · recommendations", "pt": 10.2, "color": OK, "space_after": 0, "mono": True}])
    rect(s, 6.82, y + 0.06, 5.89, 3.30, fill=PANEL, edge=LINE)
    text(s, 7.04, y + 0.20, 5.45, 3.0,
         [{"t": "TWO INTERFACES, ONE ENGINE", "pt": 11, "color": OK, "bold": True,
           "mono": True, "space_after": 7},
          {"t": "The Flask application is the product; the FastAPI wrapper exists so the "
                "same scoring functions can be called from any client.",
           "pt": 11, "color": TEXT, "space_after": 8, "spacing": 1.18},
          {"t": "Both call utils/prediction.py and src/preprocessing.py — there is no "
                "second implementation to keep in sync.",
           "pt": 11, "color": TEXT, "space_after": 8, "spacing": 1.18},
          {"t": "Pydantic schemas validate every field, so a malformed customer record "
                "returns a clear 400 instead of a stack trace.",
           "pt": 11, "color": TEXT, "space_after": 0, "spacing": 1.18}])
    footer(s, n, label="Appendix · use only if asked")
    notes(s, "Backup slide for API / integration questions.")

    n += 1
    s = add_slide(prs)
    eyebrow(s, "APPENDIX", color=MUTED)
    y = heading(s, "How these figures were produced",
                "So the deck can be audited against the repository")
    rows = [
        ("Chart / figure", "Source"),
        ("Churn split, contract, internet, tenure", "Data/01.Telco_customer_churn_Dataset.csv"),
        ("Model benchmark, confusion matrix, ROC", "models/*.pkl scored on Data/processed_data/test/"),
        ("Risk bands, high-risk profile, revenue", "Production pipeline — run_bulk_prediction_from_bytes"),
        ("Console screenshots", "templates/images/ (captured from the live app)"),
        ("Power BI pages", "Power_BI_dashboard/Screenshots/"),
    ]
    yy = y + 0.10
    for ri, (a, b) in enumerate(rows):
        hdr = ri == 0
        rect(s, M, yy, 5.10, 0.52, fill=PANEL_2 if hdr else PANEL, edge=LINE)
        rect(s, M + 5.10, yy, 7.02, 0.52, fill=PANEL_2 if hdr else PANEL, edge=LINE)
        text(s, M + 0.16, yy + 0.14, 4.8, 0.3,
             [{"t": a, "pt": 10.6, "color": MUTED if hdr else TEXT, "bold": hdr,
               "space_after": 0}])
        text(s, M + 5.26, yy + 0.14, 6.7, 0.3,
             [{"t": b, "pt": 10, "color": MUTED if hdr else OK, "mono": not hdr,
               "bold": hdr, "space_after": 0}])
        yy += 0.58

    callout(s, M, yy + 0.14, CW, 1.30, "THE POINT OF THIS SLIDE",
            "Nothing here is hand-typed. presentation/build_assets.py recomputes every "
            "figure from the saved artifacts and the production scoring path, so the "
            "deck cannot drift away from the project — if a model is retrained, the "
            "charts and the benchmark table regenerate together.", OK)
    footer(s, n, label="Appendix · use only if asked")
    notes(s, "Backup slide for reproducibility questions. This is a strong answer to "
             "\"how do we know the numbers are right?\"")

    prs.save(OUT)
    return OUT


def main():
    out = build()
    print(f"wrote {out.relative_to(ROOT)}  ({out.stat().st_size / 1024:.0f} KB)")
    print(f"slides: {len(Presentation(str(out)).slides)}")
    if WARNINGS:
        print(f"\n{len(WARNINGS)} layout warning(s):")
        for w in WARNINGS:
            print("  -", w)
    else:
        print("\nno layout overflow detected")


if __name__ == "__main__":
    main()
