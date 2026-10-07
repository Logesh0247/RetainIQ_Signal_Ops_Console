"""
Build the RetainIQ storytelling deck (.pptx).

Narrative: the deck follows one telecom's weakening customer signal from raw
data to a prioritised, explained, costed retention action, laid out as the
project's own 14-phase development cycle across three acts.

Theme: mirrors the application's "Signal Ops" dark console — ink background,
signal-green accent, monospace data — so slides and product look like one thing.

Facts: every number comes from the repository (raw dataset, saved models,
held-out split, production scoring pipeline). See presentation/build_assets.py.

Run from the repository root:

    python presentation/build_assets.py     # charts + metric tables
    python presentation/build_qr.py         # live-app + repository QR codes
    python presentation/build_deck.py       # this file
"""
from __future__ import annotations

from pathlib import Path

from PIL import Image, ImageFont
from pptx import Presentation
from pptx.dml.color import RGBColor
from pptx.enum.shapes import MSO_SHAPE
from pptx.enum.text import MSO_ANCHOR, PP_ALIGN
from pptx.util import Inches, Pt

HERE = Path(__file__).resolve().parent
ROOT = HERE.parent
ASSETS = HERE / "assets"
IMAGES = ROOT / "templates" / "images"
PBI = ROOT / "Power_BI_dashboard" / "Screenshots"
OUT = HERE / "RetainIQ_Project_Review.pptx"

URL_FILE = HERE / "live_url.txt"
LIVE_URL = (URL_FILE.read_text().strip() if URL_FILE.exists()
            else "https://retainiq-predictive-customer-retention-zq6x.onrender.com")
LIVE_URL_DISPLAY = LIVE_URL.replace("https://", "")
REPO_DISPLAY = "github.com/Logesh0247/RetainIQ_Signal_Ops_Console"

# --- palette (static/style.css tokens) --------------------------------------
INK = RGBColor(0x0A, 0x0F, 0x1C)
PANEL = RGBColor(0x11, 0x1A, 0x2C)
PANEL_2 = RGBColor(0x16, 0x21, 0x36)
LINE = RGBColor(0x22, 0x30, 0x49)
LINE_SOFT = RGBColor(0x1A, 0x25, 0x40)
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

RAIL_Y = 6.90
FOOT_RULE_Y = 7.06
FOOT_TEXT_Y = 7.12

PHASES = [
    "Business Understanding", "Data Collection", "Data Cleaning & Preprocessing",
    "Exploratory Data Analysis", "Feature Engineering", "Model Development",
    "Model Evaluation", "Model Selection", "Explainability", "Risk Segmentation",
    "Retention Intelligence", "Business Intelligence", "Web Application", "Deployment",
]

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
    shape = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(x), Inches(y),
                                   Inches(w), Inches(h))
    shape.fill.solid()
    shape.fill.fore_color.rgb = color
    shape.line.fill.background()
    shape.shadow.inherit = False
    return shape


def text(slide, x, y, w, h, blocks, align=PP_ALIGN.LEFT, anchor=MSO_ANCHOR.TOP,
         line_spacing=1.12, audit=True):
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

        if audit:
            pt = item["pt"]
            needed = wrap_lines(item["t"], pt, w, item.get("bold", False),
                                item.get("mono", False))
            required_h = needed * pt * 1.22 * item.get("spacing", line_spacing) / 72
            if required_h > h + 0.06:
                WARNINGS.append(
                    f"possible vertical overflow: {needed} lines need "
                    f"{required_h:.2f}in in {h:.2f}in box -> \"{item['t'][:56]}...\"")
        elif False:
            pass
        width_needed = text_width_pt(item["t"], item["pt"], item.get("bold", False),
                                     item.get("mono", False))
        if width_needed > w * 72 + 1 and " " not in item["t"]:
            WARNINGS.append(
                f"unbreakable text wider than box ({width_needed / 72:.2f}in > "
                f"{w:.2f}in) -> \"{item['t'][:50]}\"")
    return box


def eyebrow(slide, label, x=M, y=0.40, color=OK):
    text(slide, x, y, 9.0, 0.26,
         [{"t": label, "pt": 10, "color": color, "bold": True, "mono": True,
           "space_after": 0}])


def heading(slide, title, sub=None, title_pt=26):
    text(slide, M, 0.63, CW, 0.52,
         [{"t": title, "pt": title_pt, "color": TEXT, "bold": True, "space_after": 0}])
    y = 1.16
    if sub:
        text(slide, M, y, CW, 0.30,
             [{"t": sub, "pt": 12.5, "color": MUTED, "space_after": 0}])
        y += 0.34
    bar(slide, M, y, CW, 0.014, LINE)
    return y + 0.16


def rail(slide, current=None, complete_through=None):
    """The 14-phase progress rail: the development cycle made visible."""
    n = len(PHASES)
    gap = 0.055
    tw = (CW - (n - 1) * gap) / n
    for i in range(n):
        x = M + i * (tw + gap)
        if complete_through is not None and i < complete_through:
            color, hgt, y = OK, 0.048, RAIL_Y
        elif current is not None and i == current:
            color, hgt, y = OK, 0.075, RAIL_Y - 0.014
        else:
            color, hgt, y = LINE, 0.048, RAIL_Y
        bar(slide, x, y, tw, hgt, color)


def footer(slide, number, phase=None, complete_through=None, label="RetainIQ · Signal Ops Console"):
    rail(slide, current=phase, complete_through=complete_through)
    bar(slide, M, FOOT_RULE_Y, CW, 0.010, LINE)
    left = label
    if phase is not None:
        left = f"{label}   ·   PHASE {phase + 1:02d} · {PHASES[phase].upper()}"
    text(slide, M, FOOT_TEXT_Y, 10.4, 0.24,
         [{"t": left, "pt": 9, "color": MUTED, "space_after": 0, "mono": False}], audit=False)
    text(slide, SW - M - 1.2, FOOT_TEXT_Y, 1.2, 0.24,
         [{"t": f"{number:02d}", "pt": 9, "color": MUTED, "mono": True,
           "space_after": 0, "align": PP_ALIGN.RIGHT}], audit=False)


def kpi(slide, x, y, w, h, value, label, color=OK, value_pt=25, label_pt=9.5):
    rect(slide, x, y, w, h, fill=PANEL, edge=LINE)
    text(slide, x + 0.16, y + 0.14, w - 0.28, h - 0.22,
         [{"t": value, "pt": value_pt, "color": color, "bold": True, "space_after": 2},
          {"t": label, "pt": label_pt, "color": MUTED, "space_after": 0}])


def callout(slide, x, y, w, h, title, body, color=WARN, body_pt=10.8):
    rect(slide, x, y, w, h, fill=PANEL_2, edge=color, edge_w=1.2)
    text(slide, x + 0.20, y + 0.14, w - 0.40, h - 0.26,
         [{"t": title, "pt": 11.5, "color": color, "bold": True, "space_after": 3},
          {"t": body, "pt": body_pt, "color": TEXT, "space_after": 0, "spacing": 1.16}])


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
    rect(slide, x, y, w, h, fill=fill, edge=edge)
    image_fit(slide, path, x + pad, y + pad, w - 2 * pad, h - 2 * pad)


def signal_meter(slide, x, y, filled=3, total=5, scale=1.0, label=None, color=OK):
    """The deck's recurring motif: a customer's signal strength, 1-5 bars."""
    bw = 0.115 * scale
    gap = 0.085 * scale
    heights = [0.22, 0.38, 0.54, 0.70, 0.86]
    base = y + max(heights) * scale
    for i in range(total):
        hgt = heights[i] * scale
        c = color if i < filled else LINE
        bar(slide, x + i * (bw + gap), base - hgt, bw, hgt, c)
    if label:
        text(slide, x, base + 0.10 * scale, 2.2, 0.24,
             [{"t": label, "pt": 8.5 * scale, "color": MUTED, "mono": True,
               "space_after": 0}], audit=False)


# --- deck -------------------------------------------------------------------
def build():
    prs = Presentation()
    prs.slide_width = Inches(SW)
    prs.slide_height = Inches(SH)
    n = 0

    # =========================================================== 01 TITLE
    n += 1
    s = add_slide(prs)
    bar(s, 0, 0, SW, 0.085, OK)
    text(s, M, 1.10, 9.6, 0.28,
         [{"t": "FINAL YEAR PROJECT   ·   B.SC. DATA SCIENCE", "pt": 11,
           "color": OK, "bold": True, "mono": True, "space_after": 0}])
    text(s, M, 1.46, 10.4, 1.62,
         [{"t": "RetainIQ", "pt": 54, "color": TEXT, "bold": True, "space_after": 2},
          {"t": "Signal Ops Console", "pt": 28, "color": OK, "bold": True,
           "space_after": 0}])
    bar(s, M, 3.16, 2.1, 0.035, OK)
    text(s, M, 3.42, 8.1, 1.0,
         [{"t": "An explainable machine learning platform for customer churn "
                "prediction, risk intelligence and retention strategy.",
           "pt": 14.5, "color": MUTED, "space_after": 0, "spacing": 1.25}])

    signal_meter(s, 11.05, 1.62, filled=5, scale=1.0,
                 label="signal strength", color=OK)

    kpi(s, M, 4.72, 2.36, 1.02, "80.34%", "TEST ACCURACY", OK)
    kpi(s, M + 2.50, 4.72, 2.36, 1.02, "84.91%", "ROC-AUC", OK)
    kpi(s, M + 5.00, 4.72, 2.36, 1.02, "7,043", "CUSTOMERS SCORED", ACTION)
    kpi(s, M + 7.50, 4.72, 4.59, 1.02, "$92,539",
        "MONTHLY REVENUE AT RISK IN THE HIGH-RISK BAND", BAD, 25, 9)

    text(s, M, 6.00, 12.1, 0.72,
         [{"t": "Logesh S.   ·   B.Sc. Data Science", "pt": 12.5, "color": TEXT,
           "bold": True, "space_after": 3},
          {"t": f"Live application: {LIVE_URL_DISPLAY}     ·     Source: {REPO_DISPLAY}",
           "pt": 10.5, "color": MUTED, "mono": True, "space_after": 0}])
    notes(s, """
OPENING (0:00–0:35). "Good morning. A telecom knows it lost a quarter of its
customers last year. What it cannot tell you is who leaves next month, why, or
what to do about it. That gap is the whole project.

RetainIQ reads a customer's weakening signal — from raw records, to a predicted
probability, to an explained reason, to a recommended action, and finally to the
revenue that action protects. 80.34% accuracy, 84.91% ROC-AUC, 7,043 customers
scored, and $92,539 a month of recurring revenue sitting in the high-risk band."

Then move — do not read the tiles.
""")

    # =========================================================== 02 THE CYCLE
    n += 1
    s = add_slide(prs)
    eyebrow(s, "THE DEVELOPMENT CYCLE")
    y = heading(s, "Fourteen phases, three acts",
                "This is the map for the whole talk — every slide that follows is one of these boxes")
    image_fit(s, ASSETS / "chart_cycle.png", M, y - 0.04, CW, 3.30)
    callout(s, M, y + 3.38, CW, 1.14, "HOW TO READ THE DECK",
            "Act I shows the signal fading — the business problem and the data that "
            "records it. Act II is reading the signal — building the model and "
            "interrogating how far to trust it. Act III is the signal recovered — what "
            "the intelligence is worth and the product that delivers it.", OK)
    footer(s, n)
    notes(s, """
(0:35–1:05) "Everything I show you maps to one of these fourteen boxes, and the
boxes group into three acts."
Point along the arcs: "Act one, the signal fades — the problem and the data. Act
two, reading the signal — the model work. Act three — what it's worth and the
product."
Promise the structure and then keep it. Do not narrate every box.
""")

    # =========================================================== 03 PHASE 01
    n += 1
    s = add_slide(prs)
    eyebrow(s, "PHASE 01  ·  BUSINESS UNDERSTANDING")
    y = heading(s, "The problem: retention is still reactive",
                "Reports explain what happened. They do not say who leaves next.")
    bullets(s, M, y + 0.06, 6.55, 2.1, [
        ("Churn compounds.",
         "26.54% of this customer base churned — 1,869 of 7,043 customers."),
        ("Dashboards describe, they do not prioritise.",
         "Knowing last quarter's churn rate does not tell a retention team whom to call today."),
        ("Generic campaigns waste budget.",
         "Blanket offers discount loyal customers and still miss the ones about to leave."),
    ], pt=11.5, gap=8)

    qs = [("Who is likely to churn?", "Predicted churn probability per customer", ACTION),
          ("Why are they at risk?", "Feature contributions behind each prediction", WARN),
          ("What should we do?", "A retention action matched to the risk drivers", OK),
          ("What is at stake?", "Monthly recurring revenue exposed in each risk band", BAD)]
    for i, (q, sub, color) in enumerate(qs):
        x = M + (i % 2) * 3.34
        yy = y + 2.42 + (i // 2) * 1.30
        rect(s, x, yy, 3.16, 1.14, fill=PANEL, edge=LINE)
        bar(s, x, yy + 0.14, 0.035, 0.86, color)
        text(s, x + 0.20, yy + 0.16, 2.86, 0.9,
             [{"t": q, "pt": 12, "color": TEXT, "bold": True, "space_after": 3},
              {"t": sub, "pt": 9.8, "color": MUTED, "space_after": 0, "spacing": 1.14}])

    rect(s, 7.62, y + 0.06, 5.09, 4.90, fill=PANEL, edge=LINE)
    text(s, 7.86, y + 0.22, 4.6, 0.3,
         [{"t": "THE BASE RATE", "pt": 10, "color": MUTED, "bold": True, "mono": True,
           "space_after": 0}])
    image_fit(s, ASSETS / "chart_churn_split.png", 7.72, y + 0.50, 4.89, 3.90)
    text(s, 7.86, y + 4.40, 4.6, 0.5,
         [{"t": "Predicting \"nobody churns\" is already 73.46% accurate — which is why "
                "accuracy alone decides nothing in this project.",
           "pt": 10.5, "color": MUTED, "space_after": 0, "spacing": 1.16}])
    footer(s, n, phase=0)
    notes(s, """
(1:05–1:55) Frame the pain, then the four questions — they structure the entire
solution. Finish on the donut: "Notice the trap. If I predict nobody churns I'm
73.46% accurate and completely useless. That single number is why no model in
this deck is chosen on accuracy."
That line earns credibility early and sets up Phase 07.
""")

    # =========================================================== 04 PHASE 02
    n += 1
    s = add_slide(prs)
    eyebrow(s, "PHASE 02  ·  DATA COLLECTION")
    y = heading(s, "The data that records the signal",
                "IBM Telco Customer Churn — 7,043 customers, 34 raw fields, one binary target")
    tiles = [("7,043", "CUSTOMERS", TEXT), ("34", "RAW FEATURES", TEXT),
             ("30", "ML FEATURES", TEXT), ("26.54%", "CHURN RATE", BAD),
             ("80 / 20", "STRATIFIED SPLIT", ACTION)]
    tw = (CW - 4 * 0.16) / 5
    for i, (v, l, c) in enumerate(tiles):
        kpi(s, M + i * (tw + 0.16), y + 0.04, tw, 1.06, v, l, c, 20, 9)

    yy = y + 1.30
    text(s, M, yy, 5.9, 2.6,
         [{"t": "What a row contains", "pt": 13, "color": OK, "bold": True,
           "space_after": 6},
          {"t": "  ▸  Customer profile — gender, senior citizen, partner, dependents",
           "pt": 11.3, "color": TEXT, "space_after": 4},
          {"t": "  ▸  Services — phone, internet, security, backup, protection, support, streaming",
           "pt": 11.3, "color": TEXT, "space_after": 4, "spacing": 1.16},
          {"t": "  ▸  Contract & billing — contract type, payment method, paperless, charges",
           "pt": 11.3, "color": TEXT, "space_after": 4},
          {"t": "  ▸  Value — tenure, monthly charges, total charges, CLTV",
           "pt": 11.3, "color": TEXT, "space_after": 4},
          {"t": "Target: Churn Label (Yes / No), converted to 1 / 0.",
           "pt": 11.3, "color": MUTED, "space_after": 0, "space_before": 8}])

    callout(s, 6.85, yy, 5.86, 1.30, "CLASS BALANCE AND THE SPLIT",
            "1,869 churned against 5,174 retained. The split is stratified, so both "
            "halves carry the identical 26.54% churn rate: 5,634 rows train, 1,409 "
            "rows are held back and touched only once.", ACTION)
    callout(s, 6.85, yy + 1.48, 5.86, 1.30, "TRAIN / SERVE PARITY",
            "The 30 encoded columns are frozen in feature_columns.pkl and rebuilt by "
            "the same preprocessing code for every upload — training and live scoring "
            "cannot drift apart. There is a unit test for exactly this.", OK)
    footer(s, n, phase=1)
    notes(s, """
(1:55–2:35) Quick pass on the tiles. Spend the time on the right-hand callouts:
"the split is stratified, so both halves carry exactly the same 26.54% churn rate
— the model never sees a distorted class balance."
And the parity point: "the 30-column contract is saved next to the model and
tested, so the app cannot silently score on the wrong schema."
""")

    # =========================================================== 05 PHASE 03
    n += 1
    s = add_slide(prs)
    eyebrow(s, "PHASE 03  ·  DATA CLEANING & PREPROCESSING")
    y = heading(s, "From raw records to 30 model-ready features",
                "One reproducible pipeline, used identically in training and in the live app")
    steps = [("Raw upload", "7,043 records · 34 columns", ACTION),
             ("Clean", "types, blanks, duplicates, category names", ACTION),
             ("Encode", "16 categoricals → 27 flags", WARN),
             ("30 features", "frozen column contract", OK)]
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
                   "align": PP_ALIGN.CENTER}], audit=False)

    yy = y + 1.44
    callout(s, M, yy, CW, 1.34,
            "LEAKAGE PREVENTION — the detail that makes the results believable",
            "Churn Score, Churn Reason, Churn Category and Customer Status were removed "
            "before training. All four are only known after a customer has already "
            "churned, so keeping them would inflate the metrics and break the model in "
            "production. This is the difference between a demo and a deployable model.")
    bullets(s, M, yy + 1.54, 12.09, 1.5, [
        ("Cleaning", "Data-type correction, 11 blank TotalCharges values (all zero-tenure new "
                     "customers), duplicate checks, categorical normalisation, removal of ID "
                     "and geographic columns that carry no predictive value."),
        ("Encoding", "Every service column becomes an explicit flag — Contract_Two year, "
                     "Tech Support_Yes, Internet Service_Fiber optic — so no ordering is "
                     "invented where none exists: 3 numeric features + 27 flags = 30."),
    ], pt=11.0, gap=9)
    footer(s, n, phase=2)
    notes(s, """
(2:35–3:20) Walk the four boxes quickly, then slow down on the leakage callout —
reviewers consistently reward this. "Churn Score and Churn Reason are the
dataset's most seductive columns and they're poison: they only exist after the
customer has left. I dropped them. That is why 80% here is a real 80%."
One sentence on encoding and move on.
""")

    # =========================================================== 06 PHASE 04
    n += 1
    s = add_slide(prs)
    eyebrow(s, "PHASE 04  ·  EXPLORATORY DATA ANALYSIS")
    y = heading(s, "What the data said",
                "Three patterns that shaped the model — and the retention strategy")
    image_fit(s, ASSETS / "chart_contract.png", M, y + 0.04, 5.95, 3.24)
    image_fit(s, ASSETS / "chart_tenure.png", M + 6.15, y + 0.04, 5.94, 3.24)
    yy = y + 3.40
    cards = [
        ("Contract is the strongest signal",
         "42.7% churn on month-to-month vs 2.8% on two-year contracts — a 15× gap.", OK),
        ("Churn is front-loaded",
         "47.4% of first-year customers churn; 55.5% of all churners are in year one. "
         "Churned customers average 18.0 months vs 37.6.", WARN),
        ("It is a service-experience story",
         "Fiber optic 41.9% vs 7.4% with no internet; customers without tech support or "
         "online security churn at 41.6% and 41.8%.", ACTION),
    ]
    cw = (CW - 2 * 0.18) / 3
    for i, (t, b, c) in enumerate(cards):
        x = M + i * (cw + 0.18)
        rect(s, x, yy, cw, 1.66, fill=PANEL, edge=LINE)
        bar(s, x, yy + 0.16, 0.035, 1.34, c)
        text(s, x + 0.20, yy + 0.18, cw - 0.40, 1.34,
             [{"t": t, "pt": 11.5, "color": c, "bold": True, "space_after": 4},
              {"t": b, "pt": 10.2, "color": TEXT, "space_after": 0, "spacing": 1.18}])
    footer(s, n, phase=3)
    notes(s, """
(3:20–4:05) Three takeaways only, no bar-by-bar narration.
"Month-to-month is a 15× risk multiplier."
"Risk is front-loaded — year one is where retention money belongs."
"And the story is service experience, not just price: fiber optic, no security,
no support."
Bridge out: "these are the patterns the model later recovered on its own — and
what the recommendation engine acts on."
""")

    # =========================================================== 07 PHASE 05
    n += 1
    s = add_slide(prs)
    eyebrow(s, "PHASE 05  ·  FEATURE ENGINEERING")
    y = heading(s, "Thirty features, one frozen contract",
                "Three numeric measures, twenty-seven explicit flags — and a test that guards them")
    kpi(s, M, y + 0.04, 3.92, 1.10, "3", "NUMERIC FEATURES", ACTION, 26, 9.5)
    kpi(s, M + 4.10, y + 0.04, 3.92, 1.10, "27", "BINARY FLAGS", WARN, 26, 9.5)
    kpi(s, M + 8.20, y + 0.04, 3.89, 1.10, "30", "TOTAL MODEL INPUTS", OK, 26, 9.5)

    yy = y + 1.32
    rect(s, M, yy, 5.95, 2.72, fill=PANEL, edge=LINE)
    text(s, M + 0.22, yy + 0.16, 5.51, 2.4,
         [{"t": "THE THREE NUMERIC MEASURES", "pt": 10.5, "color": ACTION, "bold": True,
           "mono": True, "space_after": 6},
          {"t": "Tenure Months, Monthly Charges, Total Charges — the only continuous "
                "quantities the model sees.", "pt": 10.8, "color": TEXT, "space_after": 12,
           "spacing": 1.18},
          {"t": "WHY FLAGS, NOT NUMBERS", "pt": 10.5, "color": WARN, "bold": True,
           "mono": True, "space_after": 6},
          {"t": "Assigning 1, 2, 3 to DSL / Fiber / None would invent an ordering that "
                "does not exist. Flags keep the model honest about what these categories "
                "actually are.", "pt": 10.8, "color": TEXT, "space_after": 0,
           "spacing": 1.18}])

    rect(s, 6.75, yy, 5.96, 2.72, fill=PANEL_2, edge=OK, edge_w=1.2)
    text(s, 6.95, yy + 0.16, 5.56, 2.4,
         [{"t": "THE COLUMN CONTRACT", "pt": 10.5, "color": OK, "bold": True,
           "mono": True, "space_after": 6},
          {"t": "The exact 30 column names are saved to feature_columns.pkl alongside the "
                "model. The app rebuilds a feature frame and checks it against that list "
                "before scoring.", "pt": 10.8, "color": TEXT, "space_after": 12,
           "spacing": 1.18},
          {"t": "If the two ever disagree, the request fails loudly rather than silently "
                "scoring on a misaligned frame — the most common way a student project "
                "works in the notebook and quietly breaks in the app.", "pt": 10.8,
           "color": MUTED, "space_after": 0, "spacing": 1.18}])
    footer(s, n, phase=4)
    notes(s, """
(4:05–4:35) Short slide, but it answers a question examiners like: "why did you
encode it that way?"
"Three continuous measures. Everything else is a yes/no flag — because numbering
DSL as 1 and Fiber as 2 would invent an ordering that isn't real."
Then the contract point: "the 30 column names are saved with the model and
checked at scoring time, so a schema mismatch fails loudly instead of silently
producing garbage."
""")

    # =========================================================== 08 PHASE 06
    n += 1
    s = add_slide(prs)
    eyebrow(s, "PHASE 06  ·  MODEL DEVELOPMENT")
    y = heading(s, "Four algorithms, one pipeline",
                "Same features, same split, same scoring code — only the algorithm changes")
    image_fit(s, ASSETS / "chart_models.png", M, y + 0.04, 7.85, 3.40)
    bullets(s, 8.75, y + 0.10, 3.96, 3.30, [
        ("Logistic Regression", "linear baseline; interpretable coefficients"),
        ("Random Forest", "bagged trees; captures non-linearity"),
        ("XGBoost & LightGBM", "gradient boosting; strong on tabular data"),
    ], pt=10.8, gap=8)
    rect(s, 8.75, y + 2.42, 3.96, 1.02, fill=PANEL_2, edge=WARN, edge_w=1.2)
    text(s, 8.93, y + 2.54, 3.60, 0.84,
         [{"t": "Benchmark ladder (F1)", "pt": 10.2, "color": WARN, "bold": True,
           "space_after": 3},
          {"t": "LR 60.6 · LightGBM 59.6 · XGBoost 58.5 · RF 57.1",
           "pt": 10.2, "color": TEXT, "space_after": 0, "mono": True}])
    callout(s, M, y + 3.58, CW, 1.06, "WHY FOUR MODELS AND NOT ONE",
            "A single model proves nothing. Four algorithms on the same pipeline show the "
            "result is a property of the features and the problem — not of one lucky "
            "configuration. Accuracy, precision, recall, F1 and ROC-AUC were all "
            "recorded, and the winner was chosen on F1 and ROC-AUC, never accuracy.")
    footer(s, n, phase=5)
    notes(s, """
(4:35–5:05) "Same features, same split, same code path — the only thing that
changes is the algorithm."
Then the key observation: Logistic Regression leads on F1 and ROC-AUC, with
LightGBM a close second. Say why it matters: "churn here is largely linear in the
encoded features, so the simplest model wins — and that also buys me
explainability for free." That sets up Phase 09.
""")

    # =========================================================== 09 PHASE 07
    n += 1
    s = add_slide(prs)
    eyebrow(s, "PHASE 07  ·  MODEL EVALUATION")
    y = heading(s, "Measuring the model — and its stability",
                "Held-out test set of 1,409 customers, plus five-fold cross-validation")
    image_fit(s, ASSETS / "chart_confusion.png", M, y + 0.02, 4.62, 2.72)
    image_fit(s, ASSETS / "chart_cv.png", M + 4.80, y + 0.02, 7.31, 2.72)

    yy = y + 2.86
    tiles = [("80.34%", "ACCURACY", TEXT), ("64.74%", "PRECISION", TEXT),
             ("56.95%", "RECALL", WARN), ("60.60%", "F1", OK), ("84.91%", "ROC-AUC", OK)]
    tw = (CW - 4 * 0.16) / 5
    for i, (v, l, c) in enumerate(tiles):
        kpi(s, M + i * (tw + 0.16), yy, tw, 0.90, v, l, c, 18, 9)

    text(s, M, yy + 1.00, CW, 0.50,
         [{"t": "Honest reading: of 374 churners in the test set the model catches 213. "
                "The deployed 60.60% F1 sits inside a cross-validated range of "
                "62.11% ± 2.96% (accuracy 81.19% ± 1.18%), so the single split is not "
                "carrying the result — and recall is the number this project works on next.",
           "pt": 10.8, "color": MUTED, "space_after": 0, "spacing": 1.16}])
    footer(s, n, phase=6)
    notes(s, """
(5:05–5:55) The credibility slide — slow down.
"1,409 customers held back entirely. 919 stayed and were correctly left alone;
116 were flagged unnecessarily; 161 churners were missed; 213 caught."
Then the stability point, which most projects never show: "five-fold
cross-validation gives 62.11% F1 with a 2.96 standard deviation, and my deployed
result sits inside it. So this is a stable result, not a lucky split."
Then own the weakness: "recall is 56.95% — that is the one number I would fix
first, and I'll show you exactly how on the next slide."
""")

    # =========================================================== 10 PHASE 08a
    n += 1
    s = add_slide(prs)
    eyebrow(s, "PHASE 08  ·  MODEL SELECTION")
    y = heading(s, "Why Logistic Regression ships",
                "Four models, five metrics, one decision — made on F1 and ROC-AUC")
    rows = [
        ("Model", "Accuracy", "Precision", "Recall", "F1", "ROC-AUC", True),
        ("Logistic Regression  ← selected", "80.34%", "64.74%", "56.95%", "60.60%", "84.91%", False),
        ("LightGBM", "80.06%", "64.49%", "55.35%", "59.57%", "84.71%", False),
        ("XGBoost", "79.06%", "61.72%", "55.61%", "58.51%", "83.34%", False),
        ("Random Forest", "78.99%", "62.34%", "52.67%", "57.10%", "83.48%", False),
    ]
    colw = [4.35, 1.56, 1.56, 1.42, 1.42, 1.79]
    yy = y + 0.06
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
                   "mono": hdr or ci > 0}], audit=False)
            x += colw[ci]
        yy += 0.60

    callout(s, M, yy + 0.14, CW, 1.30, "THE DECISION RULE",
            "Accuracy alone would have picked LightGBM by one hundredth of a point. The "
            "choice was made on F1 and ROC-AUC, because a retention team cares about "
            "ranking risk correctly and catching churners — not about the share of all "
            "customers classified correctly. Logistic Regression also wins on being "
            "explainable, which the next phase depends on.", OK)
    footer(s, n, phase=7)
    notes(s, """
(5:55–6:25) "Four algorithms, identical treatment. Look at how close they are —
LightGBM is within 0.28 of a point on accuracy. If I had picked on accuracy I'd
have picked a harder model to explain, for no real gain."
Then state the rule: "I selected on F1 and ROC-AUC, because what matters
operationally is ranking risk, not classifying the majority correctly."
""")

    # =========================================================== 11 PHASE 08b
    n += 1
    s = add_slide(prs)
    eyebrow(s, "PHASE 08  ·  MODEL SELECTION  ·  THE OPERATING POINT")
    y = heading(s, "The cut-off was a default — not a decision",
                "0.5 is arbitrary. Measuring it turns the model's weakest number into a choice")
    image_fit(s, ASSETS / "chart_threshold.png", M, y + 0.02, 7.90, 4.30)

    rect(s, 8.72, y + 0.02, 3.99, 2.04, fill=PANEL, edge=LINE)
    text(s, 8.92, y + 0.16, 3.63, 1.74,
         [{"t": "WHAT THE SWEEP SHOWS", "pt": 10.5, "color": OK, "bold": True,
           "mono": True, "space_after": 6},
          {"t": "At 0.35 the model catches 71.7% of churners instead of 57.0%, and F1 "
                "actually improves to 62.98%. Both recall and F1 are better than the "
                "deployed setting.", "pt": 10.6, "color": TEXT, "space_after": 0,
           "spacing": 1.18}])

    callout(s, 8.72, y + 2.20, 3.99, 2.12,
            "THE PROOF THAT MATTERS",
            "Training with class_weight='balanced' lifts recall to 77.81% while ROC-AUC "
            "barely moves: 84.89 against 84.90. The model already ranks customers "
            "correctly — the 0.5 cut-off was throwing churners away.", WARN)
    footer(s, n, phase=7)
    notes(s, """
(6:25–7:05) This is the strongest slide in the deck — give it time.
"0.5 is a library default, not a business decision. So I measured the whole range."
Point at the crossing: "at 0.35 I catch 71.7% of churners instead of 57% — and F1
actually goes UP, not down."
Then the diagnostic proof, which is the real insight: "I retrained with balanced
class weights. Recall goes from 56.95 to 77.81, and ROC-AUC doesn't move —
84.89 versus 84.90. The model was already ranking customers correctly. Only the
cut-off was discarding them."
Then be honest: "it stays at 0.5 in production today because moving it is a
business decision about the cost of a wasted offer versus a lost customer — not
something I should silently change."
""")

    # =========================================================== 12 PHASE 09
    n += 1
    s = add_slide(prs)
    eyebrow(s, "PHASE 09  ·  EXPLAINABILITY")
    y = heading(s, "Why the model flags a customer",
                "The profile of the high-risk band, and a per-customer explanation in the app")
    image_fit(s, ASSETS / "chart_high_risk_profile.png", M, y + 0.02, 7.55, 3.78)
    rect(s, 8.40, y + 0.04, 4.31, 1.90, fill=PANEL, edge=LINE)
    text(s, 8.60, y + 0.18, 3.95, 1.66,
         [{"t": "PER-CUSTOMER EXPLANATION", "pt": 10.5, "color": OK, "bold": True,
           "mono": True, "space_after": 6},
          {"t": "For each prediction the console shows the top drivers, computed as "
                "coefficient × customer value — the honest linear explanation for a "
                "logistic model, with human labels like \"Fiber optic internet\" or "
                "\"Account tenure\".", "pt": 10.6, "color": TEXT, "space_after": 0,
           "spacing": 1.18}])
    callout(s, 8.40, y + 2.06, 4.31, 1.74, "ON SHAP — STATED PRECISELY",
            "SHAP was explored during the project; summary, bar, dependence and waterfall "
            "plots are in the repository. Production explanation is aligned to the "
            "deployed Logistic Regression instead — accurate for a linear model, and it "
            "costs nothing at request time.", WARN)
    footer(s, n, phase=8)
    notes(s, """
(7:05–7:40) Left chart: "this is who the high-risk band is — every one of them is
month-to-month, 91% are on fiber optic, 95% have no online security, and they
average under ten months of tenure."
Right: "and for a single customer the app names which of those factors pushed
this particular score."
If asked about SHAP: "I ran the SHAP analysis during experimentation; the
deployed model is linear, so I explain it with coefficients — same information,
no runtime cost." Do not overclaim here.
""")

    # =========================================================== 13 PHASE 10
    n += 1
    s = add_slide(prs)
    eyebrow(s, "PHASE 10  ·  RISK SEGMENTATION")
    y = heading(s, "Turning a probability into a work queue",
                "Every scored customer lands in one of three operational bands")
    bands = [("< 30%", "LOW RISK", "4,435 customers", OK),
             ("30 – 60%", "MEDIUM RISK", "1,473 customers", WARN),
             ("≥ 60%", "HIGH RISK", "1,135 customers", BAD)]
    bw = (CW - 2 * 0.24) / 3
    for i, (rng, label, count, c) in enumerate(bands):
        x = M + i * (bw + 0.24)
        rect(s, x, y + 0.02, bw, 0.94, fill=PANEL, edge=c, edge_w=1.2)
        text(s, x + 0.22, y + 0.14, bw - 0.44, 0.72,
             [{"t": f"{rng}   {label}", "pt": 11.5, "color": c, "bold": True,
               "mono": True, "space_after": 2},
              {"t": count, "pt": 12.5, "color": TEXT, "space_after": 0}])
    image_fit(s, ASSETS / "chart_risk_bands.png", M, y + 1.04, 7.72, 3.36)
    rect(s, 8.62, y + 1.06, 4.09, 1.58, fill=PANEL_2, edge=OK, edge_w=1.2)
    text(s, 8.82, y + 1.20, 3.73, 1.34,
         [{"t": "VALIDATION", "pt": 10.5, "color": OK, "bold": True, "mono": True,
           "space_after": 5},
          {"t": "The bands track reality: actual churn is 9.5% in low risk, 42.0% in "
                "medium and 73.0% in high risk. The segmentation is not decoration — it "
                "orders the portfolio correctly.", "pt": 10.6, "color": TEXT,
           "space_after": 0, "spacing": 1.18}])
    rect(s, 8.62, y + 2.78, 4.09, 1.62, fill=PANEL, edge=BAD, edge_w=1.2)
    text(s, 8.82, y + 2.92, 3.73, 1.38,
         [{"t": "THE OPERATIONAL PAYOFF", "pt": 10.5, "color": BAD, "bold": True,
           "mono": True, "space_after": 5},
          {"t": "Instead of calling 7,043 customers, the team starts with 1,135 — 16.1% "
                "of the base, where 73% will actually leave — carrying $92,539 of "
                "recurring monthly charges.", "pt": 10.6, "color": TEXT, "space_after": 0,
           "spacing": 1.18}])
    footer(s, n, phase=9)
    notes(s, """
(7:40–8:15) "A probability isn't something a retention agent can act on, so I cut
it into three bands."
Then the validation, which is the point of the slide: "I checked the bands against
the real churn labels. Low risk actually churns 9.5%, medium 42%, high 73%. The
band is a genuine ordering of risk, not a cosmetic split."
Close operationally: "the team starts with 1,135 customers instead of 7,043."
""")

    # =========================================================== 14 PHASE 11
    n += 1
    s = add_slide(prs)
    eyebrow(s, "PHASE 11  ·  RETENTION INTELLIGENCE")
    y = heading(s, "From risk to action — and to money",
                "2,608 at-risk customers each received a specific recommendation")
    image_fit(s, ASSETS / "chart_revenue.png", M, y + 0.04, 7.10, 3.04)
    rect(s, 7.95, y + 0.02, 4.76, 3.08, fill=PANEL, edge=LINE)
    text(s, 8.15, y + 0.16, 4.40, 2.84,
         [{"t": "RECOMMENDATIONS ISSUED (FULL PORTFOLIO)", "pt": 10.5, "color": OK,
           "bold": True, "mono": True, "space_after": 7},
          {"t": "Promote Long-Term Contract            792", "pt": 10.4, "color": TEXT, "space_after": 3, "mono": True},
          {"t": "5% Discount Offer                            663", "pt": 10.4, "color": TEXT, "space_after": 3, "mono": True},
          {"t": "Offer 15% Discount                          658", "pt": 10.4, "color": TEXT, "space_after": 3, "mono": True},
          {"t": "Welcome Retention Package          430", "pt": 10.4, "color": TEXT, "space_after": 3, "mono": True},
          {"t": "Free Online Security                       45", "pt": 10.4, "color": TEXT, "space_after": 3, "mono": True},
          {"t": "Switch to Autopay / Check-in          18", "pt": 10.4, "color": TEXT, "space_after": 3, "mono": True},
          {"t": "Free Premium Support                      2", "pt": 10.4, "color": TEXT, "space_after": 8, "mono": True},
          {"t": "Driven by each customer's own risk factors — a month-to-month customer "
                "with no security gets a different offer than a loyal high-value one.",
           "pt": 10.2, "color": MUTED, "space_after": 0, "spacing": 1.16}])
    yy = y + 3.24
    callout(s, M, yy, 7.10, 1.20, "SPEAK PRECISELY ABOUT THIS NUMBER",
            "RetainIQ reports monthly revenue at risk — the recurring charges sitting "
            "inside flagged customers. $92,539/month in high risk, $200,302 including "
            "medium. It is exposure, not money already saved.", WARN)
    rect(s, 7.95, yy, 4.76, 1.20, fill=PANEL, edge=LINE)
    text(s, 8.15, yy + 0.14, 4.40, 0.98,
         [{"t": "DECISION FRAMEWORK", "pt": 10.5, "color": OK, "bold": True,
           "mono": True, "space_after": 5},
          {"t": "Data → insight → prediction → explanation → action → business value. "
                "The platform deliberately does not stop at step three.", "pt": 10.4,
           "color": TEXT, "space_after": 0, "spacing": 1.18}])
    footer(s, n, phase=10)
    notes(s, """
(8:15–8:50) "Prediction alone doesn't retain anyone."
Show the offer list — note the biggest group is a contract-upgrade nudge, which
follows directly from the Phase 04 finding about month-to-month risk.
Then be precise about money: "this is monthly revenue at risk — exposure, $92,539
a month sitting in the high-risk band. I deliberately do not call it savings."
""")

    # =========================================================== 15 PHASE 12
    n += 1
    s = add_slide(prs)
    eyebrow(s, "PHASE 12  ·  BUSINESS INTELLIGENCE")
    y = heading(s, "The descriptive layer beneath the predictions",
                "A four-page Power BI dashboard for the people who fund the retention team")
    picture_card(s, PBI / "page1_executive_overview.png.png", M, y + 0.04, 6.30, 2.36)
    picture_card(s, PBI / "page3_risk_intelligence.png.png", M + 6.48, y + 0.04, 5.61, 2.36)

    yy = y + 2.56
    cards = [
        ("Executive overview", "Portfolio KPIs, churn rate and revenue in one view for management.", OK),
        ("Customer insights", "Where churn concentrates — contract, tenure, service mix.", ACTION),
        ("Risk intelligence", "The scored risk distribution carried into BI.", WARN),
        ("Retention strategy", "Which actions are recommended, and to how many customers.", BAD),
    ]
    cw = (CW - 3 * 0.18) / 4
    for i, (t, b, c) in enumerate(cards):
        x = M + i * (cw + 0.18)
        rect(s, x, yy, cw, 1.62, fill=PANEL, edge=LINE)
        bar(s, x, yy + 0.14, 0.035, 1.34, c)
        text(s, x + 0.20, yy + 0.16, cw - 0.42, 1.32,
             [{"t": t, "pt": 11.5, "color": c, "bold": True, "space_after": 4},
              {"t": b, "pt": 10.0, "color": TEXT, "space_after": 0, "spacing": 1.16}])
    footer(s, n, phase=11)
    notes(s, """
(8:50–9:15) "The machine learning app answers what is likely to happen. Power BI
answers what already happened and why — and it's what a non-technical
stakeholder actually opens."
Four pages: executive overview, customer insights, risk intelligence, retention
strategy. One sentence each, don't linger.
""")

    # =========================================================== 16 PHASE 13
    n += 1
    s = add_slide(prs)
    eyebrow(s, "PHASE 13  ·  WEB APPLICATION")
    y = heading(s, "The product: a Signal Ops Console",
                "A working Flask application — bulk scoring, dashboard, private reports")
    image_fit(s, ASSETS / "chart_architecture.png", M, y - 0.02, CW, 2.20)
    yy = y + 2.28
    ph = 2.58
    image_fit(s, IMAGES / "06.prediction_dashboard.png", M, yy, 5.40, ph)
    image_fit(s, IMAGES / "04.Bulk_prediction.png", M + 5.58, yy, 3.90, ph)
    rect(s, M + 9.66, yy, 2.43, ph, fill=PANEL, edge=LINE)
    text(s, M + 9.82, yy + 0.12, 2.11, ph - 0.24,
         [{"t": "WHAT A USER DOES", "pt": 9.5, "color": OK, "bold": True, "mono": True,
           "space_after": 5},
          {"t": "1  Upload a CSV", "pt": 9.3, "color": TEXT, "space_after": 3},
          {"t": "2  Validate the file", "pt": 9.3, "color": TEXT, "space_after": 3},
          {"t": "3  Score the portfolio", "pt": 9.3, "color": TEXT, "space_after": 3},
          {"t": "4  Read the KPIs", "pt": 9.3, "color": TEXT, "space_after": 3},
          {"t": "5  Drill into a customer", "pt": 9.3, "color": TEXT, "space_after": 3},
          {"t": "6  Download the report", "pt": 9.3, "color": TEXT, "space_after": 7},
          {"t": "Reports are scoped to the visitor by a signed cookie — never a shared "
                "folder.", "pt": 8.7, "color": MUTED, "space_after": 0, "spacing": 1.12}])
    footer(s, n, phase=12)
    notes(s, """
(9:15–9:45) This is the demo slide. If the live site is warm, switch to it for
thirty seconds; otherwise walk the screenshots.
"Upload a portfolio, validate it before scoring, then read churn rate, risk
distribution and revenue at risk — and drill into any single customer."
One engineering detail worth a sentence: "reports are scoped to the visitor with
a signed cookie, so users never see each other's data — that has its own unit
test."
""")

    # =========================================================== 17 PHASE 14
    n += 1
    s = add_slide(prs)
    eyebrow(s, "PHASE 14  ·  DEPLOYMENT")
    y = heading(s, "Making it survive production",
                "Free-tier hosting is unforgiving — these are the four problems it creates")
    problems = [
        ("Slow scoring on a shared CPU", "A large CSV can take far longer than the default 30-second worker timeout, so the worker is killed mid-request and the browser shows a 502.", "300-second timeout", OK),
        ("Reloading the model per request", "Reading the pickled model on every request wastes the little memory the instance has.", "worker preload", OK),
        ("Memory creeping past the limit", "A long-lived worker grows until the platform OOM-kills it, which surfaces as a random error.", "request recycling", OK),
        ("The free tier sleeps", "An idle instance is suspended, so the first visitor waits for a cold start.", "warm it before demoing", WARN),
    ]
    ch = (CW - 0.26) / 2
    for i, (t, b, fix, c) in enumerate(problems):
        x = M + (i % 2) * (ch + 0.26)
        yy = y + 0.04 + (i // 2) * 1.78
        rect(s, x, yy, ch, 1.62, fill=PANEL, edge=LINE)
        bar(s, x, yy + 0.14, 0.035, 1.34, c)
        text(s, x + 0.22, yy + 0.16, ch - 0.46, 1.30,
             [{"t": t, "pt": 12, "color": TEXT, "bold": True, "space_after": 4},
              {"t": b, "pt": 10.2, "color": MUTED, "space_after": 5, "spacing": 1.16},
              {"t": f"→ fixed by {fix}", "pt": 10.4, "color": c, "bold": True,
               "mono": True, "space_after": 0}])

    callout(s, M, y + 3.66, CW, 1.00, "SHIPPED AS",
            "Gunicorn with a threaded worker · Dockerfile and Procfile · version-pinned "
            "requirements · hosted on Render, where the application answers at the URL on "
            "the closing slide.", ACTION)
    footer(s, n, phase=13)
    notes(s, """
(9:45–10:10) Don't list technology here — that's the next slide. This slide is
about the production problems, which is what actually separates a notebook from
a deployed service.
"Four things broke or nearly broke when this left my laptop: slow scoring hitting
the worker timeout, the model being reloaded per request, memory creeping until
the platform killed the worker, and the free tier sleeping."
Then the honest operational note: "which means if you scan the QR at the end, and
the app has been idle, you'll wait about a minute for it to wake up. I'll warm it
before we start."
""")

    # =========================================================== 18 TECH STACK
    n += 1
    s = add_slide(prs)
    eyebrow(s, "TECHNOLOGY STACK")
    y = heading(s, "Everything the project is built on",
                "One requirement drove the whole stack: the same code must train the "
                "model and serve it")
    groups = [
        ("LANGUAGE & DATA", OK, ["Python 3.11", "Pandas — data manipulation",
                                 "NumPy — numerical computing",
                                 "SQL Server (extraction layer)",
                                 "SQLAlchemy — DB connection"]),
        ("MACHINE LEARNING", WARN, ["Scikit-learn — training & evaluation",
                                    "Logistic Regression (deployed)",
                                    "Random Forest · XGBoost · LightGBM",
                                    "Joblib — model persistence",
                                    "SHAP — explainability research"]),
        ("WEB APPLICATION", ACTION, ["Flask — application server",
                                     "Jinja2 — server-side templates",
                                     "HTML5 · CSS3 (custom design system)",
                                     "Vanilla JavaScript — dashboards",
                                     "Gunicorn — WSGI production server"]),
        ("APIs & DELIVERY", BAD, ["FastAPI + Pydantic — typed API",
                                  "Docker — container image",
                                  "Render — cloud hosting",
                                  "Git & GitHub — version control",
                                  "Microsoft Power BI (.pbix / .pbit)"]),
    ]
    cw = (CW - 3 * 0.20) / 4
    for i, (title, coll, items) in enumerate(groups):
        x = M + i * (cw + 0.20)
        rect(s, x, y + 0.04, cw, 2.62, fill=PANEL, edge=LINE)
        bar(s, x, y + 0.04, cw, 0.045, coll)
        blocks = [{"t": title, "pt": 10, "color": coll, "bold": True, "mono": True,
                   "space_after": 8}]
        blocks += [{"t": f"·  {item}", "pt": 10.4, "color": TEXT, "space_after": 5,
                    "spacing": 1.10} for item in items]
        text(s, x + 0.20, y + 0.24, cw - 0.40, 2.28, blocks)

    callout(s, M, y + 2.82, CW, 1.14, "THE DESIGN CHOICE BEHIND THE STACK",
            "Preprocessing, the saved column contract and the scoring functions are a "
            "single Python package imported by the notebooks, the Flask app, the FastAPI "
            "wrapper and the evaluation scripts. That is what makes training and "
            "production behaviour the same thing rather than two similar things.", OK)
    footer(s, n, complete_through=len(PHASES))
    notes(s, """
(10:10–10:35) Do not read the columns. One sentence each.
"Python and pandas for the data work. Scikit-learn, four algorithms, Logistic
Regression deployed. Flask with a hand-built CSS design system for the product.
FastAPI, Docker and Render around it. Power BI for the business layer."
If asked about a tool that isn't here: only claim what can be opened in the
repository.
""")

    # =========================================================== 19 RESPONSIBLE AI
    n += 1
    s = add_slide(prs)
    eyebrow(s, "RESPONSIBLE AI, LIMITS & THE NEXT CYCLE")
    y = heading(s, "What this model should not be trusted with",
                "Where it breaks, who it might treat unfairly, and what the next cycle fixes")
    col = (CW - 2 * 0.26) / 3

    rect(s, M, y + 0.02, col, 3.62, fill=PANEL, edge=WARN, edge_w=1.2)
    text(s, M + 0.22, y + 0.16, col - 0.44, 3.32,
         [{"t": "LIMITATIONS — STATED PLAINLY", "pt": 11, "color": WARN, "bold": True,
           "mono": True, "space_after": 7},
          {"t": "▸  Trained on historical data from one telecom; behaviour may not transfer",
           "pt": 10.6, "color": TEXT, "space_after": 5, "spacing": 1.16},
          {"t": "▸  56.95% recall at the deployed threshold — two in five churners missed",
           "pt": 10.6, "color": TEXT, "space_after": 5, "spacing": 1.16},
          {"t": "▸  No live behavioural feed, no drift monitoring, no automatic retraining",
           "pt": 10.6, "color": TEXT, "space_after": 5, "spacing": 1.16},
          {"t": "▸  Probabilities are not guarantees; recommendations are decision support",
           "pt": 10.6, "color": TEXT, "space_after": 5, "spacing": 1.16},
          {"t": "▸  Revenue figures are exposure, not realised savings",
           "pt": 10.6, "color": TEXT, "space_after": 0, "spacing": 1.16}])

    rect(s, M + col + 0.26, y + 0.02, col, 3.62, fill=PANEL, edge=BAD, edge_w=1.2)
    text(s, M + col + 0.48, y + 0.16, col - 0.44, 3.32,
         [{"t": "FAIRNESS — AN OPEN QUESTION", "pt": 11, "color": BAD, "bold": True,
           "mono": True, "space_after": 7},
          {"t": "▸  Senior citizens churn at 41.7% — nearly double the base rate. A model "
                "optimising for churn risk will systematically concentrate its retention "
                "offers on them.", "pt": 10.6, "color": TEXT, "space_after": 5,
           "spacing": 1.16},
          {"t": "▸  A blanket discount aimed at that group is not neutral; it can read as "
                "exploitative pricing.", "pt": 10.6, "color": TEXT, "space_after": 5,
           "spacing": 1.16},
          {"t": "▸  No bias audit has been run. Before this touched real customers, that "
                "would be a prerequisite, not a nice-to-have.", "pt": 10.6, "color": TEXT,
           "space_after": 0, "spacing": 1.16}])

    rect(s, M + 2 * (col + 0.26), y + 0.02, col, 3.62, fill=PANEL, edge=ACTION, edge_w=1.2)
    text(s, M + 2 * (col + 0.26) + 0.22, y + 0.16, col - 0.44, 3.32,
         [{"t": "THE NEXT CYCLE", "pt": 11, "color": ACTION, "bold": True, "mono": True,
           "space_after": 7},
          {"t": "▸  Move the operating point deliberately — 0.35 catches 71.7% of churners "
                "at no cost to F1", "pt": 10.6, "color": TEXT, "space_after": 5,
           "spacing": 1.16},
          {"t": "▸  Decide the threshold with the business: a wasted offer is cheap, a lost "
                "customer is not", "pt": 10.6, "color": TEXT, "space_after": 5,
           "spacing": 1.16},
          {"t": "▸  Bias audit across age and other protected attributes",
           "pt": 10.6, "color": TEXT, "space_after": 5, "spacing": 1.16},
          {"t": "▸  Drift detection and scheduled retraining — which sends the project "
                "back to Phase 02", "pt": 10.6, "color": TEXT, "space_after": 0,
           "spacing": 1.16}])
    footer(s, n, complete_through=len(PHASES))
    notes(s, """
(10:35–11:05) Most students hide this slide. Put it up and read it.
"Three things I would not do with this model today. I wouldn't trust it on a
different market — it has only ever seen one telecom. I wouldn't quote revenue as
savings, only as exposure. And I wouldn't ship a retention campaign aimed at
senior citizens without a bias audit — they churn at 41.7%, nearly double the
base rate, so a risk-optimising model will concentrate offers on them."
Then close the loop: "and the fix for the recall number is already measured — 0.35
catches 71.7%. That's the next cycle, which starts back at Phase 02 with drift
monitoring."
""")

    # =========================================================== 20 THANK YOU
    n += 1
    s = add_slide(prs)
    bar(s, 0, 0, SW, 0.085, OK)
    eyebrow(s, "THANK YOU")
    y = heading(s, "See it running",
                "Scan to open the live application and score a customer portfolio")
    rect(s, M, y + 0.06, 5.82, 3.58, fill=PANEL, edge=OK, edge_w=1.4)
    text(s, M + 0.28, y + 0.28, 5.26, 3.16,
         [{"t": "LIVE APPLICATION", "pt": 10.5, "color": OK, "bold": True, "mono": True,
           "space_after": 8},
          {"t": LIVE_URL_DISPLAY, "pt": 11, "color": TEXT, "bold": True,
           "space_after": 10, "spacing": 1.16},
          {"t": "Bulk score a customer CSV, read the dashboard, and download a private "
                "report. The model, risk bands, explanations and recommendations run "
                "live — the same code paths that produced every figure in this deck.",
           "pt": 10.8, "color": MUTED, "space_after": 12, "spacing": 1.20},
          {"t": "SOURCE CODE", "pt": 10.5, "color": ACTION, "bold": True, "mono": True,
           "space_after": 5},
          {"t": REPO_DISPLAY, "pt": 10.6, "color": TEXT, "mono": True, "space_after": 0,
           "spacing": 1.16}])

    rect(s, 6.54, y + 0.06, 2.98, 3.58, fill=WHITE, edge=LINE)
    image_fit(s, ASSETS / "qr_live_app.png", 6.70, y + 0.24, 2.66, 2.66)
    text(s, 6.70, y + 3.00, 2.66, 0.52,
         [{"t": "SCAN FOR THE", "pt": 9.5, "color": INK, "bold": True, "mono": True,
           "space_after": 2, "align": PP_ALIGN.CENTER},
          {"t": "LIVE APPLICATION", "pt": 9.5, "color": INK, "bold": True, "mono": True,
           "space_after": 0, "align": PP_ALIGN.CENTER}])

    rect(s, 9.62, y + 0.06, 3.09, 3.58, fill=PANEL, edge=LINE)
    image_fit(s, ASSETS / "qr_repository.png", 10.54, y + 0.24, 1.25, 1.25)
    text(s, 9.82, y + 0.28, 0.66, 1.00,
         [{"t": "SOURCE", "pt": 8.5, "color": MUTED, "bold": True, "mono": True,
           "space_after": 3},
          {"t": "REPO", "pt": 8.5, "color": MUTED, "bold": True, "mono": True,
           "space_after": 0}], audit=False)
    text(s, 9.82, y + 1.52, 2.69, 2.00,
         [{"t": "Thank you", "pt": 19, "color": TEXT, "bold": True, "space_after": 5},
          {"t": "Questions welcome.", "pt": 11, "color": OK, "space_after": 10},
          {"t": "RetainIQ — from a churn probability to a prioritised, explained, costed "
                "retention action.", "pt": 10.4, "color": MUTED, "space_after": 0,
           "spacing": 1.18}])

    signal_meter(s, 11.05, 5.62, filled=5, scale=0.85, label=None, color=OK)
    text(s, M, y + 3.80, CW, 0.40,
         [{"t": "Logesh S.  ·  B.Sc. Data Science  ·  RetainIQ Signal Ops Console  ·  "
                "80.34% accuracy  ·  84.91% ROC-AUC",
           "pt": 10.5, "color": MUTED, "mono": True, "space_after": 0}])
    notes(s, """
CLOSING (11:05–11:30) Leave this slide up for questions. It puts the live URL and
a scannable QR in front of the examiner and ends on the thesis, not a tool list.
"The contribution isn't the 80% — it's that a probability becomes a prioritised,
explained, costed action, in a product a retention team can actually open. Thank
you — happy to take questions."
Offer the demo: "if you'd like, I can upload a customer file and score it live."
Remember the free tier sleeps — warm the app before you present.

LIKELY QUESTIONS
· Why not deep learning? 1,409 test rows and a largely linear signal; the four-way
  benchmark supports the simpler model, and it stays explainable.
· Why is recall low? The 0.5 threshold is a default. Phase 08 shows 0.35 catches
  71.7%, and balanced class weights reach 77.81% recall at identical ROC-AUC.
· Isn't revenue-at-risk optimistic? It's exposure, not savings — the deck and the
  app both say so; validating it needs a live campaign.
· How do you know it isn't leakage? Churn Score, Reason, Category and Status are
  post-outcome fields and were dropped before training.
""")

    # =========================================================== A1 BENCHMARK
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
    yy = y + 0.20
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
                   "mono": hdr or ci > 0}], audit=False)
            x += colw[ci]
        yy += 0.60

    text(s, M, yy + 0.34, CW, 1.4,
         [{"t": "All four models were trained on the identical 5,634-row training split "
                "with the identical 30-feature frame, and evaluated with the same scoring "
                "code — the only variable is the algorithm.", "pt": 11, "color": MUTED,
           "space_after": 6, "spacing": 1.18},
          {"t": "Reproduce:  python presentation/build_assets.py  →  writes "
                "presentation/assets/model_metrics.csv", "pt": 10.5, "color": OK,
           "mono": True, "space_after": 0}])
    footer(s, n, complete_through=len(PHASES), label="Appendix · use only if asked")
    notes(s, "Backup slide. Use if anyone questions the benchmark or wants the full "
             "metric set.")

    # =========================================================== A2 HYPERPARAMS
    n += 1
    s = add_slide(prs)
    eyebrow(s, "APPENDIX", color=MUTED)
    y = heading(s, "Hyperparameters",
                "Read directly from the saved model artifacts and the training notebook")
    rows = [
        ("Model", "Settings"),
        ("Logistic Regression", "max_iter=5000, C=1.0, solver='lbfgs'"),
        ("Random Forest", "n_estimators=200, random_state=42"),
        ("XGBoost", "XGBClassifier(random_state=42) — library defaults otherwise"),
        ("LightGBM", "LGBMClassifier(random_state=42) — library defaults otherwise"),
    ]
    yy = y + 0.20
    for ri, (a, b) in enumerate(rows):
        hdr = ri == 0
        rect(s, M, yy, 3.60, 0.54, fill=PANEL_2 if hdr else PANEL, edge=LINE)
        rect(s, M + 3.60, yy, 8.49, 0.54, fill=PANEL_2 if hdr else PANEL, edge=LINE)
        text(s, M + 0.16, yy + 0.15, 3.35, 0.30,
             [{"t": a, "pt": 10.8, "color": MUTED if hdr else TEXT, "bold": hdr,
               "space_after": 0}], audit=False)
        text(s, M + 3.76, yy + 0.15, 8.20, 0.30,
             [{"t": b, "pt": 10.2, "color": MUTED if hdr else OK, "mono": not hdr,
               "bold": hdr, "space_after": 0}], audit=False)
        yy += 0.62
    callout(s, M, yy + 0.20, CW, 1.30, "A NOTE ON TUNING",
            "No automated hyperparameter search was run — the algorithms were compared "
            "under their default configurations, which keeps the benchmark a fair test of "
            "the algorithms rather than of how much tuning effort each received. The "
            "XGBoost and LightGBM instances were constructed with only random_state set, "
            "so their remaining settings are library defaults; that is what the saved "
            "artifacts contain. A tuning sweep is on the next-cycle list.", WARN)
    footer(s, n, complete_through=len(PHASES), label="Appendix · use only if asked")
    notes(s, "Backup slide for 'what parameters did you use?'. Be straightforward: these "
             "are defaults, and that was deliberate so the comparison tested algorithms "
             "rather than tuning effort.")

    # =========================================================== A3 FEATURES
    n += 1
    s = add_slide(prs)
    eyebrow(s, "APPENDIX", color=MUTED)
    y = heading(s, "The 30 model inputs",
                "Three numeric measures plus twenty-seven encoded flags, in model order")
    import joblib
    features = list(joblib.load(ROOT / "models" / "feature_columns.pkl"))
    half = (len(features) + 1) // 2
    for col_i, chunk in enumerate([features[:half], features[half:]]):
        x = M + col_i * (CW / 2 + 0.10)
        blocks = []
        start = col_i * half + 1
        for k, feat in enumerate(chunk):
            blocks.append({"t": f"{start + k:>2d}.  {feat}", "pt": 9.4, "color": TEXT,
                           "mono": True, "space_after": 1.2})
        text(s, x, y + 0.10, CW / 2 - 0.10, 4.4, blocks)
    footer(s, n, complete_through=len(PHASES), label="Appendix · use only if asked")
    notes(s, "Backup slide. Any question about which variables the model sees can be "
             "answered by pointing here.")

    # =========================================================== A4 API + PROVENANCE
    n += 1
    s = add_slide(prs)
    eyebrow(s, "APPENDIX", color=MUTED)
    y = heading(s, "Prediction API & where every figure came from",
                "The same engine, exposed programmatically — and the provenance of this deck")
    rect(s, M, y + 0.04, 5.30, 2.42, fill=PANEL_2, edge=LINE)
    text(s, M + 0.20, y + 0.18, 4.90, 2.14,
         [{"t": "POST /api/predict", "pt": 11, "color": OK, "bold": True, "mono": True,
           "space_after": 7},
          {"t": "{ \"gender\": \"Male\", \"SeniorCitizen\": 1,", "pt": 9.6, "color": TEXT, "space_after": 1, "mono": True},
          {"t": "  \"tenure\": 8, \"Contract\": \"Month-to-month\",", "pt": 9.6, "color": TEXT, "space_after": 1, "mono": True},
          {"t": "  \"InternetService\": \"Fiber optic\", ... }", "pt": 9.6, "color": TEXT, "space_after": 9, "mono": True},
          {"t": "→  probability · prediction · risk_level", "pt": 9.8, "color": OK, "space_after": 1, "mono": True},
          {"t": "→  top_drivers · recommendations", "pt": 9.8, "color": OK, "space_after": 0, "mono": True}])
    rect(s, 6.12, y + 0.04, 6.59, 2.42, fill=PANEL, edge=LINE)
    text(s, 6.32, y + 0.18, 6.19, 2.14,
         [{"t": "FIGURE PROVENANCE", "pt": 11, "color": OK, "bold": True, "mono": True,
           "space_after": 7},
          {"t": "Churn split, contract, internet, tenure  →  Data/01.Telco_customer_churn_Dataset.csv",
           "pt": 9.8, "color": TEXT, "space_after": 4, "spacing": 1.14},
          {"t": "Benchmark, confusion matrix, ROC  →  models/*.pkl on the held-out split",
           "pt": 9.8, "color": TEXT, "space_after": 4, "spacing": 1.14},
          {"t": "Threshold sweep, CV, imbalance  →  presentation/build_assets.py",
           "pt": 9.8, "color": TEXT, "space_after": 4, "spacing": 1.14},
          {"t": "Risk bands, profile, revenue  →  the production scoring pipeline",
           "pt": 9.8, "color": TEXT, "space_after": 4, "spacing": 1.14},
          {"t": "QR codes  →  build_qr.py, decoded back to verify", "pt": 9.8,
           "color": TEXT, "space_after": 0, "spacing": 1.14}])
    callout(s, M, y + 2.62, CW, 1.10, "THE POINT",
            "Nothing in this deck is hand-typed. build_assets.py recomputes every figure "
            "from the saved artifacts and the production scoring path, so the deck cannot "
            "drift away from the project — retrain a model and the charts and the "
            "benchmark table regenerate together.", OK)
    footer(s, n, complete_through=len(PHASES), label="Appendix · use only if asked")
    notes(s, "Backup slide for API and reproducibility questions. Strong answer to "
             "'how do we know these numbers are right?'")

    prs.save(OUT)
    return OUT


def main():
    out = build()
    print(f"wrote {out.relative_to(ROOT)}  ({out.stat().st_size / 1024:.0f} KB)")
    print(f"slides: {len(Presentation(str(out)).slides)}")
    if WARNINGS:
        print(f"\n{len(WARNINGS)} layout warning(s):")
        for w in WARNINGS[:25]:
            print("  -", w)
    else:
        print("\nno layout overflow detected")


if __name__ == "__main__":
    main()
