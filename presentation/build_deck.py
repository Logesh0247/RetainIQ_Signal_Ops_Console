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
REPO_URL = "https://github.com/Logesh0247/RetainIQ_Signal_Ops_Console"
PBIT_URL = ("https://github.com/Logesh0247/RetainIQ_Signal_Ops_Console/blob/main/"
            "Power_BI_dashboard/dashboard_template.pbit")

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

# --- vertical layout engine -------------------------------------------------
# Every slide's content must occupy exactly this band, so no slide ends early.
CONTENT_TOP = 1.66
CONTENT_BOTTOM = 6.80
CAV = CONTENT_BOTTOM - CONTENT_TOP          # 5.14in of usable height


def stack(rows, top=CONTENT_TOP, bottom=CONTENT_BOTTOM):
    """rows: [(height, gap_after), ...] -> [(y, h), ...] filling top..bottom.

    Heights are treated as relative weights, so a slide designed to a nominal
    height self-corrects to fill the band exactly.
    """
    total_gaps = sum(g for _, g in rows)
    avail = (bottom - top) - total_gaps
    total_w = sum(h for h, _ in rows)
    out, y = [], top
    for h, g in rows:
        hh = avail * (h / total_w)
        out.append((y, hh))
        y += hh + g
    return out


def img_h(path, w):
    """Height an image occupies at a given width, preserving aspect ratio."""
    with Image.open(path) as im:
        iw, ih = im.size
    return w * ih / iw


def img_w(path, h):
    with Image.open(path) as im:
        iw, ih = im.size
    return h * iw / ih


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
        if item.get("href"):
            run.hyperlink.address = item["href"]
            run.font.underline = item.get("underline", True)

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
    text(s, M, 1.22, 9.6, 0.28,
         [{"t": "FINAL YEAR PROJECT   ·   B.SC. DATA SCIENCE", "pt": 11,
           "color": OK, "bold": True, "mono": True, "space_after": 0}])
    text(s, M, 1.60, 10.4, 1.68,
         [{"t": "RetainIQ", "pt": 54, "color": TEXT, "bold": True, "space_after": 2},
          {"t": "Signal Ops Console", "pt": 28, "color": OK, "bold": True,
           "space_after": 0}])
    bar(s, M, 3.42, 2.1, 0.035, OK)
    text(s, M, 3.70, 8.1, 1.0,
         [{"t": "An explainable machine learning platform for customer churn "
                "prediction, risk intelligence and retention strategy.",
           "pt": 14.5, "color": MUTED, "space_after": 0, "spacing": 1.25}])

    signal_meter(s, 11.05, 1.78, filled=5, scale=1.05,
                 label="signal strength", color=OK)

    _rk = stack([(1.18, 0.14), (0.84, 0)], top=4.92, bottom=6.84)
    (r_kpi, r_cred) = _rk[0], _rk[1]
    kpi(s, M, r_kpi[0], 2.36, r_kpi[1], "80.34%", "TEST ACCURACY", OK, 27, 9.5)
    kpi(s, M + 2.50, r_kpi[0], 2.36, r_kpi[1], "84.91%", "ROC-AUC", OK, 27, 9.5)
    kpi(s, M + 5.00, r_kpi[0], 2.36, r_kpi[1], "7,043", "CUSTOMERS SCORED", ACTION, 27, 9.5)
    kpi(s, M + 7.50, r_kpi[0], 4.59, r_kpi[1], "$92,539",
        "MONTHLY REVENUE AT RISK IN THE HIGH-RISK BAND", BAD, 27, 9)

    text(s, M, r_cred[0], 12.1, r_cred[1],
         [{"t": "Logesh S.   ·   B.Sc. Data Science", "pt": 11.5, "color": TEXT,
           "bold": True, "space_after": 2},
          {"t": f"LIVE APP  ↗  {LIVE_URL_DISPLAY}", "pt": 9.8, "color": ACTION,
           "mono": True, "space_after": 2, "href": LIVE_URL},
          {"t": f"SOURCE  ↗  {REPO_DISPLAY}", "pt": 9.8, "color": MUTED,
           "mono": True, "space_after": 0, "href": REPO_URL}])
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
    _r = stack([(4.38, 0.08), (0.68, 0)]); r_chart, r_cap = _r[0], _r[1]
    image_fit(s, ASSETS / "chart_cycle.png", M, r_chart[0], CW, img_h(ASSETS / "chart_cycle.png", CW))
    callout(s, M, r_cap[0], CW, r_cap[1], "HOW TO READ THE DECK",
            "Three acts — the signal fades (I), reading the signal (II), the signal "
            "recovered (III). Every slide names its phase in the footer.", OK)
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
    _r = stack([(2.40, 0.20), (2.54, 0)]); r_bul, r_cards = _r[0], _r[1]
    bullets(s, M, r_bul[0], 6.55, r_bul[1], [
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
    card_h = (r_cards[1] - 0.14) / 2
    for i, (q, sub, color) in enumerate(qs):
        x = M + (i % 2) * 3.34
        yy = r_cards[0] + (i // 2) * (card_h + 0.14)
        rect(s, x, yy, 3.16, card_h, fill=PANEL, edge=LINE)
        bar(s, x, yy + 0.16, 0.035, card_h - 0.32, color)
        text(s, x + 0.22, yy + 0.18, 2.82, card_h - 0.34,
             [{"t": q, "pt": 12, "color": TEXT, "bold": True, "space_after": 3},
              {"t": sub, "pt": 9.8, "color": MUTED, "space_after": 0, "spacing": 1.14}])

    text(s, 7.86, CONTENT_TOP + 0.20, 4.6, 0.30,
         [{"t": "THE BASE RATE", "pt": 10.5, "color": MUTED, "bold": True, "mono": True,
           "space_after": 0}])
    chart_h = CAV - 0.44 - 0.66 - 0.14
    image_fit(s, ASSETS / "chart_churn_split.png", 7.72, CONTENT_TOP + 0.48, 4.89, chart_h)
    text(s, 7.86, CONTENT_TOP + CAV - 0.66, 4.6, 0.66,
         [{"t": "1,869 churned · 5,174 retained of 7,043. Predicting \"nobody churns\" is "
                "already 73.46% accurate — which is why accuracy alone decides nothing here.",
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
    _r = stack([(1.18, 0.20), (3.76, 0)]); r_tile, r_body = _r[0], _r[1]
    for i, (v, l, c) in enumerate(tiles):
        kpi(s, M + i * (tw + 0.16), r_tile[0], tw, r_tile[1], v, l, c, 21, 9)

    yy = r_body[0]
    text(s, M, yy, 5.9, r_body[1],
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

    callout(s, 6.85, yy, 5.86, r_body[1] / 2 - 0.08, "CLASS BALANCE AND THE SPLIT",
            "1,869 churned against 5,174 retained. The split is stratified, so both "
            "halves carry the identical 26.54% churn rate: 5,634 rows train, 1,409 "
            "rows are held back and touched only once.", ACTION)
    callout(s, 6.85, yy + r_body[1] / 2 + 0.08, 5.86, r_body[1] / 2 - 0.08, "TRAIN / SERVE PARITY",
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
    _r = stack([(1.30, 0.20), (1.50, 0.16), (1.98, 0)]); r_flow, r_leak, r_bul = _r[0], _r[1], _r[2]
    for i, (t, sub, c) in enumerate(steps):
        x = M + i * (sw_ + 0.30)
        rect(s, x, r_flow[0], sw_, r_flow[1], fill=PANEL_2, edge=c, edge_w=1.2)
        text(s, x + 0.20, r_flow[0] + 0.22, sw_ - 0.40, r_flow[1] - 0.40,
             [{"t": t, "pt": 13.5, "color": TEXT, "bold": True, "space_after": 4},
              {"t": sub, "pt": 10, "color": MUTED, "space_after": 0, "spacing": 1.14}])
        if i < 3:
            text(s, x + sw_ + 0.02, r_flow[0] + r_flow[1] / 2 - 0.20, 0.28, 0.4,
                 [{"t": "→", "pt": 17, "color": OK, "bold": True, "space_after": 0,
                   "align": PP_ALIGN.CENTER}], audit=False)

    yy = r_leak[0]
    callout(s, M, yy, CW, r_leak[1],
            "LEAKAGE PREVENTION — the detail that makes the results believable",
            "Churn Score, Churn Reason, Churn Category and Customer Status were removed "
            "before training. All four are only known after a customer has already "
            "churned, so keeping them would inflate the metrics and break the model in "
            "production. This is the difference between a demo and a deployable model.")
    bullets(s, M, r_bul[0], 12.09, r_bul[1], [
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
    _r = stack([(3.40, 0.16), (1.58, 0)]); r_img, r_card = _r[0], _r[1]
    box_h = r_img[1]
    for i, chart in enumerate(["chart_contract.png", "chart_tenure.png"]):
        cw_ = min(5.95, img_w(ASSETS / chart, box_h))
        image_fit(s, ASSETS / chart, M + i * 6.15, r_img[0], cw_, box_h)
    yy = r_card[0]
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
        rect(s, x, yy, cw, r_card[1], fill=PANEL, edge=LINE)
        bar(s, x, yy + 0.16, 0.035, r_card[1] - 0.32, c)
        text(s, x + 0.20, yy + 0.18, cw - 0.40, r_card[1] - 0.34,
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
    _r = stack([(1.24, 0.20), (3.70, 0)]); r_tile, r_panel = _r[0], _r[1]
    kpi(s, M, r_tile[0], 3.92, r_tile[1], "3", "NUMERIC FEATURES", ACTION, 28, 9.5)
    kpi(s, M + 4.10, r_tile[0], 3.92, r_tile[1], "27", "BINARY FLAGS", WARN, 28, 9.5)
    kpi(s, M + 8.20, r_tile[0], 3.89, r_tile[1], "30", "TOTAL MODEL INPUTS", OK, 28, 9.5)

    yy = r_panel[0]
    rect(s, M, yy, 5.95, r_panel[1], fill=PANEL, edge=LINE)
    text(s, M + 0.22, yy + 0.18, 5.51, r_panel[1] - 0.36,
         [{"t": "THE THREE NUMERIC MEASURES", "pt": 10.5, "color": ACTION, "bold": True,
           "mono": True, "space_after": 6},
          {"t": "Tenure Months, Monthly Charges, Total Charges — the only continuous "
                "quantities the model sees.", "pt": 10.8, "color": TEXT, "space_after": 12,
           "spacing": 1.18},
          {"t": "WHY FLAGS, NOT NUMBERS", "pt": 10.5, "color": WARN, "bold": True,
           "mono": True, "space_after": 6},
          {"t": "Assigning 1, 2, 3 to DSL / Fiber / None would invent an ordering that "
                "does not exist. Flags keep the model honest about what these categories "
                "actually are.", "pt": 10.8, "color": TEXT, "space_after": 12,
           "spacing": 1.18},
          {"t": "Eleven Total Charges values are blank — every one of them belongs to a "
                "zero-tenure customer, so the missing rows never hide a churn signal.",
           "pt": 10.8, "color": MUTED, "space_after": 0, "spacing": 1.18}])

    rect(s, 6.75, yy, 5.96, r_panel[1], fill=PANEL_2, edge=OK, edge_w=1.2)
    text(s, 6.95, yy + 0.18, 5.56, r_panel[1] - 0.36,
         [{"t": "THE COLUMN CONTRACT", "pt": 10.5, "color": OK, "bold": True,
           "mono": True, "space_after": 6},
          {"t": "The exact 30 column names are saved to feature_columns.pkl alongside the "
                "model. The app rebuilds a feature frame and checks it against that list "
                "before scoring.", "pt": 10.8, "color": TEXT, "space_after": 12,
           "spacing": 1.18},
          {"t": "If the two ever disagree, the request fails loudly rather than silently "
                "scoring on a misaligned frame — the most common way a student project "
                "works in the notebook and quietly breaks in the app.", "pt": 10.8,
           "color": MUTED, "space_after": 10, "spacing": 1.18},
          {"t": "Uploads are validated against that same list, so a mismatched CSV is "
                "rejected before a single row is scored.", "pt": 10.8, "color": OK,
           "space_after": 0, "spacing": 1.18}])
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
    _r = stack([(4.02, 0.16), (0.96, 0)]); r_row, r_call = _r[0], _r[1]
    ch_w = min(7.11, img_w(ASSETS / "chart_models.png", r_row[1]))
    image_fit(s, ASSETS / "chart_models.png", M, r_row[0], ch_w, r_row[1])
    bullets(s, M + ch_w + 0.28, r_row[0] + 0.06, CW - ch_w - 0.28, r_row[1] - 0.12, [
        ("Logistic Regression", "linear baseline; interpretable coefficients"),
        ("Random Forest", "bagged trees; captures non-linearity"),
        ("XGBoost & LightGBM", "gradient boosting; strong on tabular data"),
    ], pt=10.8, gap=8)
    rx = M + ch_w + 0.28
    rw = CW - ch_w - 0.28
    rect(s, rx, r_row[0] + r_row[1] - 1.14, rw, 1.14, fill=PANEL_2, edge=WARN, edge_w=1.2)
    text(s, rx + 0.18, r_row[0] + r_row[1] - 1.02, rw - 0.36, 0.96,
         [{"t": "Benchmark ladder (F1)", "pt": 10.2, "color": WARN, "bold": True,
           "space_after": 3},
          {"t": "LR 60.6 · LightGBM 59.6 · XGBoost 58.5 · RF 57.1",
           "pt": 10.2, "color": TEXT, "space_after": 0, "mono": True}])
    callout(s, M, r_call[0], CW, r_call[1], "WHY FOUR MODELS AND NOT ONE",
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
    _r = stack([(2.88, 0.14), (1.16, 0.10), (0.86, 0)])
    r_img, r_tile, r_txt = _r[0], _r[1], _r[2]
    ih = r_img[1]
    w1 = img_w(ASSETS / "chart_confusion.png", ih)
    w2 = img_w(ASSETS / "chart_cv.png", ih)
    scale = min(1.0, (CW - 0.20) / (w1 + w2))
    image_fit(s, ASSETS / "chart_confusion.png", M, r_img[0], w1 * scale, ih * scale)
    image_fit(s, ASSETS / "chart_cv.png", M + w1 * scale + 0.20, r_img[0],
              w2 * scale, ih * scale)

    yy = r_tile[0]
    tiles = [("80.34%", "ACCURACY", TEXT), ("64.74%", "PRECISION", TEXT),
             ("56.95%", "RECALL", WARN), ("60.60%", "F1", OK), ("84.91%", "ROC-AUC", OK)]
    tw = (CW - 4 * 0.16) / 5
    for i, (v, l, c) in enumerate(tiles):
        kpi(s, M + i * (tw + 0.16), yy, tw, r_tile[1], v, l, c, 20, 9)

    text(s, M, r_txt[0], CW, r_txt[1],
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
    _r = stack([(3.70, 0.16), (1.28, 0)]); r_tbl, r_call = _r[0], _r[1]
    row_h = (r_tbl[1] - 0.06) / 5
    yy = r_tbl[0]
    for ri, row in enumerate(rows):
        x = M
        hdr = row[-1]
        for ci, cell in enumerate(row[:-1]):
            if ri == 0:
                rect(s, x, yy, colw[ci], row_h, fill=PANEL_2, edge=LINE)
            elif ri == 1:
                rect(s, x, yy, colw[ci], row_h, fill=PANEL, edge=OK, edge_w=1.2)
            else:
                rect(s, x, yy, colw[ci], row_h, fill=PANEL, edge=LINE)
            color = MUTED if hdr else (OK if ri == 1 and ci > 0 else TEXT)
            text(s, x + 0.14, yy + (row_h - 0.32) / 2, colw[ci] - 0.24, 0.32,
                 [{"t": cell, "pt": 11.5 if not hdr else 10.5, "color": color,
                   "bold": hdr or (ri == 1 and ci == 0), "space_after": 0,
                   "mono": hdr or ci > 0}], audit=False)
            x += colw[ci]
        yy += row_h + 0.015

    callout(s, M, r_call[0], CW, r_call[1], "THE DECISION RULE",
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
    ch_w = 8.40
    image_fit(s, ASSETS / "chart_threshold.png", M, CONTENT_TOP, ch_w,
              img_h(ASSETS / "chart_threshold.png", ch_w))
    px = M + ch_w + 0.22
    pw = SW - M - px
    _p = stack([(2.49, 0.16), (2.49, 0)]); p1, p2 = _p[0], _p[1]

    rect(s, px, p1[0], pw, p1[1], fill=PANEL, edge=LINE)
    text(s, px + 0.20, p1[0] + 0.16, pw - 0.40, p1[1] - 0.32,
         [{"t": "WHAT THE SWEEP SHOWS", "pt": 10.5, "color": OK, "bold": True,
           "mono": True, "space_after": 6},
          {"t": "At 0.35 the model catches 71.7% of churners instead of 57.0%, and F1 "
                "actually improves to 62.98%. Both recall and F1 are better than the "
                "deployed setting.", "pt": 10.6, "color": TEXT, "space_after": 0,
           "spacing": 1.18}])

    callout(s, px, p2[0], pw, p2[1],
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
    _l = stack([(4.12, 0.14), (0.88, 0)]); lc, lcap = _l[0], _l[1]
    cw_ = 8.00
    image_fit(s, ASSETS / "chart_high_risk_profile.png", M, lc[0], cw_,
              img_h(ASSETS / "chart_high_risk_profile.png", cw_))
    px = M + cw_ + 0.29
    pw = SW - M - px
    _p = stack([(2.49, 0.16), (2.49, 0)]); p1, p2 = _p[0], _p[1]
    rect(s, px, p1[0], pw, p1[1], fill=PANEL, edge=LINE)
    text(s, px + 0.20, p1[0] + 0.16, pw - 0.40, p1[1] - 0.32,
         [{"t": "PER-CUSTOMER EXPLANATION", "pt": 10.5, "color": OK, "bold": True,
           "mono": True, "space_after": 6},
          {"t": "For each prediction the console shows the top drivers, computed as "
                "coefficient × customer value — the honest linear explanation for a "
                "logistic model, with human labels like \"Fiber optic internet\" or "
                "\"Account tenure\".", "pt": 10.6, "color": TEXT, "space_after": 0,
           "spacing": 1.18}])
    callout(s, px, p2[0], pw, p2[1], "ON SHAP — STATED PRECISELY",
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
    _r = stack([(1.00, 0.18), (3.96, 0)]); r_band, r_left = _r[0], _r[1]
    r_panel_top = r_left[0]
    r_panel_h = r_left[1]
    for i, (rng, label, count, c) in enumerate(bands):
        x = M + i * (bw + 0.24)
        rect(s, x, r_band[0], bw, r_band[1], fill=PANEL, edge=c, edge_w=1.2)
        text(s, x + 0.22, r_band[0] + 0.16, bw - 0.44, r_band[1] - 0.32,
             [{"t": f"{rng}   {label}", "pt": 11.5, "color": c, "bold": True,
               "mono": True, "space_after": 2},
              {"t": count, "pt": 12.5, "color": TEXT, "space_after": 0}])
    lw = 7.72
    image_fit(s, ASSETS / "chart_risk_bands.png", M, r_left[0], lw,
              img_h(ASSETS / "chart_risk_bands.png", lw))
    px, pw = 8.62, SW - M - 8.62
    _p = stack([(2.49, 0.16), (2.49, 0)], top=r_panel_top + 0.06,
               bottom=r_panel_top + r_panel_h)
    p1, p2 = _p[0], _p[1]
    rect(s, px, p1[0], pw, p1[1], fill=PANEL_2, edge=OK, edge_w=1.2)
    text(s, px + 0.20, p1[0] + 0.16, pw - 0.40, p1[1] - 0.32,
         [{"t": "VALIDATION", "pt": 10.5, "color": OK, "bold": True, "mono": True,
           "space_after": 5},
          {"t": "The bands track reality: actual churn is 9.5% in low risk, 42.0% in "
                "medium and 73.0% in high risk. The segmentation is not decoration — it "
                "orders the portfolio correctly.", "pt": 10.6, "color": TEXT,
           "space_after": 0, "spacing": 1.18}])
    rect(s, px, p2[0], pw, p2[1], fill=PANEL, edge=BAD, edge_w=1.2)
    text(s, px + 0.20, p2[0] + 0.16, pw - 0.40, p2[1] - 0.32,
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
    _r = stack([(3.34, 0.16), (1.64, 0)]); r_chart, r_call = _r[0], _r[1]
    image_fit(s, ASSETS / "chart_revenue.png", M, r_chart[0], 7.10,
              img_h(ASSETS / "chart_revenue.png", 7.10))
    rect(s, 7.95, CONTENT_TOP, 4.76, CAV, fill=PANEL, edge=LINE)
    text(s, 8.15, CONTENT_TOP + 0.18, 4.40, CAV - 0.36,
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
    callout(s, M, r_call[0], 7.10, r_call[1], "SPEAK PRECISELY ABOUT THIS NUMBER",
            "RetainIQ reports monthly revenue at risk — the recurring charges sitting "
            "inside flagged customers. $92,539/month in high risk, $200,302 including "
            "medium. It is exposure, not money already saved.", WARN)
    rect(s, 7.95, r_call[0], 4.76, r_call[1], fill=PANEL, edge=LINE)
    text(s, 8.15, r_call[0] + 0.16, 4.40, r_call[1] - 0.32,
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
    _r = stack([(3.26, 0.12), (1.30, 0.12), (0.34, 0)])
    r_pbi, r_cards, r_link = _r[0], _r[1], _r[2]
    for i, src in enumerate(["page1_executive_overview.png.png",
                             "page3_risk_intelligence.png.png"]):
        path = PBI / src
        card_h = r_pbi[1]
        card_w = min(6.00, img_w(path, card_h - 0.18) + 0.18)
        x = M + i * 6.30 + (6.00 - card_w) / 2
        picture_card(s, path, x, r_pbi[0], card_w, card_h)
    yy = r_cards[0]
    cards = [
        ("Executive overview", "Portfolio KPIs, churn rate and revenue in one view for management.", OK),
        ("Customer insights", "Where churn concentrates — contract, tenure, service mix.", ACTION),
        ("Risk intelligence", "The scored risk distribution carried into BI.", WARN),
        ("Retention strategy", "Which actions are recommended, and to how many customers.", BAD),
    ]
    cw = (CW - 3 * 0.18) / 4
    for i, (t, b, c) in enumerate(cards):
        x = M + i * (cw + 0.18)
        rect(s, x, yy, cw, r_cards[1], fill=PANEL, edge=LINE)
        bar(s, x, yy + 0.14, 0.035, r_cards[1] - 0.28, c)
        text(s, x + 0.20, yy + 0.16, cw - 0.42, r_cards[1] - 0.30,
             [{"t": t, "pt": 11.5, "color": c, "bold": True, "space_after": 4},
              {"t": b, "pt": 10.0, "color": TEXT, "space_after": 0, "spacing": 1.16}])
    rect(s, M, r_link[0], CW, r_link[1], fill=PANEL_2, edge=ACTION, edge_w=1.0)
    text(s, M + 0.20, r_link[0] + 0.02, CW - 0.40, r_link[1] - 0.04,
         [{"t": "OPEN THE EDITABLE POWER BI TEMPLATE (.PBIT)  ↗", "pt": 10.5,
           "color": ACTION, "bold": True, "mono": True, "space_after": 0,
           "href": PBIT_URL}], anchor=MSO_ANCHOR.MIDDLE, audit=False)
    footer(s, n, phase=11)
    notes(s, """
(8:50–9:15) "The machine learning app answers what is likely to happen. Power BI
answers what already happened and why — and it's what a non-technical
stakeholder actually opens."
Four pages: executive overview, customer insights, risk intelligence, retention
strategy. One sentence each, don't linger. Point out the underlined .pbit link —
it opens the reusable report template in the project repository.
""")

    # =========================================================== 16 PHASE 13
    n += 1
    s = add_slide(prs)
    eyebrow(s, "PHASE 13  ·  WEB APPLICATION")
    y = heading(s, "The product: a Signal Ops Console",
                "A working Flask application — bulk scoring, dashboard, private reports")
    _r = stack([(3.86, 0.16), (1.12, 0)]); r_img, r_strip = _r[0], _r[1]
    a1 = img_w(IMAGES / "06.prediction_dashboard.png", 1.0)
    a2 = img_w(IMAGES / "04.Bulk_prediction.png", 1.0)
    ph = (CW - 0.22) / (a1 + a2)
    d1 = a1 * ph
    image_fit(s, IMAGES / "06.prediction_dashboard.png", M, r_img[0], d1, ph)
    image_fit(s, IMAGES / "04.Bulk_prediction.png", M + d1 + 0.22, r_img[0],
              a2 * ph, ph)

    rect(s, M, r_strip[0], CW, r_strip[1], fill=PANEL, edge=LINE)
    text(s, M + 0.24, r_strip[0] + 0.14, CW - 2.72, r_strip[1] - 0.26,
         [{"t": "WHAT A REVIEWER DOES", "pt": 9.8, "color": OK, "bold": True,
           "mono": True, "space_after": 4},
          {"t": "Upload a CSV  →  validate the file before scoring  →  score the whole "
                "portfolio\nRead churn rate, risk split and revenue at risk  →  drill "
                "into a customer  →  download a private report",
           "pt": 10.4, "color": TEXT, "space_after": 0, "spacing": 1.16}])
    text(s, SW - M - 2.20, r_strip[0] + 0.14, 2.20, 0.34,
         [{"t": "REPORTS ARE PRIVATE — SCOPED BY SIGNED COOKIE", "pt": 8.4,
           "color": MUTED, "bold": True, "mono": True, "space_after": 0,
           "align": PP_ALIGN.RIGHT}], audit=False)
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
    _r = stack([(1.94, 0.16), (1.94, 0.16), (0.94, 0)])
    r1, r2, r_call = _r[0], _r[1], _r[2]
    for i, (t, b, fix, c) in enumerate(problems):
        x = M + (i % 2) * (ch + 0.26)
        yy, hh = (r1 if i < 2 else r2)
        rect(s, x, yy, ch, hh, fill=PANEL, edge=LINE)
        bar(s, x, yy + 0.16, 0.035, hh - 0.32, c)
        text(s, x + 0.24, yy + 0.18, ch - 0.50, hh - 0.36,
             [{"t": t, "pt": 12, "color": TEXT, "bold": True, "space_after": 4},
              {"t": b, "pt": 10.2, "color": MUTED, "space_after": 5, "spacing": 1.16},
              {"t": f"→ fixed by {fix}", "pt": 10.4, "color": c, "bold": True,
               "mono": True, "space_after": 0}])

    callout(s, M, r_call[0], CW, r_call[1], "SHIPPED AS",
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
    y = heading(s, "Tools behind the RetainIQ workflow",
                "A shared Python path keeps model training and live scoring in sync")

    logo_dir = ASSETS / "tool_logos"
    logo_panel_w, logo_gap = 3.10, 0.20
    logo_x, logo_y = M, CONTENT_TOP
    rect(s, logo_x, logo_y, logo_panel_w, CAV, fill=PANEL, edge=LINE)
    text(s, logo_x + 0.16, logo_y + 0.12, logo_panel_w - 0.32, 0.24,
         [{"t": "TOOLS USED", "pt": 9.8, "color": MUTED, "bold": True,
           "mono": True, "space_after": 0}], audit=False)

    logo_items = [
        ("python", "Python"), ("pandas", "Pandas"),
        ("scikitlearn", "Scikit-learn"), ("flask", "Flask"),
        ("fastapi", "FastAPI"), ("docker", "Docker"),
        ("powerbi", "Power BI"), ("github", "GitHub"),
    ]
    logo_pad, logo_col_gap = 0.16, 0.12
    tile_w = (logo_panel_w - 2 * logo_pad - logo_col_gap) / 2
    tile_h, tile_gap, grid_top = 1.02, 0.08, logo_y + 0.47
    for i, (icon, label) in enumerate(logo_items):
        col, row = i % 2, i // 2
        x = logo_x + logo_pad + col * (tile_w + logo_col_gap)
        yy = grid_top + row * (tile_h + tile_gap)
        rect(s, x, yy, tile_w, tile_h, fill=PANEL_2, edge=LINE_SOFT)
        icon_path = logo_dir / f"{icon}.png"
        image_fit(s, icon_path, x + (tile_w - 0.42) / 2, yy + 0.09, 0.42, 0.42)
        text(s, x + 0.05, yy + 0.58, tile_w - 0.10, 0.25,
             [{"t": label, "pt": 9.2, "color": TEXT, "bold": True,
               "space_after": 0, "align": PP_ALIGN.CENTER}], audit=False)

    right_x = logo_x + logo_panel_w + logo_gap
    right_w = SW - M - right_x
    groups = [
        ("LANGUAGE & DATA", OK,
         ["Python 3.11 · Pandas · NumPy · SQL Server · SQLAlchemy"]),
        ("MACHINE LEARNING", WARN,
         ["Scikit-learn · Logistic Regression (deployed) · Random Forest",
          "XGBoost · LightGBM · Joblib · SHAP (research only)"]),
        ("WEB APPLICATION", ACTION,
         ["Flask · Jinja2 · HTML5/CSS3 · Vanilla JavaScript · Gunicorn"]),
        ("API & DELIVERY", BAD,
         ["FastAPI + Pydantic · Docker · Render · Git/GitHub",
          "Power BI · .pbix report · .pbit template"]),
    ]
    row_gap = 0.14
    card_h = (CAV - 3 * row_gap) / 4
    for i, (title, color, details) in enumerate(groups):
        yy = CONTENT_TOP + i * (card_h + row_gap)
        rect(s, right_x, yy, right_w, card_h, fill=PANEL, edge=LINE)
        bar(s, right_x, yy, 0.045, card_h, color)
        blocks = [{"t": title, "pt": 10.2, "color": color, "bold": True,
                   "mono": True, "space_after": 5}]
        blocks += [{"t": line, "pt": 12.7, "color": TEXT,
                    "space_after": 0, "spacing": 1.18} for line in details]
        text(s, right_x + 0.22, yy + 0.14, right_w - 0.42, card_h - 0.26,
             blocks)

    footer(s, n, complete_through=len(PHASES))
    notes(s, """
(10:10–10:35) Use the logos as visual anchors; do not read the lists line by line.
"Python, pandas and NumPy handle the data. Scikit-learn trains the deployed
Logistic Regression; Flask and FastAPI expose it; Docker, Render and GitHub carry
it to production; Power BI presents the business view."
SHAP was explored only in research; the production explanations come from the
Logistic Regression coefficients. One shared Python package keeps training and
serving on the same preprocessing and scoring path.
""")

    # =========================================================== 19 RESPONSIBLE AI
    n += 1
    s = add_slide(prs)
    eyebrow(s, "RESPONSIBLE AI, LIMITS & THE NEXT CYCLE")
    y = heading(s, "What this model should not be trusted with",
                "Where it breaks, who it might treat unfairly, and what the next cycle fixes")
    col = (CW - 2 * 0.26) / 3

    rect(s, M, CONTENT_TOP, col, CAV, fill=PANEL, edge=WARN, edge_w=1.2)
    text(s, M + 0.22, CONTENT_TOP + 0.18, col - 0.44, CAV - 0.36,
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

    rect(s, M + col + 0.26, CONTENT_TOP, col, CAV, fill=PANEL, edge=BAD, edge_w=1.2)
    text(s, M + col + 0.48, CONTENT_TOP + 0.18, col - 0.44, CAV - 0.36,
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

    rect(s, M + 2 * (col + 0.26), CONTENT_TOP, col, CAV, fill=PANEL, edge=ACTION, edge_w=1.2)
    text(s, M + 2 * (col + 0.26) + 0.22, CONTENT_TOP + 0.18, col - 0.44, CAV - 0.36,
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
    _r = stack([(4.34, 0.20), (0.60, 0)]); r_pan, r_credit = _r[0], _r[1]
    rect(s, M, r_pan[0], 5.82, r_pan[1], fill=PANEL, edge=OK, edge_w=1.4)
    text(s, M + 0.28, r_pan[0] + 0.24, 5.26, r_pan[1] - 0.48,
         [{"t": "LIVE APPLICATION", "pt": 10.5, "color": OK, "bold": True, "mono": True,
           "space_after": 8},
          {"t": LIVE_URL_DISPLAY, "pt": 11, "color": ACTION, "bold": True,
           "space_after": 10, "spacing": 1.16, "href": LIVE_URL},
          {"t": "Bulk score a customer CSV, read the dashboard, and download a private "
                "report. The model, risk bands, explanations and recommendations run "
                "live — the same code paths that produced every figure in this deck.",
           "pt": 10.8, "color": MUTED, "space_after": 12, "spacing": 1.20},
          {"t": "SOURCE CODE", "pt": 10.5, "color": ACTION, "bold": True, "mono": True,
           "space_after": 5},
          {"t": REPO_DISPLAY, "pt": 10.6, "color": ACTION, "mono": True,
           "space_after": 12, "spacing": 1.16, "href": REPO_URL},
          {"t": "Notebooks 01–09, the Flask console, the FastAPI wrapper and the "
                "nine-test suite are all in the repository.", "pt": 10.4,
           "color": MUTED, "space_after": 10, "spacing": 1.18},
          {"t": "The host is a free tier — open it once a minute before the review so "
                "the first scan is instant.", "pt": 10.4, "color": WARN,
           "space_after": 0, "spacing": 1.18}])

    rect(s, 6.54, r_pan[0], 2.98, r_pan[1], fill=WHITE, edge=LINE)
    qs = min(2.86, r_pan[1] - 0.92)
    image_fit(s, ASSETS / "qr_live_app.png", 6.54 + (2.98 - qs) / 2, r_pan[0] + 0.22, qs, qs)
    text(s, 6.70, r_pan[0] + r_pan[1] - 0.68, 2.66, 0.52,
         [{"t": "SCAN FOR THE", "pt": 9.5, "color": INK, "bold": True, "mono": True,
           "space_after": 2, "align": PP_ALIGN.CENTER},
          {"t": "LIVE APPLICATION", "pt": 9.5, "color": INK, "bold": True, "mono": True,
           "space_after": 0, "align": PP_ALIGN.CENTER}])

    rect(s, 9.62, r_pan[0], 3.09, r_pan[1], fill=PANEL, edge=LINE)
    image_fit(s, ASSETS / "qr_repository.png", 10.54, r_pan[0] + 0.24, 1.25, 1.25)
    text(s, 9.82, r_pan[0] + 0.28, 0.66, 1.00,
         [{"t": "SOURCE", "pt": 8.5, "color": MUTED, "bold": True, "mono": True,
           "space_after": 3},
          {"t": "REPO", "pt": 8.5, "color": MUTED, "bold": True, "mono": True,
           "space_after": 0}], audit=False)
    text(s, 9.82, r_pan[0] + 1.62, 2.69, 2.40,
         [{"t": "Thank you", "pt": 19, "color": TEXT, "bold": True, "space_after": 5},
          {"t": "Questions welcome.", "pt": 11, "color": OK, "space_after": 10},
          {"t": "RetainIQ — from a churn probability to a prioritised, explained, costed "
                "retention action.", "pt": 10.4, "color": MUTED, "space_after": 0,
           "spacing": 1.18}])

    signal_meter(s, 11.05, r_pan[0] + r_pan[1] - 1.30, filled=5, scale=0.85,
                 label=None, color=OK)
    text(s, M, r_credit[0], CW, r_credit[1],
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

    # Slide 20 is the final content slide; the five appendix slides are omitted.
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
