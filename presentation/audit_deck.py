"""
Geometry audit for the generated deck.

There is no renderer in this environment, so this script acts as the QA pass:
it flags shapes that fall outside the slide, and text blocks whose estimated
rendered height collides with another text block. Run after build_deck.py.
"""
from __future__ import annotations

import sys
from collections import Counter
from pathlib import Path

from PIL import ImageFont
from pptx import Presentation
from pptx.util import Emu

HERE = Path(__file__).resolve().parent
DECK = HERE / "RetainIQ_Project_Review.pptx"

_TTF = Path(__import__("matplotlib").__file__).parent / "mpl-data" / "fonts" / "ttf"
FONTS = {
    (False, False): ImageFont.truetype(str(_TTF / "DejaVuSans.ttf"), 100),
    (True, False): ImageFont.truetype(str(_TTF / "DejaVuSans-Bold.ttf"), 100),
    (False, True): ImageFont.truetype(str(_TTF / "DejaVuSansMono.ttf"), 100),
    (True, True): ImageFont.truetype(str(_TTF / "DejaVuSansMono-Bold.ttf"), 100),
}

SW, SH = 13.333, 7.5
TOL = 0.02


def width_pt(text, pt, bold, mono):
    return FONTS[(bold, mono)].getlength(text) / 100 * pt


def block_metrics(tf, box_w_in, box_h_in):
    """Estimated rendered height (inches) and max line width for a text frame."""
    total_h = 0.0
    widest = 0.0
    limit = box_w_in * 72
    for para in tf.paragraphs:
        text = "".join(r.text for r in para.runs)
        if not text.strip():
            continue
        pt = max((r.font.size.pt for r in para.runs if r.font.size), default=12)
        bold = any(r.font.bold for r in para.runs)
        mono = any((r.font.name or "").lower().startswith("consolas") for r in para.runs)
        words, lines, cur = text.split(), 0, ""
        for word in words:
            trial = f"{cur} {word}".strip()
            if width_pt(trial, pt, bold, mono) <= limit or not cur:
                cur = trial
            else:
                widest = max(widest, width_pt(cur, pt, bold, mono))
                lines += 1
                cur = word
        widest = max(widest, width_pt(cur, pt, bold, mono))
        lines += 1
        spacing = para.line_spacing or 1.0
        if isinstance(spacing, float):
            total_h += lines * pt * 1.22 * spacing / 72
        else:
            total_h += lines * pt * 1.22 / 72
        if para.space_after is not None:
            total_h += para.space_after.pt / 72
        if para.space_before is not None:
            total_h += para.space_before.pt / 72
    return min(total_h, box_h_in * 3), widest


def main():
    prs = Presentation(str(DECK))
    problems = []
    expected_slides = 22
    if len(prs.slides) != expected_slides:
        problems.append(f"expected {expected_slides} slides, found {len(prs.slides)}")
    else:
        for idx, slide in enumerate(prs.slides, 1):
            targets = []
            named_nav = 0
            for shape in slide.shapes:
                target = shape.click_action.target_slide
                if target is not None:
                    targets.append(target.slide_id)
                nv_sp = getattr(shape._element, "nvSpPr", None)
                c_nv_pr = getattr(nv_sp, "cNvPr", None) if nv_sp is not None else None
                name = c_nv_pr.get("name") if c_nv_pr is not None else None
                if name and name.startswith("Slide navigation "):
                    named_nav += 1
            expected_targets = Counter(s.slide_id for s in prs.slides)
            expected_targets.update(expected_targets)
            if Counter(targets) != expected_targets or named_nav != expected_slides:
                problems.append(
                    f"slide {idx}: expected {expected_slides} clickable nav dashes "
                    f"(two links per destination), found {named_nav} named targets "
                    f"and {len(targets)} internal links")

    for idx, slide in enumerate(prs.slides, 1):
        texts = []
        for shape in slide.shapes:
            x, y = shape.left / 914400, shape.top / 914400
            w, h = shape.width / 914400, shape.height / 914400

            if x < -TOL or y < -TOL or x + w > SW + TOL or y + h > SH + TOL:
                problems.append(
                    f"slide {idx}: shape outside slide at ({x:.2f},{y:.2f}) "
                    f"size {w:.2f}x{h:.2f}")

            if shape.has_text_frame and shape.text_frame.text.strip():
                est_h, widest = block_metrics(shape.text_frame, w, h)
                if widest > w * 72 + 1:
                    problems.append(
                        f"slide {idx}: unbreakable text wider than box "
                        f"({widest / 72:.2f}in > {w:.2f}in) -> "
                        f"\"{shape.text_frame.text[:50]}\"")
                # text taller than its own box spills visually (no autofit is set)
                if est_h > h + 0.08:
                    problems.append(
                        f"slide {idx}: text taller than its box "
                        f"({est_h:.2f}in > {h:.2f}in) -> "
                        f"\"{shape.text_frame.text[:46]}\"")
                texts.append((shape, x, y, w, h, est_h))

        for i in range(len(texts)):
            for j in range(i + 1, len(texts)):
                a, ax, ay, aw, ah, ahh = texts[i]
                b, bx, by, bw, bh, bhh = texts[j]
                ax2, ay2 = ax + aw, ay + min(ahh, ah)
                bx2, by2 = bx + bw, by + min(bhh, bh)
                ox = min(ax2, bx2) - max(ax, bx)
                oy = min(ay2, by2) - max(ay, by)
                if ox > 0.06 and oy > 0.06:
                    problems.append(
                        f"slide {idx}: text overlap {ox:.2f}x{oy:.2f}in -> "
                        f"\"{a.text_frame.text[:38]}\" vs \"{b.text_frame.text[:38]}\"")

    print(f"audited {len(prs.slides)} slides")
    if problems:
        print(f"\n{len(problems)} issue(s):")
        for p in problems:
            print("  -", p)
    else:
        print("clean: no out-of-bounds shapes, no text collisions")
        print("verified: all 22 footer dashes link to all 22 slides on every slide")
    return 1 if problems else 0


if __name__ == "__main__":
    sys.exit(main())
