"""
Layout quality audit — measures the things a geometry check misses.

For every slide it reports:
  * dead space: the gap between the lowest content and the footer rule
  * image utilisation: how much of its allotted box each picture actually fills
  * density: how much of the content area carries content

These are the problems that only show up when you look at a rendered slide.

    python presentation/audit_layout.py
"""
from __future__ import annotations

from pathlib import Path

from pptx import Presentation

HERE = Path(__file__).resolve().parent
DECK = HERE / "RetainIQ_Project_Review.pptx"

CONTENT_TOP = 1.66          # below the header rule
FOOTER_RULE = 7.06
RAIL = 6.90
M = 0.62
SW, SH = 13.333, 7.5


def inches(v):
    return v / 914400


def main():
    prs = Presentation(str(DECK))
    print(f"{'slide':>5} {'content bottom':>15} {'dead space':>11} "
          f"{'images (placed / box)':>26}  notes")
    print("-" * 100)

    worst = []
    for i, slide in enumerate(prs.slides, 1):
        bottom = 0.0
        pics = []
        for sh in slide.shapes:
            x, y = inches(sh.left), inches(sh.top)
            w, h = inches(sh.width), inches(sh.height)
            # ignore the background, footer furniture and rail ticks
            if w > SW - 0.1 and h > SH - 0.1:
                continue
            if abs(y - RAIL) < 0.06 or y > FOOTER_RULE - 0.02:
                continue
            bottom = max(bottom, y + h)
            if sh.shape_type == 13:
                try:
                    img_w, img_h = sh.image.size
                except Exception:
                    continue
                pics.append((w, h, img_w / img_h))

        dead = FOOTER_RULE - bottom
        notes = []
        if dead > 0.75:
            notes.append(f"DEAD SPACE {dead:.2f}in")

        pic_note = []
        for w, h, aspect in pics:
            if h * aspect < w:          # height-limited
                fill = (h * aspect) / w
                pic_note.append(f"{h * aspect:.1f}x{h:.1f}/{w:.1f}x{h:.1f} ({fill*100:.0f}%)")
                if fill < 0.75:
                    notes.append(f"image only {fill*100:.0f}% of width")
            else:
                fill = (w / aspect) / h
                pic_note.append(f"{w:.1f}x{w/aspect:.1f}/{w:.1f}x{h:.1f} ({fill*100:.0f}%)")

        print(f"{i:>5} {bottom:>15.2f} {dead:>11.2f} {('; '.join(pic_note)):>26}  {' | '.join(notes)}")
        if notes:
            worst.append((i, notes))

    print()
    print(f"{len(worst)} slide(s) with layout notes:")
    for i, notes in worst:
        print(f"  slide {i}: {'; '.join(notes)}")


if __name__ == "__main__":
    main()
