"""
Generate the QR codes used by the closing slide.

Each code is rendered with the project's ink colour and then decoded again with
OpenCV, so the deck can only ever ship a QR that actually resolves to the URL
printed beside it. Run from the repository root:

    python presentation/build_qr.py
"""
from __future__ import annotations

from pathlib import Path

import cv2
import qrcode
from qrcode.constants import ERROR_CORRECT_H

HERE = Path(__file__).resolve().parent
ASSETS = HERE / "assets"
ASSETS.mkdir(parents=True, exist_ok=True)

URL_FILE = HERE / "live_url.txt"
LIVE_URL = (URL_FILE.read_text().strip() if URL_FILE.exists()
            else "https://retainiq-predictive-customer-retention-zq6x.onrender.com")
REPO_URL = "https://github.com/Logesh0247/RetainIQ_Signal_Ops_Console"

INK = (0x0A, 0x0F, 0x1C)
WHITE = (0xFF, 0xFF, 0xFF)


def make(url: str, name: str) -> Path:
    qr = qrcode.QRCode(
        version=None,
        error_correction=ERROR_CORRECT_H,   # 30% damage tolerance — survives projection
        box_size=16,
        border=4,                           # quiet zone required by the spec
    )
    qr.add_data(url)
    qr.make(fit=True)
    img = qr.make_image(fill_color=INK, back_color=WHITE).convert("RGB")
    path = ASSETS / name
    img.save(path)
    return path


def verify(path: Path, expected: str) -> bool:
    detector = cv2.QRCodeDetector()
    image = cv2.imread(str(path))
    decoded, _, _ = detector.detectAndDecode(image)
    ok = decoded.strip() == expected.strip()
    print(f"{'OK  ' if ok else 'FAIL'} {path.name}")
    print(f"     encoded : {expected}")
    print(f"     decoded : {decoded.strip() or '(nothing decoded)'}")
    return ok


def main() -> int:
    print("generating QR codes ...\n")
    live = make(LIVE_URL, "qr_live_app.png")
    repo = make(REPO_URL, "qr_repository.png")

    print("verifying by decoding the rendered PNGs ...\n")
    ok_live = verify(live, LIVE_URL)
    ok_repo = verify(repo, REPO_URL)

    if ok_live and ok_repo:
        print("\nboth QR codes decode to the intended URLs")
        return 0
    print("\nWARNING: a QR code did not decode correctly")
    return 1


if __name__ == "__main__":
    raise SystemExit(main())
