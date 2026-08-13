"""
Read the ESTV DataMatrix from page 1 of a scanned PDF - fast and robust.

Strategy (why it's quick):
  1. Only page 1 is touched, and only the top-right corner is rasterised.
     Decoding a full A4 page at 400 dpi is what makes naive solutions slow;
     cropping cuts the pixel count by ~10x before the decoder ever sees it.
  2. zxing-cpp (C++) is tried first - typically 2-10 ms on such a crop.
     pylibdmtx (libdmtx) is only used as a fallback: it is much slower but
     sometimes reads badly printed / skewed codes that zxing rejects.
  3. Escalation ladder: low dpi -> higher dpi -> binarised -> full page.
     The common case exits after the very first attempt.

Install:
    pip install pymupdf zxing-cpp numpy
    pip install pylibdmtx opencv-python-headless    # optional fallbacks
"""

from __future__ import annotations

import time
from dataclasses import dataclass
from typing import Optional

import fitz  # PyMuPDF
import numpy as np

try:
    import zxingcpp
except ImportError:
    zxingcpp = None

try:
    from pylibdmtx.pylibdmtx import decode as dmtx_decode
except ImportError:
    dmtx_decode = None

try:
    import cv2
except ImportError:
    cv2 = None


# Region of interest on page 1, as fractions of page width/height.
# "far right, close to the top" -> right 45 %, top 30 %. Widen if a scan
# is shifted; a slightly too large ROI costs almost nothing.
ROI = (0.55, 0.00, 1.00, 0.30)  # x0, y0, x1, y1


@dataclass
class DmtxResult:
    text: str
    dpi: int
    stage: str
    seconds: float


def _render(page: fitz.Page, dpi: int, roi: Optional[tuple] = None) -> np.ndarray:
    """Rasterise (part of) a page to a grayscale numpy array."""
    clip = None
    if roi:
        r = page.rect
        clip = fitz.Rect(
            r.x0 + roi[0] * r.width,
            r.y0 + roi[1] * r.height,
            r.x0 + roi[2] * r.width,
            r.y0 + roi[3] * r.height,
        )
    pix = page.get_pixmap(dpi=dpi, colorspace=fitz.csGRAY, clip=clip, annots=False)
    return np.frombuffer(pix.samples, dtype=np.uint8).reshape(pix.height, pix.width)


def _zxing(img: np.ndarray) -> Optional[str]:
    if zxingcpp is None:
        return None
    results = zxingcpp.read_barcodes(
        img,
        formats=zxingcpp.BarcodeFormat.DataMatrix,
        try_rotate=True,     # scans are often 90/180 degrees off
        try_downscale=True,  # helps when the code is rendered large
    )
    return results[0].text if results else None


def _libdmtx(img: np.ndarray, timeout_ms: int = 3000) -> Optional[str]:
    if dmtx_decode is None:
        return None
    h, w = img.shape
    res = dmtx_decode(
        (img.tobytes(), w, h),
        max_count=1,          # stop after the first hit - big speedup
        timeout=timeout_ms,   # never let libdmtx run away on noise
        shrink=1,
    )
    return res[0].data.decode("utf-8", "replace") if res else None


def _binarise(img: np.ndarray) -> np.ndarray:
    """Otsu threshold - rescues low-contrast or grey-ish scans."""
    if cv2 is None:
        thr = int(img.mean())
        return np.where(img > thr, 255, 0).astype(np.uint8)
    img = cv2.medianBlur(img, 3)
    _, out = cv2.threshold(img, 0, 255, cv2.THRESH_BINARY + cv2.THRESH_OTSU)
    return out


def read_estv_datamatrix(pdf_path: str, roi: tuple = ROI) -> Optional[DmtxResult]:
    t0 = time.perf_counter()
    with fitz.open(pdf_path) as doc:
        page = doc[0]

        # (dpi, use ROI, binarise)  - cheapest first
        plan = [
            (300, True,  False),
            (400, True,  False),
            (400, True,  True),
            (600, True,  True),
            (300, False, False),   # rotated / mis-scanned page: search all of it
            (400, False, True),
        ]

        for dpi, use_roi, binar in plan:
            img = _render(page, dpi, roi if use_roi else None)
            if binar:
                img = _binarise(img)

            stage = f"{'roi' if use_roi else 'full'}{'+otsu' if binar else ''}"

            text = _zxing(img)
            if text:
                return DmtxResult(text, dpi, f"zxing/{stage}", time.perf_counter() - t0)

            # libdmtx only on the small crops - it is too slow on full pages
            if use_roi:
                text = _libdmtx(img)
                if text:
                    return DmtxResult(text, dpi, f"dmtx/{stage}", time.perf_counter() - t0)

    return None


# --- optional: highest-fidelity variant -------------------------------------
# If the scan is a single embedded image per page, you can skip rasterising
# altogether and work on the original pixels (no resampling loss at all):
def read_from_native_image(pdf_path: str, roi: tuple = ROI) -> Optional[str]:
    with fitz.open(pdf_path) as doc:
        page = doc[0]
        images = page.get_images(full=True)
        if len(images) != 1:
            return None
        pix = fitz.Pixmap(doc, images[0][0])
        if pix.n - pix.alpha > 1:
            pix = fitz.Pixmap(fitz.csGRAY, pix)
        img = np.frombuffer(pix.samples, np.uint8).reshape(pix.height, pix.width)
        h, w = img.shape
        crop = img[int(roi[1] * h):int(roi[3] * h), int(roi[0] * w):int(roi[2] * w)]
        return _zxing(crop) or _libdmtx(crop)


if __name__ == "__main__":
    import sys

    for path in sys.argv[1:]:
        r = read_estv_datamatrix(path)
        if r:
            print(f"{path}\t{r.stage}@{r.dpi}dpi\t{r.seconds * 1000:.0f} ms\t{r.text}")
        else:
            print(f"{path}\tNO DATAMATRIX FOUND")
