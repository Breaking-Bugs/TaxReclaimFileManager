r"""
Batch processing of ESTV tax-voucher PDFs.

For every PDF in the input folder:
  1. decode the DataMatrix on page 1 (ROI top-right, zxing-cpp -> libdmtx ladder)
  2. read field 7 = Rta Claim and field 2 = ESTV ID
  3. look the Rta Claim up in export.csv and verify the ESTV ID
  4. on a unique match  -> rename to <Ident>_2.pdf, move to the output folder,
     append Ident to idents.txt
     otherwise          -> move to the manual-review folder, record the reason

Artifacts written to the logs folder:
    processing_log.csv   one row per PDF (appended across runs)
    exceptions.csv       one row per failure (appended across runs)
    idents.txt           successfully matched Idents, one per line (deduplicated)
    run_summary.txt      human-readable summary of this run

Usage (Windows):
    python batch_estv_reclaims.py --input .\input --export .\export.csv ^
        --output .\output --review .\manual_review --logs .\logs --dry-run

Install:
    pip install pymupdf zxing-cpp numpy
    pip install pylibdmtx opencv-python-headless      # optional fallbacks
"""

from __future__ import annotations

import argparse
import csv
import os
import re
import shutil
import sys
import time
from collections import Counter, defaultdict
from concurrent.futures import ProcessPoolExecutor
from dataclasses import dataclass, field
from datetime import datetime
from pathlib import Path
from typing import Iterable, Optional

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


# --------------------------------------------------------------------------
# DataMatrix decoding (page 1, top-right corner)
# --------------------------------------------------------------------------

ROI = (0.55, 0.00, 1.00, 0.30)  # x0, y0, x1, y1 as fractions of the page

# (dpi, use ROI, binarise) - cheapest first, the normal case exits on step 1
DECODE_PLAN = [
    (300, True, False),
    (400, True, False),
    (400, True, True),
    (600, True, True),
    (300, False, False),   # rotated / badly scanned page: search the whole page
    (400, False, True),
]


def _render(page: fitz.Page, dpi: int, roi: Optional[tuple]) -> np.ndarray:
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
    res = zxingcpp.read_barcodes(
        img,
        formats=zxingcpp.BarcodeFormat.DataMatrix,
        try_rotate=True,
        try_downscale=True,
    )
    return res[0].text if res else None


def _libdmtx(img: np.ndarray, timeout_ms: int = 3000) -> Optional[str]:
    if dmtx_decode is None:
        return None
    h, w = img.shape
    res = dmtx_decode((img.tobytes(), w, h), max_count=1, timeout=timeout_ms, shrink=1)
    return res[0].data.decode("utf-8", "replace") if res else None


def _binarise(img: np.ndarray) -> np.ndarray:
    if cv2 is None:
        return np.where(img > int(img.mean()), 255, 0).astype(np.uint8)
    return cv2.threshold(cv2.medianBlur(img, 3), 0, 255,
                         cv2.THRESH_BINARY + cv2.THRESH_OTSU)[1]


@dataclass
class Decoded:
    path: str
    raw: Optional[str] = None
    stage: str = ""
    dpi: int = 0
    ms: float = 0.0
    error: str = ""


def decode_pdf(path: str, roi: tuple = ROI) -> Decoded:
    """Decode the DataMatrix on page 1. Module-level so it is picklable."""
    t0 = time.perf_counter()
    try:
        with fitz.open(path) as doc:
            if doc.page_count == 0:
                return Decoded(path, error="empty_pdf", ms=(time.perf_counter() - t0) * 1000)
            page = doc[0]
            for dpi, use_roi, binar in DECODE_PLAN:
                img = _render(page, dpi, roi if use_roi else None)
                if binar:
                    img = _binarise(img)
                stage = ("roi" if use_roi else "full") + ("+otsu" if binar else "")

                text = _zxing(img)
                if text:
                    return Decoded(path, text, f"zxing/{stage}", dpi,
                                   (time.perf_counter() - t0) * 1000)
                if use_roi:  # libdmtx only on the small crop - too slow on full pages
                    text = _libdmtx(img)
                    if text:
                        return Decoded(path, text, f"dmtx/{stage}", dpi,
                                       (time.perf_counter() - t0) * 1000)
    except Exception as exc:  # corrupt/encrypted file etc.
        return Decoded(path, error=f"{type(exc).__name__}: {exc}",
                       ms=(time.perf_counter() - t0) * 1000)
    return Decoded(path, ms=(time.perf_counter() - t0) * 1000)


# --------------------------------------------------------------------------
# DataMatrix payload -> fields
# --------------------------------------------------------------------------

def parse_fields(raw: str) -> dict[str, str]:
    """'2=052.0308.7797;4=;7=DE-4703-4568-00;' -> {'2': '052.0308.7797', ...}"""
    out: dict[str, str] = {}
    for chunk in re.split(r"[;\r\n\x1d]+", raw):
        chunk = chunk.strip()
        if not chunk or "=" not in chunk:
            continue
        key, _, value = chunk.partition("=")
        out[key.strip()] = value.strip()
    return out


def canon(value: str) -> str:
    """Comparison key: uppercase, punctuation/whitespace removed.

    Makes '052.0308.7797' == '052 0308 7797' and 'DE-4703-4568-00' ==
    'DE4703456800'. Raw values are always kept in the logs so nothing is lost.
    """
    return re.sub(r"[^0-9A-Z]", "", (value or "").upper())


# --------------------------------------------------------------------------
# export.csv
# --------------------------------------------------------------------------

@dataclass
class ExportRow:
    ident: str
    rta_claim: str
    estv_id: str
    line: int


REQUIRED_COLUMNS = {"ident": "Ident", "rtaclaim": "Rta Claim", "estvid": "ESTV ID"}


def _norm_header(name: str) -> str:
    return re.sub(r"[^a-z0-9]", "", (name or "").lower())


def load_export(path: Path) -> tuple[dict[str, list[ExportRow]], list[ExportRow]]:
    """Return (index by canonical Rta Claim, all rows)."""
    data = path.read_bytes()
    for encoding in ("utf-8-sig", "cp1252", "latin-1"):
        try:
            text = data.decode(encoding)
            break
        except UnicodeDecodeError:
            continue
    else:
        raise SystemExit(f"Cannot decode {path}")

    sample = text[:4096]
    try:
        dialect = csv.Sniffer().sniff(sample, delimiters=";,\t")
    except csv.Error:
        dialect = csv.excel
        dialect.delimiter = ";" if sample.count(";") >= sample.count(",") else ","

    reader = csv.DictReader(text.splitlines(), dialect=dialect)
    headers = {_norm_header(h): h for h in (reader.fieldnames or [])}
    missing = [label for key, label in REQUIRED_COLUMNS.items() if key not in headers]
    if missing:
        raise SystemExit(
            f"export.csv is missing column(s): {', '.join(missing)}. "
            f"Found: {reader.fieldnames}"
        )

    rows: list[ExportRow] = []
    index: dict[str, list[ExportRow]] = defaultdict(list)
    for line_no, rec in enumerate(reader, start=2):
        row = ExportRow(
            ident=(rec.get(headers["ident"]) or "").strip(),
            rta_claim=(rec.get(headers["rtaclaim"]) or "").strip(),
            estv_id=(rec.get(headers["estvid"]) or "").strip(),
            line=line_no,
        )
        if not row.ident and not row.rta_claim:
            continue
        rows.append(row)
        index[canon(row.rta_claim)].append(row)
    return index, rows


# --------------------------------------------------------------------------
# Matching
# --------------------------------------------------------------------------

@dataclass
class Outcome:
    pdf: Path
    status: str = "OK"           # OK | REVIEW
    reason: str = ""             # exception code on failure
    detail: str = ""
    ident: str = ""
    rta_claim: str = ""
    estv_id: str = ""
    stage: str = ""
    dpi: int = 0
    ms: float = 0.0
    target: str = ""


def match(dec: Decoded, index: dict[str, list[ExportRow]]) -> Outcome:
    out = Outcome(pdf=Path(dec.path), stage=dec.stage, dpi=dec.dpi, ms=dec.ms)

    if dec.error:
        return _fail(out, "PDF_ERROR", dec.error)
    if not dec.raw:
        return _fail(out, "NO_DATAMATRIX", "no DataMatrix decoded on page 1")

    fields = parse_fields(dec.raw)
    out.rta_claim = fields.get("7", "")
    out.estv_id = fields.get("2", "")

    if not out.rta_claim:
        return _fail(out, "MISSING_FIELD_7", f"payload: {dec.raw[:120]}")
    if not out.estv_id:
        return _fail(out, "MISSING_FIELD_2", f"payload: {dec.raw[:120]}")

    candidates = index.get(canon(out.rta_claim), [])
    if not candidates:
        return _fail(out, "NO_MATCH_RTA", f"Rta Claim {out.rta_claim} not in export.csv")

    confirmed = [r for r in candidates if canon(r.estv_id) == canon(out.estv_id)]
    if not confirmed:
        found = ", ".join(sorted({r.estv_id for r in candidates})) or "(empty)"
        return _fail(out, "ESTV_MISMATCH",
                     f"DataMatrix ESTV ID {out.estv_id} vs export.csv {found}")

    idents = sorted({r.ident for r in confirmed if r.ident})
    if len(idents) != 1:
        return _fail(out, "AMBIGUOUS_MATCH",
                     f"{len(idents)} Idents for this Rta Claim: {', '.join(idents) or '(none)'}")

    out.ident = idents[0]
    return out


def _fail(out: Outcome, reason: str, detail: str) -> Outcome:
    out.status, out.reason, out.detail = "REVIEW", reason, detail
    return out


# --------------------------------------------------------------------------
# Run
# --------------------------------------------------------------------------

def iter_pdfs(folder: Path, recursive: bool) -> list[Path]:
    it: Iterable[Path] = folder.rglob("*") if recursive else folder.glob("*")
    return sorted(p for p in it if p.is_file() and p.suffix.lower() == ".pdf")


def _open_csv(path: Path, header: list[str]):
    """Append mode; write the header only when the file is new."""
    new = not path.exists() or path.stat().st_size == 0
    fh = path.open("a", newline="", encoding="utf-8-sig")
    writer = csv.writer(fh, delimiter=";")
    if new:
        writer.writerow(header)
    return fh, writer


def main(argv: Optional[list[str]] = None) -> int:
    ap = argparse.ArgumentParser(description="Batch-match ESTV DataMatrix PDFs against export.csv")
    ap.add_argument("--input", required=True, type=Path)
    ap.add_argument("--export", required=True, type=Path, help="export.csv")
    ap.add_argument("--output", required=True, type=Path)
    ap.add_argument("--review", required=True, type=Path, help="manual review folder")
    ap.add_argument("--logs", type=Path, default=Path("./logs"))
    ap.add_argument("--recursive", action="store_true")
    ap.add_argument("--copy", action="store_true", help="copy instead of move (keeps originals)")
    ap.add_argument("--dry-run", action="store_true", help="decide everything, touch nothing")
    ap.add_argument("--workers", type=int, default=1, help="parallel decode processes")
    args = ap.parse_args(argv)

    if zxingcpp is None and dmtx_decode is None:
        return _die("Neither zxing-cpp nor pylibdmtx is installed - nothing can be decoded.")
    if not args.input.is_dir():
        return _die(f"Input folder not found: {args.input}")
    if not args.export.is_file():
        return _die(f"export.csv not found: {args.export}")

    index, export_rows = load_export(args.export)
    pdfs = iter_pdfs(args.input, args.recursive)
    if not pdfs:
        return _die(f"No PDFs found in {args.input}")

    args.logs.mkdir(parents=True, exist_ok=True)
    if not args.dry_run:
        args.output.mkdir(parents=True, exist_ok=True)
        args.review.mkdir(parents=True, exist_ok=True)

    prefix = "dryrun_" if args.dry_run else ""
    log_path = args.logs / f"{prefix}processing_log.csv"
    exc_path = args.logs / f"{prefix}exceptions.csv"
    ident_path = args.logs / f"{prefix}idents.txt"
    summary_path = args.logs / f"{prefix}run_summary.txt"

    run_id = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    t_start = time.perf_counter()

    # --- decode (optionally in parallel; libdmtx is not thread-safe, so processes)
    if args.workers > 1:
        with ProcessPoolExecutor(max_workers=args.workers) as pool:
            decoded = list(pool.map(decode_pdf, [str(p) for p in pdfs], chunksize=4))
    else:
        decoded = [decode_pdf(str(p)) for p in pdfs]

    # --- match, then move sequentially
    existing_idents = set()
    if ident_path.exists():
        existing_idents = {l.strip() for l in ident_path.read_text(encoding="utf-8").splitlines() if l.strip()}

    outcomes: list[Outcome] = []
    new_idents: list[str] = []
    used_targets: set[Path] = set()

    for dec in decoded:
        out = match(dec, index)

        if out.status == "OK":
            target = args.output / f"{out.ident}_2.pdf"
            if target.exists() or target in used_targets:
                _fail(out, "TARGET_EXISTS", f"{target.name} already exists in the output folder")
            else:
                out.target = str(target)
                used_targets.add(target)
                if not args.dry_run:
                    try:
                        (shutil.copy2 if args.copy else shutil.move)(str(out.pdf), str(target))
                    except Exception as exc:
                        _fail(out, "MOVE_FAILED", f"{type(exc).__name__}: {exc}")
                        out.target = ""
                if out.status == "OK" and out.ident not in existing_idents:
                    existing_idents.add(out.ident)
                    new_idents.append(out.ident)

        if out.status != "OK":
            target = _unique(args.review / out.pdf.name)
            out.target = str(target)
            if not args.dry_run:
                try:
                    (shutil.copy2 if args.copy else shutil.move)(str(out.pdf), str(target))
                except Exception as exc:
                    out.detail += f" | move to review failed: {exc}"
                    out.target = ""

        outcomes.append(out)

    # --- artifacts
    log_fh, log_w = _open_csv(log_path, [
        "run", "file", "status", "reason", "detail", "ident",
        "rta_claim", "estv_id", "decode_stage", "dpi", "decode_ms", "target",
    ])
    exc_fh, exc_w = _open_csv(exc_path, [
        "run", "file", "reason", "detail", "rta_claim", "estv_id", "moved_to",
    ])
    with log_fh, exc_fh:
        for o in outcomes:
            log_w.writerow([run_id, o.pdf.name, o.status, o.reason, o.detail, o.ident,
                            o.rta_claim, o.estv_id, o.stage, o.dpi or "",
                            f"{o.ms:.0f}", o.target])
            if o.status != "OK":
                exc_w.writerow([run_id, o.pdf.name, o.reason, o.detail,
                                o.rta_claim, o.estv_id, o.target])

    if new_idents:
        with ident_path.open("a", encoding="utf-8") as fh:
            for ident in new_idents:
                fh.write(ident + "\n")

    elapsed = time.perf_counter() - t_start
    ok = sum(1 for o in outcomes if o.status == "OK")
    reasons = Counter(o.reason for o in outcomes if o.status != "OK")
    avg_ms = sum(o.ms for o in outcomes) / len(outcomes) if outcomes else 0

    summary = [
        f"Run           : {run_id}{'  [DRY RUN - no files touched]' if args.dry_run else ''}",
        f"Input folder  : {args.input.resolve()}",
        f"export.csv    : {args.export.resolve()}  ({len(export_rows)} rows)",
        f"Output folder : {args.output.resolve()}",
        f"Review folder : {args.review.resolve()}",
        "",
        f"PDFs processed: {len(outcomes)}",
        f"Matched       : {ok}",
        f"Manual review : {len(outcomes) - ok}",
        f"New Idents    : {len(new_idents)}",
        "",
        f"Duration      : {elapsed:.1f} s   (avg decode {avg_ms:.0f} ms/PDF, workers={args.workers})",
    ]
    if reasons:
        summary += ["", "Exceptions by reason:"]
        summary += [f"  {r:<16} {n}" for r, n in reasons.most_common()]
    text = "\n".join(summary) + "\n"

    with summary_path.open("a", encoding="utf-8") as fh:
        fh.write(text + "-" * 60 + "\n")
    print(text)
    return 0 if ok == len(outcomes) else 1


def _unique(path: Path) -> Path:
    """Never overwrite in the review folder."""
    if not path.exists():
        return path
    stem, suffix, n = path.stem, path.suffix, 1
    while True:
        candidate = path.with_name(f"{stem}__{n}{suffix}")
        if not candidate.exists():
            return candidate
        n += 1


def _die(msg: str) -> int:
    print(f"ERROR: {msg}", file=sys.stderr)
    return 2


if __name__ == "__main__":
    sys.exit(main())
