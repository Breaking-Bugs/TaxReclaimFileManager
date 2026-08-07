#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
process_tax_certificates.py
===========================

Ordnet deutsche Steuerbescheinigungen (PDF) anhand von XLSX-Metadaten und einer
zentralen Transactions-CSV den passenden Reclaim References zu, kopiert sie in
eine Zielstruktur und benennt sie nach dem Schema ``YYYY-MM-DD_ISIN.pdf`` um.

Kernprinzipien
--------------
* Quelldateien werden **niemals** verändert, verschoben oder gelöscht.
* Bestehende Zieldateien werden **niemals** überschrieben.
* Einzelne fehlerhafte Zeilen brechen den Lauf nicht ab, sondern werden im
  Audit-Log ``processing_log.csv`` protokolliert.
* Die Verarbeitung ist deterministisch (Ordner -> XLSX -> Excel-Zeile).

Aufruf (Windows)
----------------
    python process_tax_certificates.py --input "C:\\Daten\\Input" ^
        --csv "C:\\Daten\\Transactions_2024.csv" --output "C:\\Daten\\Output"
    python process_tax_certificates.py --dry-run
"""

from __future__ import annotations

import argparse
import csv
import hashlib
import logging
import re
import shutil
import sys
import unicodedata
from dataclasses import dataclass, field
from datetime import date, datetime, timedelta
from pathlib import Path
from typing import Dict, List, Optional, Sequence, Set, Tuple

# --------------------------------------------------------------------------- #
# Optionale Abhaengigkeiten                                                    #
# --------------------------------------------------------------------------- #

try:  # pandas wird fuer das CSV-Einlesen bevorzugt, ist aber nicht zwingend
    import pandas as pd  # type: ignore
except Exception:  # pragma: no cover - Fallback auf das csv-Modul
    pd = None  # type: ignore

try:
    import openpyxl  # type: ignore
except Exception:  # pragma: no cover
    openpyxl = None  # type: ignore


LOGGER = logging.getLogger("tax_certificates")

# --------------------------------------------------------------------------- #
# Konstanten                                                                   #
# --------------------------------------------------------------------------- #

# Positionsbasierte XLSX-Spalten (0-basiert)
COL_ISIN = 1        # Spalte B  - Financial Instrument
COL_VALUE_DATE = 2  # Spalte C  - Paymt. Date
COL_BO_NAME = 8     # Spalte I  - BO Name
COL_REQUEST_ID = 13  # Spalte N  - Clearstream Request ID
MIN_REQUIRED_COLUMNS = COL_REQUEST_ID + 1

QUERY_EXECUTED_PREFIX = "query executed:"

CSV_COL_RECLAIM = "Reclaim Reference"
CSV_COL_BO_NAME = "Beneficial Owner Name"
CSV_COL_ISIN = "ISIN"
CSV_COL_VALUE_DATE = "Value Date"
REQUIRED_CSV_COLUMNS = (
    CSV_COL_RECLAIM,
    CSV_COL_BO_NAME,
    CSV_COL_ISIN,
    CSV_COL_VALUE_DATE,
)

CSV_ENCODINGS = ("utf-8-sig", "utf-8", "cp1252", "latin-1")

LOG_FILENAME = "processing_log.csv"
LOG_COLUMNS = (
    "Timestamp",
    "Source Folder",
    "XLSX File",
    "Excel Row",
    "Original PDF Name",
    "Source PDF Path",
    "XLSX BO Name",
    "XLSX ISIN",
    "XLSX Value Date",
    "Matched CSV BO Name",
    "Reclaim Reference",
    "Match Method",
    "Match Token Count",
    "Candidate Count After ISIN And Date",
    "Candidate Reclaim References",
    "Target PDF Name",
    "Target PDF Path",
    "Status",
    "Error Message",
)

# Statuswerte
ST_SUCCESS = "SUCCESS"
ST_SUCCESS_NORM = "SUCCESS_NAME_NORMALIZED"
ST_DUPLICATE = "DUPLICATE_IDENTICAL"
ST_RENAMED = "TARGET_NAME_CONFLICT_RENAMED"
ST_SKIPPED_QUERY = "SKIPPED_QUERY_EXECUTED"
ST_PDF_NOT_FOUND = "PDF_NOT_FOUND"
ST_INVALID_DATE = "INVALID_DATE"
ST_MISSING_DATA = "MISSING_DATA"
ST_NO_ISIN_DATE = "NO_ISIN_DATE_MATCH"
ST_NO_BO_MATCH = "NO_BO_MATCH"
ST_AMBIGUOUS = "AMBIGUOUS_MATCH"
ST_COPY_ERROR = "COPY_ERROR"
ST_XLSX_ERROR = "XLSX_ERROR"
ST_ERROR = "ERROR"

SUCCESS_STATES = frozenset({ST_SUCCESS, ST_SUCCESS_NORM, ST_RENAMED})

INVALID_WINDOWS_CHARS = '<>:"/\\|?*'
RESERVED_WINDOWS_NAMES = frozenset(
    ["CON", "PRN", "AUX", "NUL"]
    + [f"COM{i}" for i in range(1, 10)]
    + [f"LPT{i}" for i in range(1, 10)]
)

ISIN_PATTERN = re.compile(r"^[A-Z]{2}[A-Z0-9]{9}[0-9]$")

DATE_FORMATS = (
    "%d.%m.%Y",
    "%Y-%m-%d",
    "%d/%m/%Y",
    "%d-%m-%Y",
    "%Y/%m/%d",
    "%d.%m.%y",
    "%Y-%m-%d %H:%M:%S",
    "%d.%m.%Y %H:%M:%S",
)

EXCEL_EPOCH = datetime(1899, 12, 30)
EXCEL_SERIAL_MAX = 2958465  # 31.12.9999


# --------------------------------------------------------------------------- #
# Hilfsfunktionen: Text- und Namensnormalisierung                              #
# --------------------------------------------------------------------------- #

_DASH_CHARS = "\u2010\u2011\u2012\u2013\u2014\u2015\u2212-"


def normalize_for_matching(value: object) -> str:
    """Normalisiert einen Namen ausschliesslich fuer Vergleichszwecke.

    * Unicode-Normalisierung (NFKD) und Entfernen von Diakritika
    * saemtliche Bindestrich-Varianten werden zu Leerzeichen
    * Interpunktion wird zu Leerzeichen
    * Kleinschreibung (casefold) und Kollabieren von Whitespace

    Der Originalwert bleibt unangetastet.
    """
    if value is None:
        return ""
    text = str(value)
    text = unicodedata.normalize("NFKD", text)
    text = "".join(ch for ch in text if not unicodedata.combining(ch))
    text = "".join(" " if ch in _DASH_CHARS else ch for ch in text)
    text = "".join(ch if (ch.isalnum() or ch.isspace()) else " " for ch in text)
    text = unicodedata.normalize("NFKC", text)
    text = re.sub(r"\s+", " ", text).strip().casefold()
    return text


def tokenize_name(value: object) -> Tuple[str, ...]:
    """Zerlegt einen normalisierten Namen in Tokens."""
    normalized = normalize_for_matching(value)
    return tuple(normalized.split()) if normalized else tuple()


def contains_token_sequence(
    haystack: Sequence[str], needle: Sequence[str], end_anchored: bool = False
) -> bool:
    """Prueft, ob ``needle`` als zusammenhaengende Wortfolge in ``haystack`` vorkommt."""
    n_h, n_n = len(haystack), len(needle)
    if n_n == 0 or n_n > n_h:
        return False
    if end_anchored:
        return tuple(haystack[-n_n:]) == tuple(needle)
    for start in range(n_h - n_n + 1):
        if tuple(haystack[start:start + n_n]) == tuple(needle):
            return True
    return False


def clean_str(value: object) -> str:
    """Wandelt einen Zellwert in einen getrimmten String um."""
    if value is None:
        return ""
    if isinstance(value, float) and value.is_integer():
        return str(int(value)).strip()
    return str(value).strip()


def single_line(value: object) -> str:
    """Entfernt Zeilenumbrueche, damit das CSV-Log einzeilig bleibt."""
    return re.sub(r"[\r\n]+", " ", clean_str(value))


# --------------------------------------------------------------------------- #
# Hilfsfunktionen: Datumsverarbeitung                                          #
# --------------------------------------------------------------------------- #

def _from_excel_serial(serial: float) -> Optional[date]:
    try:
        number = float(serial)
    except (TypeError, ValueError):
        return None
    if number <= 0 or number > EXCEL_SERIAL_MAX:
        return None
    try:
        return (EXCEL_EPOCH + timedelta(days=number)).date()
    except (OverflowError, OSError, ValueError):
        return None


def parse_date_value(value: object) -> Optional[date]:
    """Liest Excel-Serials, datetime/date-Objekte und Datumsstrings robust ein."""
    if value is None:
        return None
    if isinstance(value, datetime):
        return value.date()
    if isinstance(value, date):
        return value
    if isinstance(value, bool):
        return None
    if isinstance(value, (int, float)):
        return _from_excel_serial(value)

    text = clean_str(value)
    if not text:
        return None

    for fmt in DATE_FORMATS:
        try:
            return datetime.strptime(text, fmt).date()
        except ValueError:
            continue

    # ISO-Datum mit Zeitanteil, z. B. 2024-02-13T00:00:00
    iso_candidate = text.replace("Z", "").split("T")[0].strip()
    if iso_candidate != text:
        for fmt in DATE_FORMATS:
            try:
                return datetime.strptime(iso_candidate, fmt).date()
            except ValueError:
                continue

    # Reine Zahl als String -> Excel-Serial
    if re.fullmatch(r"\d+(\.\d+)?", text):
        return _from_excel_serial(float(text))

    return None


def to_iso(value: Optional[date]) -> str:
    return value.isoformat() if value else ""


# --------------------------------------------------------------------------- #
# Hilfsfunktionen: Dateisystem / Windows                                       #
# --------------------------------------------------------------------------- #

def sanitize_windows_component(name: str, fallback: str = "") -> str:
    """Bereinigt einen Namen so, dass er als Windows-Ordner-/Dateiname taugt."""
    cleaned = "".join(
        "_" if (ch in INVALID_WINDOWS_CHARS or ord(ch) < 32) else ch for ch in str(name)
    )
    cleaned = re.sub(r"\s+", " ", cleaned).strip()
    cleaned = cleaned.rstrip(". ")
    if not cleaned:
        return fallback
    stem = cleaned.split(".")[0].upper()
    if stem in RESERVED_WINDOWS_NAMES:
        cleaned = f"_{cleaned}"
    return cleaned[:120]


def looks_like_suspicious_filename(name: str) -> bool:
    """Erkennt Pfadbestandteile, Traversals und ungueltige Zeichen."""
    if not name:
        return True
    if "/" in name or "\\" in name:
        return True
    if name.strip() in {".", ".."}:
        return True
    if any(ord(ch) < 32 for ch in name):
        return True
    if any(ch in '<>:"|?*' for ch in name):
        return True
    return False


def is_within(child: Path, parent: Path) -> bool:
    try:
        child.resolve().relative_to(parent.resolve())
        return True
    except (ValueError, OSError):
        return False


def sha256_of_file(path: Path, chunk_size: int = 1024 * 1024) -> str:
    digest = hashlib.sha256()
    with path.open("rb") as handle:
        for chunk in iter(lambda: handle.read(chunk_size), b""):
            digest.update(chunk)
    return digest.hexdigest()


# --------------------------------------------------------------------------- #
# Datenmodelle                                                                 #
# --------------------------------------------------------------------------- #

@dataclass(frozen=True)
class CsvRecord:
    """Eine relevante Zeile der Transactions-CSV."""
    row_number: int
    reclaim_reference: str
    bo_name_raw: str
    bo_tokens: Tuple[str, ...]
    isin: str
    value_date: Optional[date]


@dataclass
class MatchResult:
    record: Optional[CsvRecord] = None
    method: str = ""
    token_count: int = 0
    status: str = ST_NO_BO_MATCH
    candidates: List[CsvRecord] = field(default_factory=list)
    message: str = ""


@dataclass
class Counters:
    folders: int = 0
    xlsx_files: int = 0
    data_rows: int = 0
    copied: int = 0
    duplicates: int = 0
    renamed: int = 0
    pdf_missing: int = 0
    no_match: int = 0
    ambiguous: int = 0
    other_errors: int = 0
    skipped_query_rows: int = 0


# --------------------------------------------------------------------------- #
# Audit-Log                                                                    #
# --------------------------------------------------------------------------- #

class AuditLog:
    """Sammelt alle Log-Zeilen und schreibt sie am Ende als CSV (UTF-8 mit BOM)."""

    def __init__(self) -> None:
        self.rows: List[Dict[str, str]] = []

    def add(self, **kwargs: object) -> Dict[str, str]:
        row = {column: "" for column in LOG_COLUMNS}
        row["Timestamp"] = datetime.now().isoformat(timespec="seconds")
        for key, value in kwargs.items():
            column = key.replace("_", " ").title()
            # Feste Zuordnung fuer Sonderfaelle der Titel-Schreibweise
            column = {
                "Xlsx File": "XLSX File",
                "Xlsx Bo Name": "XLSX BO Name",
                "Xlsx Isin": "XLSX ISIN",
                "Xlsx Value Date": "XLSX Value Date",
                "Matched Csv Bo Name": "Matched CSV BO Name",
                "Original Pdf Name": "Original PDF Name",
                "Source Pdf Path": "Source PDF Path",
                "Target Pdf Name": "Target PDF Name",
                "Target Pdf Path": "Target PDF Path",
                "Candidate Count After Isin And Date": "Candidate Count After ISIN And Date",
                "Candidate Reclaim References": "Candidate Reclaim References",
            }.get(column, column)
            if column not in row:
                raise KeyError(f"Unbekannte Log-Spalte: {column}")
            row[column] = single_line(value)
        self.rows.append(row)
        return row

    def write(self, target: Path) -> None:
        with target.open("w", encoding="utf-8-sig", newline="") as handle:
            writer = csv.DictWriter(
                handle,
                fieldnames=list(LOG_COLUMNS),
                delimiter=";",
                quoting=csv.QUOTE_MINIMAL,
                lineterminator="\r\n",
            )
            writer.writeheader()
            writer.writerows(self.rows)


# --------------------------------------------------------------------------- #
# CSV einlesen                                                                 #
# --------------------------------------------------------------------------- #

def _read_csv_with_pandas(path: Path) -> Tuple[List[Dict[str, str]], List[str], str]:
    last_error: Optional[Exception] = None
    for encoding in CSV_ENCODINGS:
        try:
            frame = pd.read_csv(  # type: ignore[union-attr]
                path,
                sep=";",
                dtype=str,
                encoding=encoding,
                keep_default_na=False,
                na_values=[],
                engine="python",
            )
        except Exception as exc:  # UnicodeDecodeError, ParserError, ...
            last_error = exc
            continue
        frame.columns = [str(col).strip() for col in frame.columns]
        records = [
            {str(key): ("" if value is None else str(value)) for key, value in row.items()}
            for row in frame.to_dict(orient="records")
        ]
        return records, list(frame.columns), encoding
    raise RuntimeError(f"CSV konnte nicht gelesen werden: {last_error}")


def _read_csv_with_stdlib(path: Path) -> Tuple[List[Dict[str, str]], List[str], str]:
    last_error: Optional[Exception] = None
    for encoding in CSV_ENCODINGS:
        try:
            with path.open("r", encoding=encoding, newline="") as handle:
                reader = csv.reader(handle, delimiter=";")
                rows = list(reader)
        except Exception as exc:
            last_error = exc
            continue
        if not rows:
            return [], [], encoding
        headers = [str(col).strip() for col in rows[0]]
        records: List[Dict[str, str]] = []
        for raw in rows[1:]:
            record: Dict[str, str] = {}
            for index, header in enumerate(headers):
                record[header] = raw[index] if index < len(raw) else ""
            records.append(record)
        return records, headers, encoding
    raise RuntimeError(f"CSV konnte nicht gelesen werden: {last_error}")


def load_transactions_csv(path: Path) -> Tuple[List[CsvRecord], List[str], str, List[str]]:
    """Liest die Transactions-CSV und baut daraus normalisierte Datensaetze."""
    if pd is not None:
        records, headers, encoding = _read_csv_with_pandas(path)
    else:
        LOGGER.warning("pandas nicht verfuegbar - Fallback auf das csv-Modul.")
        records, headers, encoding = _read_csv_with_stdlib(path)

    missing = [col for col in REQUIRED_CSV_COLUMNS if col not in headers]
    if missing:
        raise RuntimeError(
            "In der CSV fehlen erforderliche Spalten: " + ", ".join(missing)
            + " | Gefundene Header: " + ", ".join(headers)
        )

    parsed: List[CsvRecord] = []
    warnings: List[str] = []
    for index, record in enumerate(records, start=2):  # Zeile 1 = Header
        isin = clean_str(record.get(CSV_COL_ISIN))
        raw_date = record.get(CSV_COL_VALUE_DATE)
        value_date = parse_date_value(raw_date)
        reclaim = clean_str(record.get(CSV_COL_RECLAIM))
        bo_name = clean_str(record.get(CSV_COL_BO_NAME))

        if raw_date and value_date is None:
            warnings.append(f"CSV-Zeile {index}: Value Date nicht lesbar ({clean_str(raw_date)!r})")
        if not isin or value_date is None:
            continue
        parsed.append(
            CsvRecord(
                row_number=index,
                reclaim_reference=reclaim,
                bo_name_raw=bo_name,
                bo_tokens=tokenize_name(bo_name),
                isin=isin,
                value_date=value_date,
            )
        )
    return parsed, headers, encoding, warnings


def build_csv_index(records: Sequence[CsvRecord]) -> Dict[Tuple[str, str], List[CsvRecord]]:
    index: Dict[Tuple[str, str], List[CsvRecord]] = {}
    for record in records:
        key = (record.isin, to_iso(record.value_date))
        index.setdefault(key, []).append(record)
    return index


# --------------------------------------------------------------------------- #
# Matching                                                                     #
# --------------------------------------------------------------------------- #

def match_beneficial_owner(
    candidates: Sequence[CsvRecord], xlsx_bo_name: str
) -> MatchResult:
    """Ermittelt aus den ISIN/Datum-Kandidaten den eindeutigen Beneficial Owner."""
    result = MatchResult(candidates=list(candidates))

    if not candidates:
        result.status = ST_NO_ISIN_DATE
        result.message = "Keine CSV-Zeile mit passender ISIN und passendem Value Date."
        return result

    tokens = tokenize_name(xlsx_bo_name)
    if not tokens:
        result.status = ST_MISSING_DATA
        result.message = "BO Name aus der XLSX ist nach der Normalisierung leer."
        return result

    # 1) Exakte Uebereinstimmung des vollstaendig normalisierten Namens
    exact = [c for c in candidates if c.bo_tokens == tokens]
    if exact and len({c.reclaim_reference for c in exact}) == 1:
        result.record = exact[0]
        result.method = "EXACT_NORMALIZED_NAME"
        result.token_count = len(tokens)
        result.status = ST_SUCCESS
        return result

    # 2) Suffix-Tokens, schrittweise erweitert
    start = min(2, len(tokens))
    last_hits: List[CsvRecord] = []
    for count in range(start, len(tokens) + 1):
        suffix = tokens[-count:]
        hits = [c for c in candidates if contains_token_sequence(c.bo_tokens, suffix)]
        if not hits:
            if last_hits:
                # Auf der vorherigen Stufe gab es mehrere unterschiedliche Reclaim
                # References, die Erweiterung kann sie nicht aufloesen -> mehrdeutig.
                result.status = ST_AMBIGUOUS
                result.candidates = last_hits
                result.token_count = count - 1
                result.method = "SUFFIX_TOKENS"
                result.message = (
                    f"Mehrdeutig: {len(last_hits)} Treffer mit "
                    f"{len({c.reclaim_reference for c in last_hits})} unterschiedlichen "
                    f"Reclaim References bei {count - 1} Tokens; Erweiterung auf {count} "
                    "Tokens liefert keine Treffer mehr."
                )
            else:
                result.status = ST_NO_BO_MATCH
                result.token_count = count
                result.message = (
                    f"Kein CSV-Kandidat enthaelt die letzten {count} Namens-Tokens "
                    f"({' '.join(suffix)})."
                )
            return result

        references = {c.reclaim_reference for c in hits}
        if len(references) == 1:
            result.record = hits[0]
            result.method = "SUFFIX_TOKENS"
            result.token_count = count
            result.status = ST_SUCCESS
            return result

        # Zusaetzliche Praezisierung: Treffer exakt am Namensende
        end_hits = [
            c for c in hits if contains_token_sequence(c.bo_tokens, suffix, end_anchored=True)
        ]
        if end_hits and len({c.reclaim_reference for c in end_hits}) == 1:
            result.record = end_hits[0]
            result.method = "SUFFIX_TOKENS_END_ANCHORED"
            result.token_count = count
            result.status = ST_SUCCESS
            return result

        last_hits = hits

    result.status = ST_AMBIGUOUS
    result.candidates = last_hits or list(candidates)
    result.token_count = len(tokens)
    result.method = "SUFFIX_TOKENS"
    result.message = (
        "Auch mit dem vollstaendigen Namen bleiben mehrere unterschiedliche "
        "Reclaim References uebrig."
    )
    return result


# --------------------------------------------------------------------------- #
# XLSX einlesen                                                                #
# --------------------------------------------------------------------------- #

@dataclass
class XlsxRow:
    excel_row: int
    isin_raw: str
    date_raw: object
    bo_name_raw: str
    request_id_raw: str
    is_query_executed: bool


def read_xlsx(path: Path) -> Tuple[List[str], List[XlsxRow]]:
    """Liest die XLSX positionsbasiert; gibt Header und Datenzeilen zurueck."""
    if openpyxl is None:
        raise RuntimeError("openpyxl ist nicht installiert (pip install openpyxl).")

    workbook = openpyxl.load_workbook(path, read_only=True, data_only=True)
    try:
        sheet = workbook.worksheets[0]
        headers: List[str] = []
        rows: List[XlsxRow] = []

        for cells in sheet.iter_rows(values_only=False):
            row_index = cells[0].row if cells else 0
            values = [cell.value for cell in cells]

            if row_index == 1:
                headers = [clean_str(value) for value in values]
                continue

            if all(clean_str(value) == "" for value in values):
                continue

            first_cell = clean_str(values[0] if values else "")
            is_query = first_cell.casefold().startswith(QUERY_EXECUTED_PREFIX)

            def get(index: int) -> object:
                return values[index] if index < len(values) else None

            rows.append(
                XlsxRow(
                    excel_row=row_index,
                    isin_raw=clean_str(get(COL_ISIN)),
                    date_raw=get(COL_VALUE_DATE),
                    bo_name_raw=clean_str(get(COL_BO_NAME)),
                    request_id_raw=clean_str(get(COL_REQUEST_ID)),
                    is_query_executed=is_query,
                )
            )
        return headers, rows
    finally:
        workbook.close()


def find_single_xlsx(folder: Path) -> Tuple[Optional[Path], Optional[str]]:
    """Sucht genau eine XLSX-Datei im Ordner (Excel-Temp-Dateien werden ignoriert)."""
    candidates = sorted(
        p for p in folder.iterdir()
        if p.is_file()
        and p.suffix.lower() in {".xlsx", ".xlsm"}
        and not p.name.startswith("~$")
    )
    if not candidates:
        return None, "Keine XLSX-Datei im Unterordner gefunden."
    if len(candidates) > 1:
        return None, (
            "Mehr als eine XLSX-Datei im Unterordner: "
            + ", ".join(p.name for p in candidates)
        )
    return candidates[0], None


# --------------------------------------------------------------------------- #
# PDF-Zuordnung und Kopieren                                                   #
# --------------------------------------------------------------------------- #

def locate_pdf(folder: Path, request_id: str) -> Tuple[Optional[Path], Optional[str]]:
    """Findet die PDF exakt anhand des Werts aus Spalte N (+ '.pdf')."""
    expected = request_id if request_id.lower().endswith(".pdf") else f"{request_id}.pdf"

    if looks_like_suspicious_filename(expected):
        return None, f"Verdaechtiger oder ungueltiger PDF-Dateiname: {expected!r}"

    files = [p for p in folder.iterdir() if p.is_file()]
    exact = [p for p in files if p.name == expected]
    if len(exact) == 1:
        return exact[0], None
    if len(exact) > 1:  # praktisch nur auf case-sensitiven Systemen moeglich
        return None, f"Mehrere PDFs mit dem Namen {expected!r} gefunden."

    insensitive = [p for p in files if p.name.lower() == expected.lower()]
    if len(insensitive) == 1:
        return insensitive[0], None
    if len(insensitive) > 1:
        return None, (
            f"Mehrere PDFs mit dem erwarteten Namen (Gross-/Kleinschreibung): "
            + ", ".join(sorted(p.name for p in insensitive))
        )
    return None, f"PDF {expected!r} nicht im Quellordner gefunden."


def resolve_target_path(
    target_dir: Path,
    filename: str,
    source_pdf: Path,
    planned: Dict[str, str],
) -> Tuple[Path, str, str]:
    """Bestimmt den finalen Zielpfad ohne jemals zu ueberschreiben.

    Rueckgabe: (Zielpfad, Status oder '', Meldung)
    Status ist ``DUPLICATE_IDENTICAL`` oder ``TARGET_NAME_CONFLICT_RENAMED`` oder ''.
    """
    target = target_dir / filename
    source_hash: Optional[str] = None

    def get_source_hash() -> str:
        nonlocal source_hash
        if source_hash is None:
            source_hash = sha256_of_file(source_pdf)
        return source_hash

    def existing_hash(path: Path) -> Optional[str]:
        key = str(path).lower()
        if key in planned:
            return planned[key]
        if path.exists():
            try:
                return sha256_of_file(path)
            except OSError:
                return None
        return None

    current_hash = existing_hash(target)
    if current_hash is None:
        return target, "", ""

    if current_hash == get_source_hash():
        return target, ST_DUPLICATE, "Zieldatei existiert bereits und ist inhaltsgleich."

    stem, suffix = target.stem, target.suffix
    for counter in range(2, 1000):
        candidate = target_dir / f"{stem}_{counter}{suffix}"
        candidate_hash = existing_hash(candidate)
        if candidate_hash is None:
            return candidate, ST_RENAMED, (
                f"Zieldatei {target.name!r} existiert bereits mit abweichendem Inhalt - "
                f"kopiert als {candidate.name!r}."
            )
        if candidate_hash == get_source_hash():
            return candidate, ST_DUPLICATE, (
                f"Inhaltsgleiche Kopie existiert bereits als {candidate.name!r}."
            )
    raise RuntimeError("Zu viele Namenskonflikte fuer dieselbe Zieldatei.")


# --------------------------------------------------------------------------- #
# Verarbeitung                                                                 #
# --------------------------------------------------------------------------- #

class Processor:
    def __init__(
        self,
        input_dir: Path,
        output_dir: Path,
        csv_index: Dict[Tuple[str, str], List[CsvRecord]],
        audit: AuditLog,
        dry_run: bool,
    ) -> None:
        self.input_dir = input_dir
        self.output_dir = output_dir
        self.csv_index = csv_index
        self.audit = audit
        self.dry_run = dry_run
        self.counters = Counters()
        self.planned_targets: Dict[str, str] = {}
        self.created_dirs: Set[Path] = set()

    # -- Ordnersuche ------------------------------------------------------- #

    def collect_folders(self) -> List[Path]:
        folders: List[Path] = []
        for path in sorted(self.input_dir.rglob("*")):
            if not path.is_dir():
                continue
            if is_within(path, self.output_dir):
                continue
            folders.append(path)
        if not folders:
            folders = [self.input_dir]
        elif any(
            p.is_file() and p.suffix.lower() in {".xlsx", ".xlsm"}
            for p in self.input_dir.iterdir()
        ):
            folders.insert(0, self.input_dir)
        return folders

    # -- Hauptlauf --------------------------------------------------------- #

    def run(self) -> Counters:
        folders = self.collect_folders()
        for folder in folders:
            try:
                entries = list(folder.iterdir())
            except OSError as exc:
                self.counters.other_errors += 1
                self.audit.add(
                    source_folder=folder, status=ST_ERROR,
                    error_message=f"Ordner nicht lesbar: {exc}",
                )
                continue

            has_xlsx = any(
                p.is_file() and p.suffix.lower() in {".xlsx", ".xlsm"}
                and not p.name.startswith("~$")
                for p in entries
            )
            has_pdf = any(p.is_file() and p.suffix.lower() == ".pdf" for p in entries)
            if not has_xlsx and not has_pdf:
                continue

            self.counters.folders += 1
            self.process_folder(folder)
        return self.counters

    def process_folder(self, folder: Path) -> None:
        xlsx_path, error = find_single_xlsx(folder)
        if xlsx_path is None:
            self.counters.other_errors += 1
            self.audit.add(
                source_folder=folder, status=ST_XLSX_ERROR, error_message=error or "",
            )
            LOGGER.error("%s: %s", folder, error)
            return

        try:
            headers, rows = read_xlsx(xlsx_path)
        except Exception as exc:
            self.counters.other_errors += 1
            self.audit.add(
                source_folder=folder, xlsx_file=xlsx_path.name,
                status=ST_XLSX_ERROR, error_message=f"XLSX nicht lesbar: {exc}",
            )
            LOGGER.error("XLSX nicht lesbar: %s (%s)", xlsx_path, exc)
            return

        self.counters.xlsx_files += 1
        LOGGER.info("Verarbeite %s (%d Zeilen)", xlsx_path, len(rows))
        LOGGER.debug("Header von %s: %s", xlsx_path.name, headers)

        if len(headers) < MIN_REQUIRED_COLUMNS:
            LOGGER.warning(
                "%s: Es wurden nur %d Spalten erkannt (erwartet mindestens %d bis Spalte N).",
                xlsx_path.name, len(headers), MIN_REQUIRED_COLUMNS,
            )

        for row in sorted(rows, key=lambda r: r.excel_row):
            try:
                self.process_row(folder, xlsx_path, row)
            except Exception as exc:  # letzte Sicherheitsnetz-Ebene
                self.counters.other_errors += 1
                self.audit.add(
                    source_folder=folder, xlsx_file=xlsx_path.name, excel_row=row.excel_row,
                    status=ST_ERROR, error_message=f"Unerwarteter Fehler: {exc!r}",
                )
                LOGGER.exception("Unerwarteter Fehler in %s Zeile %s", xlsx_path, row.excel_row)

    # -- Einzelzeile ------------------------------------------------------- #

    def process_row(self, folder: Path, xlsx_path: Path, row: XlsxRow) -> None:
        base = dict(
            source_folder=folder,
            xlsx_file=xlsx_path.name,
            excel_row=row.excel_row,
            original_pdf_name=row.request_id_raw,
            xlsx_bo_name=row.bo_name_raw,
            xlsx_isin=row.isin_raw,
        )

        if row.is_query_executed:
            self.counters.skipped_query_rows += 1
            self.audit.add(
                **base, status=ST_SKIPPED_QUERY,
                error_message="Zeile beginnt mit 'Query executed:' und wird ignoriert.",
            )
            return

        self.counters.data_rows += 1

        # --- Pflichtfelder ------------------------------------------------ #
        missing: List[str] = []
        if not row.isin_raw:
            missing.append("ISIN (Spalte B)")
        if not row.bo_name_raw:
            missing.append("BO Name (Spalte I)")
        if not row.request_id_raw:
            missing.append("Clearstream Request ID (Spalte N)")
        if missing:
            self.counters.other_errors += 1
            self.audit.add(
                **base, status=ST_MISSING_DATA,
                error_message="Fehlende Pflichtfelder: " + ", ".join(missing),
            )
            return

        # --- Value Date --------------------------------------------------- #
        value_date = parse_date_value(row.date_raw)
        if value_date is None:
            self.counters.other_errors += 1
            self.audit.add(
                **base, status=ST_INVALID_DATE,
                error_message=f"Value Date nicht lesbar: {clean_str(row.date_raw)!r}",
            )
            return
        iso_date = to_iso(value_date)
        base["xlsx_value_date"] = iso_date

        isin = row.isin_raw.strip()
        isin_note = ""
        if not ISIN_PATTERN.match(isin.upper()):
            isin_note = f"Hinweis: ISIN {isin!r} entspricht nicht dem ueblichen ISIN-Format. "

        # --- PDF finden --------------------------------------------------- #
        pdf_path, pdf_error = locate_pdf(folder, row.request_id_raw)
        if pdf_path is None:
            self.counters.pdf_missing += 1
            self.audit.add(
                **base, status=ST_PDF_NOT_FOUND, error_message=isin_note + (pdf_error or ""),
            )
            LOGGER.warning("%s Zeile %s: %s", xlsx_path.name, row.excel_row, pdf_error)
            return
        base["source_pdf_path"] = pdf_path

        # --- CSV-Matching ------------------------------------------------- #
        candidates = self.csv_index.get((isin, iso_date), [])
        match = match_beneficial_owner(candidates, row.bo_name_raw)
        candidate_refs = sorted({c.reclaim_reference for c in match.candidates if c.reclaim_reference})
        base["candidate_count_after_isin_and_date"] = len(candidates)
        base["candidate_reclaim_references"] = " | ".join(candidate_refs)

        if match.record is None:
            if match.status == ST_NO_ISIN_DATE:
                self.counters.no_match += 1
            elif match.status == ST_NO_BO_MATCH:
                self.counters.no_match += 1
            elif match.status == ST_AMBIGUOUS:
                self.counters.ambiguous += 1
            else:
                self.counters.other_errors += 1
            self.audit.add(
                **base,
                match_method=match.method,
                match_token_count=match.token_count or "",
                status=match.status,
                error_message=isin_note + match.message,
            )
            return

        record = match.record
        base["matched_csv_bo_name"] = record.bo_name_raw

        if not record.reclaim_reference:
            self.counters.other_errors += 1
            self.audit.add(
                **base, match_method=match.method, match_token_count=match.token_count,
                status=ST_MISSING_DATA,
                error_message=isin_note + (
                    f"Reclaim Reference in CSV-Zeile {record.row_number} ist leer."
                ),
            )
            return

        base["reclaim_reference"] = record.reclaim_reference

        # --- Zielpfad ------------------------------------------------------ #
        folder_name = sanitize_windows_component(record.reclaim_reference)
        if not folder_name:
            self.counters.other_errors += 1
            self.audit.add(
                **base, match_method=match.method, match_token_count=match.token_count,
                status=ST_ERROR,
                error_message="Reclaim Reference ergibt keinen gueltigen Ordnernamen.",
            )
            return

        target_dir = self.output_dir / folder_name
        if not is_within(target_dir, self.output_dir):
            self.counters.other_errors += 1
            self.audit.add(
                **base, match_method=match.method, match_token_count=match.token_count,
                status=ST_ERROR,
                error_message=f"Zielordner {target_dir} liegt ausserhalb des Output-Ordners.",
            )
            return

        target_filename = f"{iso_date}_{sanitize_windows_component(isin, 'UNKNOWN_ISIN')}.pdf"

        try:
            final_target, conflict_status, conflict_message = resolve_target_path(
                target_dir, target_filename, pdf_path, self.planned_targets
            )
        except Exception as exc:
            self.counters.other_errors += 1
            self.audit.add(
                **base, match_method=match.method, match_token_count=match.token_count,
                target_pdf_name=target_filename, status=ST_ERROR,
                error_message=f"Zielpfad konnte nicht bestimmt werden: {exc}",
            )
            return

        base["target_pdf_name"] = final_target.name
        base["target_pdf_path"] = final_target

        # Status bestimmen
        if conflict_status == ST_DUPLICATE:
            status = ST_DUPLICATE
        elif conflict_status == ST_RENAMED:
            status = ST_RENAMED
        elif match.method == "EXACT_NORMALIZED_NAME" and (
            row.bo_name_raw.strip() == record.bo_name_raw.strip()
        ):
            status = ST_SUCCESS
        else:
            status = ST_SUCCESS_NORM

        messages = [isin_note, conflict_message]

        # --- Kopieren ------------------------------------------------------ #
        if conflict_status == ST_DUPLICATE:
            self.counters.duplicates += 1
        elif self.dry_run:
            messages.append("DRY-RUN: Es wurde nichts erstellt oder kopiert.")
            if status == ST_RENAMED:
                self.counters.renamed += 1
            else:
                self.counters.copied += 1
            try:
                self.planned_targets[str(final_target).lower()] = sha256_of_file(pdf_path)
            except OSError:
                pass
        else:
            try:
                if target_dir not in self.created_dirs:
                    target_dir.mkdir(parents=True, exist_ok=True)
                    self.created_dirs.add(target_dir)
            except OSError as exc:
                self.counters.other_errors += 1
                self.audit.add(
                    **base, match_method=match.method, match_token_count=match.token_count,
                    status=ST_COPY_ERROR,
                    error_message=f"Zielordner konnte nicht erstellt werden: {exc}",
                )
                return
            try:
                shutil.copy2(pdf_path, final_target)
                self.planned_targets[str(final_target).lower()] = sha256_of_file(final_target)
            except Exception as exc:
                self.counters.other_errors += 1
                self.audit.add(
                    **base, match_method=match.method, match_token_count=match.token_count,
                    status=ST_COPY_ERROR,
                    error_message=f"PDF konnte nicht kopiert werden: {exc}",
                )
                return
            if status == ST_RENAMED:
                self.counters.renamed += 1
            else:
                self.counters.copied += 1

        self.audit.add(
            **base,
            match_method=match.method,
            match_token_count=match.token_count,
            status=status,
            error_message=" ".join(m for m in messages if m).strip(),
        )


# --------------------------------------------------------------------------- #
# CLI und Validierung                                                          #
# --------------------------------------------------------------------------- #

def discover_csv(input_dir: Path) -> Optional[Path]:
    """Sucht automatisch nach einer Datei 'Transactions*.csv'."""
    search_dirs = [Path.cwd(), input_dir, input_dir.parent]
    seen: Set[Path] = set()
    found: List[Path] = []
    for directory in search_dirs:
        try:
            resolved = directory.resolve()
        except OSError:
            continue
        if resolved in seen or not resolved.is_dir():
            continue
        seen.add(resolved)
        found.extend(sorted(resolved.glob("Transactions*.csv")))
    unique = sorted({p.resolve() for p in found})
    if len(unique) == 1:
        return unique[0]
    if len(unique) > 1:
        LOGGER.error(
            "Mehrere Transactions-CSV-Dateien gefunden - bitte --csv explizit angeben:\n  %s",
            "\n  ".join(str(p) for p in unique),
        )
    return None


def build_arg_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(
        description=(
            "Ordnet Steuerbescheinigungen (PDF) anhand von XLSX-Metadaten und der "
            "Transactions-CSV den Reclaim References zu und kopiert sie umbenannt "
            "in eine Zielstruktur."
        ),
        formatter_class=argparse.ArgumentDefaultsHelpFormatter,
    )
    parser.add_argument("--input", "-i", default="Input", help="Hauptordner mit den Unterordnern")
    parser.add_argument("--csv", "-c", default=None, help="Pfad zur Transactions-CSV")
    parser.add_argument("--output", "-o", default="Output", help="Zielordner")
    parser.add_argument(
        "--dry-run", action="store_true",
        help="Nur analysieren und protokollieren, nichts erstellen oder kopieren",
    )
    parser.add_argument("--log-file", default=None, help="Optionale Textlogdatei")
    parser.add_argument(
        "--verbose", "-v", action="store_true", help="Ausfuehrliche Konsolenausgabe (DEBUG)",
    )
    parser.add_argument(
        "--allow-nested-output", action="store_true",
        help="Erlaubt einen Output-Ordner innerhalb des Input-Ordners (nicht empfohlen)",
    )
    return parser


def configure_logging(verbose: bool, log_file: Optional[str]) -> None:
    LOGGER.setLevel(logging.DEBUG)
    formatter = logging.Formatter("%(asctime)s | %(levelname)-8s | %(message)s")

    console = logging.StreamHandler(stream=sys.stdout)
    console.setLevel(logging.DEBUG if verbose else logging.INFO)
    console.setFormatter(formatter)
    LOGGER.addHandler(console)

    if log_file:
        handler = logging.FileHandler(log_file, encoding="utf-8")
        handler.setLevel(logging.DEBUG)
        handler.setFormatter(formatter)
        LOGGER.addHandler(handler)


def print_summary(counters: Counters, log_path: Path, dry_run: bool) -> None:
    lines = [
        "",
        "=" * 62,
        "ZUSAMMENFASSUNG" + ("  (DRY RUN - es wurde nichts kopiert)" if dry_run else ""),
        "=" * 62,
        f"Gefundene Unterordner                 : {counters.folders}",
        f"Verarbeitete XLSX-Dateien             : {counters.xlsx_files}",
        f"Verarbeitete Excel-Datenzeilen        : {counters.data_rows}",
        f"Erfolgreich kopierte PDFs             : {counters.copied}",
        f"Identische Duplikate (uebersprungen)  : {counters.duplicates}",
        f"Umbenannte Dateikonflikte             : {counters.renamed}",
        f"Nicht gefundene PDFs                  : {counters.pdf_missing}",
        f"Nicht gefundene Matches               : {counters.no_match}",
        f"Mehrdeutige Matches                   : {counters.ambiguous}",
        f"Sonstige Fehler                       : {counters.other_errors}",
        f"Ignorierte 'Query executed:'-Zeilen   : {counters.skipped_query_rows}",
        "-" * 62,
        f"Audit-Log: {log_path}",
        "=" * 62,
        "",
    ]
    print("\n".join(lines))


def main(argv: Optional[Sequence[str]] = None) -> int:
    args = build_arg_parser().parse_args(argv)
    configure_logging(args.verbose, args.log_file)

    input_dir = Path(args.input).expanduser()
    output_dir = Path(args.output).expanduser()

    # --- Validierung: Input ------------------------------------------------ #
    if not input_dir.exists() or not input_dir.is_dir():
        LOGGER.error("Input-Ordner existiert nicht: %s", input_dir.resolve())
        return 2

    # --- Validierung: CSV -------------------------------------------------- #
    if args.csv:
        csv_path = Path(args.csv).expanduser()
    else:
        discovered = discover_csv(input_dir)
        if discovered is None:
            LOGGER.error(
                "Keine eindeutige 'Transactions*.csv' gefunden. Bitte --csv angeben."
            )
            return 2
        csv_path = discovered
        LOGGER.info("CSV automatisch erkannt: %s", csv_path)

    if not csv_path.exists() or not csv_path.is_file():
        LOGGER.error("CSV-Datei existiert nicht: %s", csv_path)
        return 2

    # --- Validierung: Output nicht innerhalb des Inputs -------------------- #
    if is_within(output_dir, input_dir) and not args.allow_nested_output:
        LOGGER.error(
            "Der Output-Ordner (%s) liegt innerhalb des Input-Ordners (%s). "
            "Bitte einen Pfad ausserhalb waehlen oder --allow-nested-output setzen.",
            output_dir.resolve(), input_dir.resolve(),
        )
        return 2
    if is_within(input_dir, output_dir):
        LOGGER.error(
            "Der Input-Ordner liegt innerhalb des Output-Ordners - Abbruch."
        )
        return 2

    # --- CSV laden --------------------------------------------------------- #
    try:
        records, headers, encoding, warnings = load_transactions_csv(csv_path)
    except Exception as exc:
        LOGGER.error("CSV konnte nicht verarbeitet werden: %s", exc)
        return 2

    LOGGER.info(
        "CSV geladen: %s (%d verwertbare Zeilen, Encoding %s)",
        csv_path.name, len(records), encoding,
    )
    LOGGER.debug("CSV-Header: %s", headers)
    for warning in warnings[:50]:
        LOGGER.warning(warning)
    if len(warnings) > 50:
        LOGGER.warning("... sowie %d weitere CSV-Warnungen.", len(warnings) - 50)
    if not records:
        LOGGER.error("Die CSV enthaelt keine verwertbaren Zeilen (ISIN + Value Date).")
        return 2

    csv_index = build_csv_index(records)

    # --- Output vorbereiten (auch im Dry-Run fuer das Log) ----------------- #
    try:
        output_dir.mkdir(parents=True, exist_ok=True)
    except OSError as exc:
        LOGGER.error("Output-Ordner konnte nicht erstellt werden: %s", exc)
        return 2

    audit = AuditLog()
    processor = Processor(
        input_dir=input_dir,
        output_dir=output_dir,
        csv_index=csv_index,
        audit=audit,
        dry_run=bool(args.dry_run),
    )

    if args.dry_run:
        LOGGER.info("DRY RUN aktiv - es werden keine Ordner erstellt und keine PDFs kopiert.")

    counters = processor.run()

    log_path = output_dir / LOG_FILENAME
    try:
        audit.write(log_path)
    except OSError as exc:
        LOGGER.error("Audit-Log konnte nicht geschrieben werden: %s", exc)
        return 3

    print_summary(counters, log_path, bool(args.dry_run))

    problems = (
        counters.pdf_missing + counters.no_match + counters.ambiguous + counters.other_errors
    )
    return 1 if problems else 0


if __name__ == "__main__":
    try:
        sys.exit(main())
    except KeyboardInterrupt:  # pragma: no cover
        print("\nAbbruch durch Benutzer.", file=sys.stderr)
        sys.exit(130)
