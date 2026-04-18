"""CSV ingestion, column normalization, merge, and dedup by NPI."""

from __future__ import annotations

import logging
import re
from pathlib import Path
from typing import Iterable

import pandas as pd

log = logging.getLogger(__name__)

IDENTITY_COLUMNS = [
    "HCP NPI",
    "First Name",
    "Last Name",
    "Middle Name",
    "Prefix",
    "Credential",
    "Specialty",
    "Phone Number",
    "Email",
    "Primary Site of Care",
    "Address 1",
    "Address 2",
    "City",
    "State",
    "Postal Code",
    "Medical School",
    "Medical School Graduation Year",
    "HCP URL",
]

COLUMN_ALIASES = {
    "npi": "HCP NPI",
    "hcp npi number": "HCP NPI",
    "provider npi": "HCP NPI",
    "first name": "First Name",
    "last name": "Last Name",
    "middle name": "Middle Name",
    "suffix": "Credential",
    "credential": "Credential",
    "specialty": "Specialty",
    "phone": "Phone Number",
    "phone number": "Phone Number",
    "email": "Email",
    "email address": "Email",
    "primary site of care": "Primary Site of Care",
    "practice name": "Primary Site of Care",
    "address 1": "Address 1",
    "address1": "Address 1",
    "address 2": "Address 2",
    "address2": "Address 2",
    "city": "City",
    "state": "State",
    "zip": "Postal Code",
    "zip code": "Postal Code",
    "postal code": "Postal Code",
    "medical school": "Medical School",
    "medical school graduation year": "Medical School Graduation Year",
    "hcp url": "HCP URL",
    "acuitymd url": "HCP URL",
}


def _clean_column_name(name: str) -> str:
    """Strip extra quotes AcuityMD wraps around headers."""
    if name is None:
        return name
    cleaned = str(name).strip()
    while cleaned.startswith('"') and cleaned.endswith('"') and len(cleaned) > 1:
        cleaned = cleaned[1:-1].strip()
    return cleaned


def _canonicalize(name: str) -> str:
    cleaned = _clean_column_name(name)
    alias = COLUMN_ALIASES.get(cleaned.lower())
    return alias or cleaned


def _normalize_npi(value: object) -> str:
    if value is None:
        return ""
    s = str(value).strip()
    if not s or s.lower() == "nan":
        return ""
    digits = re.sub(r"\D", "", s)
    return digits


def read_csv(path: Path) -> pd.DataFrame:
    """Read an AcuityMD CSV with all columns as strings and headers normalized."""
    df = pd.read_csv(path, dtype=str, keep_default_na=False, na_values=[""])
    df.columns = [_canonicalize(c) for c in df.columns]
    if "HCP NPI" not in df.columns:
        log.warning("No 'HCP NPI' column in %s; columns=%s", path.name, list(df.columns))
    else:
        df["HCP NPI"] = df["HCP NPI"].map(_normalize_npi)
        df = df[df["HCP NPI"] != ""]
    df["__source_file"] = path.name
    return df


def merge_frames(frames: Iterable[pd.DataFrame]) -> pd.DataFrame:
    """Outer-merge frames on HCP NPI, preserving all identity and volume columns.

    When the same NPI appears in multiple files, identity columns take the first
    non-empty value and procedure-volume columns are summed (or preserved) rather
    than duplicated.
    """
    frames = [f for f in frames if not f.empty]
    if not frames:
        return pd.DataFrame(columns=IDENTITY_COLUMNS)

    combined = pd.concat(frames, ignore_index=True, sort=False)
    combined["HCP NPI"] = combined["HCP NPI"].astype(str).map(_normalize_npi)
    combined = combined[combined["HCP NPI"] != ""]

    volume_cols = [c for c in combined.columns if "procedure volume" in c.lower() or c.lower().endswith(" vol")]
    identity_cols = [c for c in combined.columns if c not in volume_cols and c != "__source_file"]

    def _first_non_empty(series: pd.Series) -> object:
        for v in series:
            if v is not None and str(v).strip() != "" and str(v).lower() != "nan":
                return v
        return ""

    def _sum_numeric(series: pd.Series) -> object:
        total = 0.0
        saw_any = False
        for v in series:
            if v is None:
                continue
            s = str(v).replace(",", "").strip()
            if s == "" or s.lower() == "nan":
                continue
            try:
                total += float(s)
                saw_any = True
            except ValueError:
                continue
        return str(int(total)) if saw_any and total == int(total) else (str(total) if saw_any else "")

    agg = {c: _first_non_empty for c in identity_cols if c != "HCP NPI"}
    for c in volume_cols:
        agg[c] = _sum_numeric
    agg["__source_file"] = lambda s: ";".join(sorted({str(x) for x in s if x}))

    merged = combined.groupby("HCP NPI", as_index=False, sort=False).agg(agg)
    return merged


def ingest_directory(input_dir: Path) -> pd.DataFrame:
    """Read every CSV in ``input_dir`` and return a merged, deduped frame."""
    input_dir = Path(input_dir)
    csvs = sorted(input_dir.glob("*.csv"))
    if not csvs:
        log.warning("No CSV files found in %s", input_dir)
        return pd.DataFrame(columns=IDENTITY_COLUMNS)
    log.info("Ingesting %d CSV file(s) from %s", len(csvs), input_dir)
    frames = [read_csv(p) for p in csvs]
    merged = merge_frames(frames)
    log.info("Merged frame: %d unique NPIs, %d columns", len(merged), len(merged.columns))
    return merged
