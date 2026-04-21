"""Practice classification: Private Practice vs Hospital-Based."""

from __future__ import annotations

import json
import logging
from pathlib import Path
from typing import Iterable

import pandas as pd

log = logging.getLogger(__name__)

PRIVATE = "Private Practice"
HOSPITAL = "Hospital-Based"


def load_keywords(config_path: Path) -> dict:
    with open(config_path, "r", encoding="utf-8") as f:
        return json.load(f)


def _matches_any(text: str, patterns: Iterable[str]) -> bool:
    return any(p in text for p in patterns if p)


def classify_site(site_name: str, keywords: dict) -> str:
    """Return 'Private Practice' or 'Hospital-Based' for a site name."""
    if site_name is None:
        return PRIVATE
    name = str(site_name).strip().lower()
    if not name:
        return PRIVATE

    overrides = [p.lower() for p in keywords.get("private_practice_overrides", [])]
    if _matches_any(name, overrides):
        return PRIVATE

    named = [p.lower() for p in keywords.get("named_systems", [])]
    if _matches_any(name, named):
        return HOSPITAL

    generic = [p.lower() for p in keywords.get("generic_patterns", [])]
    if _matches_any(name, generic):
        return HOSPITAL

    return PRIVATE


def classify_frame(df: pd.DataFrame, keywords: dict, column: str = "Primary Site of Care") -> pd.DataFrame:
    df = df.copy()
    sites = df[column] if column in df.columns else pd.Series([""] * len(df))
    df["Practice Type"] = sites.fillna("").map(lambda s: classify_site(s, keywords))

    # If a web-verified practice name is present, let it override:
    # - "hospital-owned" / "hospital-affiliated" -> Hospital-Based
    # - any other verified name -> Private Practice (even if AcuityMD said hospital)
    # A doctor with BOTH a private office and hospital privileges is Private
    # for outreach purposes since we can sell into their private office.
    if "Web Practice" in df.columns:
        for idx in df.index:
            web = str(df.at[idx, "Web Practice"] or "").lower()
            if not web:
                continue
            if "hospital-owned" in web or "hospital-affiliated" in web or "(hospital" in web:
                df.at[idx, "Practice Type"] = HOSPITAL
            else:
                df.at[idx, "Practice Type"] = PRIVATE
    return df
