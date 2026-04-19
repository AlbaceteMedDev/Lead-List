"""Assign each lead to a product line (JR / S&N / OOS) from its source CSV."""

from __future__ import annotations

import logging
from typing import Iterable

import pandas as pd

log = logging.getLogger(__name__)

PRODUCT_LINES = ["JR", "S&N", "OOS"]
PRODUCT_LINE_LABELS = {"JR": "Joint Replacement", "S&N": "Spine & Neuro", "OOS": "Outside Ortho & Spine"}


def _classify_source(source: str) -> str:
    s = (source or "").lower()
    if "joint_replacement" in s or "joint-replacement" in s:
        return "JR"
    if "spine" in s and "outisde" not in s and "outside" not in s:
        return "S&N"
    if "outisde" in s or "outside" in s:
        return "OOS"
    return "JR"


def _lead_primary_line(sources: str) -> str:
    """If a lead appears in multiple source files, pick the most specific one.

    Priority: JR (smallest list) > S&N > OOS.
    """
    tokens = [t for t in (sources or "").split(";") if t.strip()]
    lines = {_classify_source(t) for t in tokens}
    for pref in ("JR", "S&N", "OOS"):
        if pref in lines:
            return pref
    return "JR"


def enrich_frame(df: pd.DataFrame) -> pd.DataFrame:
    df = df.copy()
    sources = df["__source_file"] if "__source_file" in df.columns else pd.Series([""] * len(df))
    df["Product Line"] = sources.fillna("").map(_lead_primary_line)
    return df


def split_by_product_line(df: pd.DataFrame) -> dict[str, pd.DataFrame]:
    """Return {line_code: subset_df} for every product line present."""
    out: dict[str, pd.DataFrame] = {}
    if "Product Line" not in df.columns:
        df = enrich_frame(df)
    for line in PRODUCT_LINES:
        subset = df[df["Product Line"] == line]
        if not subset.empty:
            out[line] = subset.copy()
    return out
