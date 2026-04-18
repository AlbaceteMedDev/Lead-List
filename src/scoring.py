"""Target Score, Incision Likelihood, Best Approach, Why Target?.

The scoring mirrors the structure used in earlier hand-built Master Tracker
iterations (see `Master_Lead_List_Tracker (3) (10).xlsx`) so the output is
familiar to Gabe and his cold caller.

    Target Score  = practice_type + tier + microlyte + volume + incision
    Target Tier   = A+/A/B/C/D bucket of Target Score
    Lead Priority = A/B/C derived from Target Tier
    Best Approach = in-person / intro call / email nurture based on score + tier
"""

from __future__ import annotations

from typing import Optional

import pandas as pd

INCISION_HIGH = "High"
INCISION_MED_HIGH = "Medium-High"
INCISION_MED = "Medium"
INCISION_LOW = "Low"


def _num(value) -> float:
    if value is None:
        return 0.0
    s = str(value).replace(",", "").strip()
    if not s or s.lower() == "nan":
        return 0.0
    try:
        return float(s)
    except ValueError:
        return 0.0


def incision_likelihood(row: pd.Series, volume_cols: dict) -> str:
    """Classify how likely a surgeon is to produce large incisions.

    Drives Microlyte / ProPacks targeting. Large joint replacement, open
    orthopedic, and open spine volume all push up the likelihood.
    """
    joint = _num(row.get(volume_cols.get("joint_repl", ""), 0))
    open_ortho = _num(row.get(volume_cols.get("open_ortho", ""), 0))
    open_spine = _num(row.get(volume_cols.get("open_spine", ""), 0))
    hip = _num(row.get(volume_cols.get("hip", ""), 0))
    knee = _num(row.get(volume_cols.get("knee", ""), 0))

    large_total = joint + open_ortho + open_spine + hip + knee
    if large_total >= 400:
        return INCISION_HIGH
    if large_total >= 200:
        return INCISION_MED_HIGH
    if large_total >= 75:
        return INCISION_MED
    return INCISION_LOW


_INCISION_POINTS = {
    INCISION_HIGH: 20,
    INCISION_MED_HIGH: 15,
    INCISION_MED: 10,
    INCISION_LOW: 5,
}

_TIER_POINTS = {
    "Tier 1 (0-30 min)": 20,
    "Tier 2 (30-60 min)": 17,
    "Tier 3 (60-120 min)": 13,
    "Tier 4 (120-180 min)": 9,
    "Tier 5 (180+ min drivable)": 6,
    "Tier 6 (Requires flight)": 3,
    "Hospital-Based": 8,
}


def _volume_percentile_points(volume: float, top10: float, top25: float) -> int:
    if volume >= top10 and volume > 0:
        return 20
    if volume >= top25 and volume > 0:
        return 12
    if volume > 0:
        return 6
    return 0


def _volume_label(volume: float, top10: float, top25: float) -> str:
    if volume >= top10 and volume > 0:
        return "Top 10% volume"
    if volume >= top25 and volume > 0:
        return "Top 25% volume"
    if volume > 0:
        return "Active volume"
    return ""


def target_tier_label(score: int) -> str:
    if score >= 85:
        return "A+"
    if score >= 70:
        return "A"
    if score >= 55:
        return "B"
    if score >= 40:
        return "C"
    return "D"


def lead_priority(target_tier: str) -> str:
    return {"A+": "A", "A": "A", "B": "B", "C": "C"}.get(target_tier, "D")


def best_approach(score: int, tier: str, microlyte: str) -> str:
    if score >= 85 and tier in ("Tier 1 (0-30 min)", "Tier 2 (30-60 min)"):
        return "In-Person Lunch & Learn"
    if score >= 70:
        return "Intro Call + Follow-up Visit"
    if score >= 55:
        return "Intro Call"
    return "Email Nurture"


def why_target(row: pd.Series, volume_label: str, incision: str) -> str:
    parts: list[str] = []
    ptype = row.get("Practice Type", "")
    if ptype:
        parts.append(str(ptype))
    tier = row.get("Tier", "")
    if tier:
        short = tier.split(" (")[0] if "(" in tier else tier
        parts.append(f"{short} drive")
    if row.get("Microlyte Eligible") == "Yes":
        parts.append("Microlyte eligible")
    if volume_label:
        parts.append(volume_label)
    if incision in (INCISION_HIGH, INCISION_MED_HIGH):
        parts.append(f"{incision} incision")
    return " • ".join(parts)


def _detect_volume_columns(columns) -> dict:
    def find(needle: str) -> Optional[str]:
        for c in columns:
            low = c.lower()
            if needle in low and ("procedure volume" in low or low.endswith("vol")):
                return c
        return None
    return {
        "joint_repl": find("joint replacement") or find("total joint") or find("joint repl"),
        "knee": find("knee"),
        "hip": find("hip"),
        "shoulder": find("shoulder"),
        "open_ortho": find("open orthopedic") or find("open ortho"),
        "open_spine": find("open spine") or find("spine"),
    }


def enrich_frame(df: pd.DataFrame) -> pd.DataFrame:
    df = df.copy()
    volume_cols = _detect_volume_columns(df.columns)
    primary_vol_col = volume_cols.get("joint_repl") or volume_cols.get("open_spine") or volume_cols.get("open_ortho")

    primary = df[primary_vol_col].map(_num) if primary_vol_col else pd.Series([0.0] * len(df))
    nonzero = primary[primary > 0]
    top10 = nonzero.quantile(0.90) if not nonzero.empty else 0.0
    top25 = nonzero.quantile(0.75) if not nonzero.empty else 0.0

    incisions: list[str] = []
    scores: list[int] = []
    target_tiers: list[str] = []
    priorities: list[str] = []
    approaches: list[str] = []
    reasons: list[str] = []

    for idx, row in df.iterrows():
        vol = primary.iloc[df.index.get_loc(idx)] if primary_vol_col else 0.0
        incision = incision_likelihood(row, volume_cols)
        vol_pts = _volume_percentile_points(vol, top10, top25)
        vol_label = _volume_label(vol, top10, top25)

        practice_pts = 25 if row.get("Practice Type") == "Private Practice" else 10
        tier_pts = _TIER_POINTS.get(row.get("Tier", ""), 5)
        microlyte_pts = 15 if row.get("Microlyte Eligible") == "Yes" else 0
        score = practice_pts + tier_pts + microlyte_pts + vol_pts + _INCISION_POINTS[incision]
        score = int(min(100, score))

        t_tier = target_tier_label(score)
        incisions.append(incision)
        scores.append(score)
        target_tiers.append(t_tier)
        priorities.append(lead_priority(t_tier))
        approaches.append(best_approach(score, row.get("Tier", ""), row.get("Microlyte Eligible", "")))
        reasons.append(why_target(row, vol_label, incision))

    df["Lg Incision Likelihood"] = incisions
    df["Target Score"] = scores
    df["Target Tier"] = target_tiers
    df["Lead Priority"] = priorities
    df["Best Approach"] = approaches
    df["Why Target?"] = reasons
    return df
