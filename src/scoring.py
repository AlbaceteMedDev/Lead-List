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


def target_tier_label(score: int, practice_type: str = "") -> str:
    """A+/A requires Private Practice - hospital systems are harder to sell into."""
    if practice_type == "Hospital-Based":
        if score >= 55:
            return "B"
        if score >= 40:
            return "C"
        if score >= 25:
            return "D"
        return "F"
    if score >= 85:
        return "A+"
    if score >= 70:
        return "A"
    if score >= 55:
        return "B"
    if score >= 40:
        return "C"
    if score >= 25:
        return "D"
    return "F"


def lead_priority(target_tier: str) -> str:
    return {"A+": "A", "A": "A", "B": "B", "C": "C", "D": "D"}.get(target_tier, "F")


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


def _tier_rationale(t_tier: str, practice_type: str) -> str:
    """One-line summary of what this grade means."""
    if practice_type == "Hospital-Based":
        capped = " (capped - hospital-based, hard to sell into)"
        return {
            "B": "Solid hospital-based prospect" + capped,
            "C": "Average hospital-based prospect" + capped,
            "D": "Low-priority hospital-based prospect" + capped,
            "F": "Skip - hospital-based and low score",
        }.get(t_tier, "")
    return {
        "A+": "Top priority - call this week",
        "A": "High priority - call within 2 weeks",
        "B": "Solid prospect - work into rotation",
        "C": "Average - email nurture first",
        "D": "Low priority - email only",
        "F": "Skip - weak fit",
    }.get(t_tier, "")


def _tier_reason_breakdown(
    t_tier: str,
    score: int,
    practice_type: str,
    tier: str,
    microlyte: str,
    incision: str,
    vol_label: str,
    vol_pts: int,
) -> str:
    """Per-lead explanation of exactly how the Target Tier was reached."""
    parts: list[str] = []
    parts.append(_tier_rationale(t_tier, practice_type))

    pos: list[str] = []
    neg: list[str] = []

    if practice_type == "Private Practice":
        pos.append("+25 Private Practice")
    elif practice_type == "Hospital-Based":
        neg.append("+10 Hospital-Based (tier capped at B)")
    else:
        neg.append(f"+10 {practice_type or 'unknown practice type'}")

    tier_pts = _TIER_POINTS.get(tier, 5)
    tier_short = tier.split(" (")[0] if tier and "(" in tier else (tier or "Unknown tier")
    if tier_pts >= 15:
        pos.append(f"+{tier_pts} {tier_short}")
    elif tier_pts >= 8:
        neg.append(f"+{tier_pts} {tier_short}")
    else:
        neg.append(f"+{tier_pts} {tier_short}")

    if microlyte == "Yes":
        pos.append("+15 Microlyte eligible (non-LCD MAC)")
    else:
        neg.append("+0 LCD state - Microlyte not eligible")

    if vol_pts >= 20:
        pos.append(f"+{vol_pts} {vol_label or 'top 10% volume'}")
    elif vol_pts >= 10:
        pos.append(f"+{vol_pts} {vol_label or 'top 25% volume'}")
    elif vol_pts > 0:
        neg.append(f"+{vol_pts} {vol_label or 'some volume'}")
    else:
        neg.append("+0 no procedure volume data")

    inc_pts = _INCISION_POINTS.get(incision, 0)
    if incision in (INCISION_HIGH, INCISION_MED_HIGH):
        pos.append(f"+{inc_pts} {incision} incision likelihood")
    else:
        neg.append(f"+{inc_pts} {incision} incision likelihood")

    parts.append(f"Score {score}/100")
    if pos:
        parts.append("Strengths: " + "; ".join(pos))
    if neg:
        parts.append("Weaknesses: " + "; ".join(neg))
    return " | ".join(p for p in parts if p)


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
    tier_reasons: list[str] = []

    for idx, row in df.iterrows():
        vol = primary.iloc[df.index.get_loc(idx)] if primary_vol_col else 0.0
        incision = incision_likelihood(row, volume_cols)
        vol_pts = _volume_percentile_points(vol, top10, top25)
        vol_label = _volume_label(vol, top10, top25)

        practice_type = row.get("Practice Type", "")
        tier = row.get("Tier", "")
        microlyte = row.get("Microlyte Eligible", "")

        practice_pts = 25 if practice_type == "Private Practice" else 10
        tier_pts = _TIER_POINTS.get(tier, 5)
        microlyte_pts = 15 if microlyte == "Yes" else 0
        score = practice_pts + tier_pts + microlyte_pts + vol_pts + _INCISION_POINTS[incision]
        score = int(min(100, score))

        t_tier = target_tier_label(score, practice_type)
        incisions.append(incision)
        scores.append(score)
        target_tiers.append(t_tier)
        priorities.append(lead_priority(t_tier))
        approaches.append(best_approach(score, tier, microlyte))
        reasons.append(why_target(row, vol_label, incision))
        tier_reasons.append(_tier_reason_breakdown(t_tier, score, practice_type, tier, microlyte, incision, vol_label, vol_pts))

    df["Lg Incision Likelihood"] = incisions
    df["Target Score"] = scores
    df["Target Tier"] = target_tiers
    df["Lead Priority"] = priorities
    df["Best Approach"] = approaches
    df["Why Target?"] = reasons
    df["Target Tier Reason"] = tier_reasons
    return df
