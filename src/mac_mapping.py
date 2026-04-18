"""MAC jurisdiction + Microlyte SAM eligibility mapping."""

from __future__ import annotations

import json
import logging
from pathlib import Path

import pandas as pd

log = logging.getLogger(__name__)

NOVA_COUNTIES = {"arlington", "fairfax", "alexandria"}


def load_mac_config(config_path: Path) -> dict:
    with open(config_path, "r", encoding="utf-8") as f:
        cfg = json.load(f)
    return {k: v for k, v in cfg.items() if not k.startswith("_")}


def _is_nova(address1: str, city: str) -> bool:
    text = f"{address1} {city}".lower()
    return any(c in text for c in NOVA_COUNTIES)


def lookup(state: str, cfg: dict, address1: str = "", city: str = "") -> dict:
    """Return MAC + Microlyte eligibility for a state row.

    Virginia is special-cased: NoVA (Arlington / Fairfax / Alexandria) falls under
    Novitas and is NOT eligible. Rest of VA is Palmetto and eligible.
    """
    state = (state or "").strip().upper()
    entry = cfg.get(state)
    if not entry:
        return {
            "MAC Jurisdiction": "Unknown",
            "Microlyte Eligible": "Unknown",
            "MAC Notes": f"No MAC mapping for state '{state}'",
        }

    mac = entry.get("mac", "Unknown")
    eligible = "Yes" if entry.get("microlyte_eligible") else "No"
    notes = []
    if entry.get("verify_note"):
        notes.append(entry["verify_note"])

    if state == "VA":
        if _is_nova(address1, city):
            mac = "Novitas"
            eligible = "No"
            notes.append("NoVA carve-out: Arlington/Fairfax/Alexandria -> Novitas (LCD L35041 active)")
        else:
            notes.append("VA requires manual county-level review")

    return {
        "MAC Jurisdiction": mac,
        "Microlyte Eligible": eligible,
        "MAC Notes": "; ".join(notes),
    }


def enrich_frame(df: pd.DataFrame, cfg: dict) -> pd.DataFrame:
    df = df.copy()
    state_col = df["State"] if "State" in df.columns else pd.Series([""] * len(df))
    addr_col = df["Address 1"] if "Address 1" in df.columns else pd.Series([""] * len(df))
    city_col = df["City"] if "City" in df.columns else pd.Series([""] * len(df))

    rows = [lookup(s, cfg, a, c) for s, a, c in zip(state_col.fillna(""), addr_col.fillna(""), city_col.fillna(""))]
    df["MAC Jurisdiction"] = [r["MAC Jurisdiction"] for r in rows]
    df["Microlyte Eligible"] = [r["Microlyte Eligible"] for r in rows]
    df["MAC Notes"] = [r["MAC Notes"] for r in rows]
    return df
