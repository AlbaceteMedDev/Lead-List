"""Personalized cold-outreach email generation."""

from __future__ import annotations

import json
import logging
from pathlib import Path
from typing import Optional

import pandas as pd

log = logging.getLogger(__name__)


def load_templates(config_path: Path) -> dict:
    with open(config_path, "r", encoding="utf-8") as f:
        return json.load(f)


def _num(value: object) -> float:
    if value is None:
        return 0.0
    s = str(value).replace(",", "").strip()
    if not s or s.lower() == "nan":
        return 0.0
    try:
        return float(s)
    except ValueError:
        return 0.0


def _first_matching_col(df_columns, needles: list[str]) -> Optional[str]:
    lowered = {c.lower(): c for c in df_columns}
    for needle in needles:
        for low, original in lowered.items():
            if needle in low:
                return original
    return None


def _procedure_focus(row: pd.Series, knee_col, hip_col, shoulder_col, rules: dict) -> str:
    knee = _num(row.get(knee_col)) if knee_col else 0.0
    hip = _num(row.get(hip_col)) if hip_col else 0.0
    shoulder = _num(row.get(shoulder_col)) if shoulder_col else 0.0
    knee_ratio = rules.get("knee_dominant_ratio", 1.5)
    hip_ratio = rules.get("hip_dominant_ratio", 1.5)
    shoulder_threshold = rules.get("shoulder_threshold", 200)

    if knee > hip * knee_ratio and knee > 0:
        return "knee replacement"
    if hip > knee * hip_ratio and hip > 0:
        return "hip replacement"
    if shoulder > shoulder_threshold:
        return "shoulder and joint replacement"
    return rules.get("default", "joint replacement")


def _volume_hook(total_vol: float, focus: str, rules: dict) -> str:
    high = rules.get("high_volume_threshold", 500)
    mid = rules.get("mid_volume_threshold", 200)
    if total_vol >= high:
        tmpl = rules.get("high", "")
    elif total_vol >= mid:
        tmpl = rules.get("mid", "")
    else:
        tmpl = rules.get("low", "")
    return tmpl.format(procedure_focus=focus)


def _practice_name(row: pd.Series) -> str:
    name = (row.get("Primary Site of Care") or "").strip()
    if not name:
        last = (row.get("Last Name") or "").strip() or "your practice"
        return f"Dr. {last}'s practice"
    return name


def generate_for_row(row: pd.Series, templates: dict, volume_cols: dict) -> dict:
    eligible = (row.get("Microlyte Eligible") or "").strip().lower()
    track_key = "track_b_propacks_plus_microlyte" if eligible == "yes" else "track_a_propacks_only"
    tmpl = templates[track_key]

    focus = _procedure_focus(
        row,
        volume_cols.get("knee"),
        volume_cols.get("hip"),
        volume_cols.get("shoulder"),
        templates.get("procedure_focus_rules", {}),
    )
    total_col = volume_cols.get("joint_repl")
    total = _num(row.get(total_col)) if total_col else 0.0
    hook = _volume_hook(total, focus, templates.get("volume_hook_rules", {}))

    merge = {
        "first_name": (row.get("First Name") or "").strip(),
        "last_name": (row.get("Last Name") or "").strip(),
        "practice_name": _practice_name(row),
        "city": (row.get("City") or "").strip(),
        "procedure_focus": focus,
        "volume_hook": hook,
        "sender_name": templates.get("sender_name", ""),
        "sender_title": templates.get("sender_title", ""),
        "sender_email": templates.get("sender_email", ""),
    }

    subject = tmpl["subject"].format(**merge)
    body = tmpl["body"].format(**merge)
    return {
        "Subject Line": subject,
        "Draft Email": body,
        "Email Track": tmpl["label"],
    }


def detect_volume_columns(columns) -> dict:
    """Locate AcuityMD procedure-volume columns regardless of exact phrasing."""
    def find(needle: str) -> Optional[str]:
        for c in columns:
            low = c.lower()
            if needle in low and "procedure volume" in low:
                return c
        return None

    joint_repl = (
        find("joint replacement") or find("total joint") or find("arthroplasty")
    )
    return {
        "joint_repl": joint_repl,
        "knee": find("knee"),
        "hip": find("hip"),
        "shoulder": find("shoulder"),
        "open_ortho": find("open ortho") or find("orthopedic"),
    }


def enrich_frame(df: pd.DataFrame, templates: dict) -> pd.DataFrame:
    df = df.copy()
    volume_cols = detect_volume_columns(df.columns)
    subjects, bodies, tracks = [], [], []
    for _, row in df.iterrows():
        out = generate_for_row(row, templates, volume_cols)
        subjects.append(out["Subject Line"])
        bodies.append(out["Draft Email"])
        tracks.append(out["Email Track"])
    df["Subject Line"] = subjects
    df["Draft Email"] = bodies
    df["Email Track"] = tracks

    rename_map = {}
    if volume_cols["joint_repl"]:
        rename_map[volume_cols["joint_repl"]] = "Joint Repl Vol"
    if volume_cols["knee"]:
        rename_map[volume_cols["knee"]] = "Knee Vol"
    if volume_cols["hip"]:
        rename_map[volume_cols["hip"]] = "Hip Vol"
    if volume_cols["shoulder"]:
        rename_map[volume_cols["shoulder"]] = "Shoulder Vol"
    if volume_cols["open_ortho"]:
        rename_map[volume_cols["open_ortho"]] = "Open Ortho Vol"
    if rename_map:
        df = df.rename(columns=rename_map)
    return df
