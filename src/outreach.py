"""Step 7: Personalized Email Generation — Track A (ProPacks) or Track B (ProPacks + Microlyte)."""

import json
import logging

import pandas as pd

logger = logging.getLogger(__name__)

DEFAULT_CONFIG = "config/email_templates.json"


def load_templates(config_path: str = DEFAULT_CONFIG) -> dict:
    """Load email templates."""
    with open(config_path) as f:
        return json.load(f)


def _safe_float(val) -> float:
    """Safely convert a value to float, returning 0 on failure."""
    if val is None:
        return 0.0
    try:
        return float(val)
    except (ValueError, TypeError):
        return 0.0


def _detect_procedure_focus(row) -> str:
    """Auto-detect procedure focus from volume columns."""
    knee = _safe_float(row.get("Knee Joint Replacement - Procedure Volume"))
    hip = _safe_float(row.get("Hip Joint Replacement - Procedure Volume"))
    shoulder = _safe_float(row.get("Shoulder Joint Replacement - Procedure Volume"))

    if shoulder > 200:
        return "shoulder and joint replacement"
    if knee > hip * 1.5 and knee > 0:
        return "knee replacement"
    if hip > knee * 1.5 and hip > 0:
        return "hip replacement"
    return "joint replacement"


def _get_volume_hook(row, procedure_focus: str) -> str:
    """Generate volume-appropriate opening language."""
    joint_vol = _safe_float(row.get("Joint Replacement - Procedure Volume"))

    if joint_vol > 500:
        return f"With the volume of {procedure_focus} cases you're managing"
    if joint_vol >= 200:
        return f"Given your {procedure_focus} practice"
    return f"As an orthopedic surgeon performing {procedure_focus} procedures"


def generate_outreach(df: pd.DataFrame, config_path: str = DEFAULT_CONFIG) -> pd.DataFrame:
    """Generate personalized subject line and draft email for each lead."""
    templates = load_templates(config_path)
    sender = templates["sender"]

    df = df.copy()
    subject_lines = []
    draft_emails = []

    track_a_count = 0
    track_b_count = 0

    for _, row in df.iterrows():
        practice_name = row.get("Primary Site of Care") or "your practice"
        first_name = row.get("First Name") or ""
        last_name = row.get("Last Name") or ""
        city = row.get("City") or ""

        procedure_focus = _detect_procedure_focus(row)
        volume_hook = _get_volume_hook(row, procedure_focus)

        microlyte_eligible = str(row.get("Microlyte Eligible", "No")).strip()

        # Select track
        if microlyte_eligible == "Yes":
            track = templates["track_b"]
            track_b_count += 1
        else:
            track = templates["track_a"]
            track_a_count += 1

        # Generate subject
        subject = track["subject"].format(
            practice_name=practice_name,
            first_name=first_name,
            last_name=last_name,
            city=city,
        )

        # Generate body
        body = track["body"].format(
            first_name=first_name,
            last_name=last_name,
            practice_name=practice_name,
            city=city,
            procedure_focus=procedure_focus,
            volume_hook=volume_hook,
            sender_name=sender["name"],
            sender_title=sender["title"],
            sender_email=sender["email"],
        )

        subject_lines.append(subject)
        draft_emails.append(body)

    df["Subject Line"] = subject_lines
    df["Draft Email"] = draft_emails

    logger.info(f"Outreach generation: {track_a_count} Track A (ProPacks), {track_b_count} Track B (ProPacks + Microlyte)")
    return df
