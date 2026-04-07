"""Step 6: MAC Jurisdiction Mapping — State → MAC contractor → Microlyte eligibility."""

import json
import logging

import pandas as pd

logger = logging.getLogger(__name__)

DEFAULT_CONFIG = "config/mac_jurisdictions.json"


def load_mac_config(config_path: str = DEFAULT_CONFIG) -> dict:
    """Load MAC jurisdiction mappings."""
    with open(config_path) as f:
        return json.load(f)


def map_mac_jurisdictions(df: pd.DataFrame, config_path: str = DEFAULT_CONFIG) -> pd.DataFrame:
    """Add MAC Jurisdiction, Microlyte Eligible, and VA Review Flag columns."""
    mac_config = load_mac_config(config_path)
    df = df.copy()

    mac_jurisdictions = []
    microlyte_eligible = []
    va_review_flags = []

    for _, row in df.iterrows():
        state = row.get("State")
        if not state or not isinstance(state, str):
            mac_jurisdictions.append("Unknown")
            microlyte_eligible.append("No")
            va_review_flags.append("")
            continue

        state = state.strip().upper()
        mapping = mac_config.get(state)

        if not mapping:
            mac_jurisdictions.append("Unknown")
            microlyte_eligible.append("No")
            va_review_flags.append("")
            logger.warning(f"No MAC mapping for state: {state}")
            continue

        mac_jurisdictions.append(mapping["mac"])
        microlyte_eligible.append("Yes" if mapping.get("microlyte_eligible", False) else "No")

        # Virginia special handling
        if state == "VA":
            va_review_flags.append("REVIEW: VA county-level MAC determination required (NoVA → Novitas/No Microlyte)")
        else:
            va_review_flags.append("")

    df["MAC Jurisdiction"] = mac_jurisdictions
    df["Microlyte Eligible"] = microlyte_eligible
    df["VA Review Flag"] = va_review_flags

    # Log summary
    elig_counts = df["Microlyte Eligible"].value_counts()
    logger.info(f"Microlyte eligibility: {elig_counts.get('Yes', 0)} eligible, {elig_counts.get('No', 0)} not eligible")
    va_count = (df["VA Review Flag"] != "").sum()
    if va_count > 0:
        logger.info(f"  {va_count} Virginia leads flagged for county-level review")

    return df
