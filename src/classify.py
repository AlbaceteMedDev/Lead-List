"""Step 2: Practice Classification — Private Practice vs Hospital-Based."""

import json
import logging

import pandas as pd

logger = logging.getLogger(__name__)

DEFAULT_CONFIG = "config/hospital_keywords.json"


def load_keywords(config_path: str = DEFAULT_CONFIG) -> dict:
    """Load hospital keyword config."""
    with open(config_path) as f:
        return json.load(f)


def is_hospital_based(site_of_care: str | None, keywords: dict) -> bool:
    """Determine if a site of care is hospital-based using keyword matching."""
    if not site_of_care or not isinstance(site_of_care, str):
        return False

    site_lower = site_of_care.lower().strip()

    # Check named systems (case-insensitive substring match)
    for system in keywords.get("named_systems", []):
        if system.lower() in site_lower:
            return True

    # Check generic patterns with special handling for "university"
    for pattern in keywords.get("generic_patterns", []):
        pattern_lower = pattern.lower()
        # "VA " pattern needs special handling — match at start or after space
        if pattern_lower == "va ":
            if site_lower.startswith("va ") or " va " in site_lower:
                return True
        elif pattern_lower in site_lower:
            return True

    return False


def classify_practices(df: pd.DataFrame, config_path: str = DEFAULT_CONFIG) -> pd.DataFrame:
    """Add 'Practice Type' column: 'Private Practice' or 'Hospital-Based'."""
    keywords = load_keywords(config_path)

    df = df.copy()
    df["Practice Type"] = df["Primary Site of Care"].apply(
        lambda x: "Hospital-Based" if is_hospital_based(x, keywords) else "Private Practice"
    )

    hospital_count = (df["Practice Type"] == "Hospital-Based").sum()
    private_count = (df["Practice Type"] == "Private Practice").sum()
    logger.info(f"Classification: {private_count} Private Practice, {hospital_count} Hospital-Based")

    return df
