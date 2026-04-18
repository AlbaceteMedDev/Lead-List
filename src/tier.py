"""Drive-time tiering from a configurable origin via zip code geocoding."""

from __future__ import annotations

import logging
import math
from typing import Optional, Tuple

import pandas as pd

log = logging.getLogger(__name__)

NYC_MIDTOWN = (40.7580, -73.9855)

TIER_THRESHOLDS = [
    (20, "Tier 1 (0-30 min)"),
    (45, "Tier 2 (30-60 min)"),
    (90, "Tier 3 (60-120 min)"),
    (140, "Tier 4 (120-180 min)"),
    (350, "Tier 5 (180+ min drivable)"),
]
TIER_FLIGHT = "Tier 6 (Requires flight)"
TIER_HOSPITAL = "Hospital-Based"
TIER_UNKNOWN_DEFAULT = "Tier 5 (180+ min drivable)"


def haversine_miles(lat1: float, lon1: float, lat2: float, lon2: float) -> float:
    r_miles = 3958.7613
    phi1, phi2 = math.radians(lat1), math.radians(lat2)
    dphi = math.radians(lat2 - lat1)
    dlambda = math.radians(lon2 - lon1)
    a = math.sin(dphi / 2) ** 2 + math.cos(phi1) * math.cos(phi2) * math.sin(dlambda / 2) ** 2
    return 2 * r_miles * math.asin(math.sqrt(a))


def miles_to_tier(miles: Optional[float]) -> str:
    if miles is None or math.isnan(miles):
        return TIER_UNKNOWN_DEFAULT
    for threshold, label in TIER_THRESHOLDS:
        if miles < threshold:
            return label
    return TIER_FLIGHT


def _normalize_zip(value: object) -> str:
    if value is None:
        return ""
    s = str(value).strip()
    if not s or s.lower() == "nan":
        return ""
    digits = "".join(c for c in s if c.isdigit())
    return digits[:5] if len(digits) >= 5 else ""


def _geocoder():
    try:
        import pgeocode  # type: ignore
    except ImportError:
        log.warning("pgeocode not installed; all leads will default to %s", TIER_UNKNOWN_DEFAULT)
        return None
    return pgeocode.Nominatim("us")


def tier_frame(
    df: pd.DataFrame,
    origin: Tuple[float, float] = NYC_MIDTOWN,
    zip_column: str = "Postal Code",
    practice_type_column: str = "Practice Type",
) -> pd.DataFrame:
    df = df.copy()
    nomi = _geocoder()

    zips = df[zip_column].map(_normalize_zip) if zip_column in df.columns else pd.Series([""] * len(df))
    unique_zips = [z for z in zips.unique() if z]

    zip_to_miles: dict[str, Optional[float]] = {}
    if nomi is not None and unique_zips:
        try:
            results = nomi.query_postal_code(unique_zips)
            if len(unique_zips) == 1:
                rows = [results]
            else:
                rows = [results.iloc[i] for i in range(len(results))]
            for z, row in zip(unique_zips, rows):
                lat = row.get("latitude") if hasattr(row, "get") else row["latitude"]
                lon = row.get("longitude") if hasattr(row, "get") else row["longitude"]
                if lat is None or lon is None or (isinstance(lat, float) and math.isnan(lat)):
                    zip_to_miles[z] = None
                else:
                    zip_to_miles[z] = haversine_miles(origin[0], origin[1], float(lat), float(lon))
        except Exception as exc:
            log.warning("pgeocode lookup failed (%s); defaulting to unknown", exc)

    miles_series = zips.map(lambda z: zip_to_miles.get(z))
    df["Miles From Origin"] = miles_series
    df["Tier"] = miles_series.map(miles_to_tier)

    if practice_type_column in df.columns:
        df.loc[df[practice_type_column] == "Hospital-Based", "Tier"] = TIER_HOSPITAL

    return df
