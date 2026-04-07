"""Step 3: Drive-Time Tiering — Zip code geocoding + haversine distance from origin."""

import logging
import math

import pandas as pd
import pgeocode

logger = logging.getLogger(__name__)

# NYC Midtown default origin
DEFAULT_LAT = 40.7580
DEFAULT_LON = -73.9855

# Straight-line mile thresholds → tier labels
TIER_THRESHOLDS = [
    (20, "Tier 1 (0-30 min)"),
    (45, "Tier 2 (30-60 min)"),
    (90, "Tier 3 (60-120 min)"),
    (140, "Tier 4 (120-180 min)"),
    (350, "Tier 5 (180+ drivable)"),
]
TIER_FLIGHT = "Tier 6 (Requires flight)"
TIER_DEFAULT = "Tier 5 (180+ drivable)"
TIER_HOSPITAL = "Hospital-Based"


def haversine_miles(lat1: float, lon1: float, lat2: float, lon2: float) -> float:
    """Calculate great-circle distance in miles between two lat/lon points."""
    R = 3958.8  # Earth radius in miles
    lat1, lon1, lat2, lon2 = map(math.radians, [lat1, lon1, lat2, lon2])
    dlat = lat2 - lat1
    dlon = lon2 - lon1
    a = math.sin(dlat / 2) ** 2 + math.cos(lat1) * math.cos(lat2) * math.sin(dlon / 2) ** 2
    return R * 2 * math.asin(math.sqrt(a))


def miles_to_tier(miles: float) -> str:
    """Convert straight-line miles to a drive-time tier."""
    for threshold, tier_name in TIER_THRESHOLDS:
        if miles <= threshold:
            return tier_name
    return TIER_FLIGHT


def assign_tiers(
    df: pd.DataFrame,
    origin_lat: float = DEFAULT_LAT,
    origin_lon: float = DEFAULT_LON,
) -> pd.DataFrame:
    """Assign drive-time tier to each lead based on zip code distance from origin."""
    df = df.copy()
    nomi = pgeocode.Nominatim("us")

    def get_tier(row):
        # Hospital-Based always gets its own tier
        if row.get("Practice Type") == "Hospital-Based":
            return TIER_HOSPITAL

        postal = row.get("Postal Code")
        if not postal or not isinstance(postal, str) or postal.strip() in ("", "nan", "None"):
            logger.warning(f"NPI {row.get('HCP NPI')}: no postal code, defaulting to {TIER_DEFAULT}")
            return TIER_DEFAULT

        # Take first 5 digits only
        zip5 = str(postal).strip()[:5]
        result = nomi.query_postal_code(zip5)

        if result is None or pd.isna(result.get("latitude")) or pd.isna(result.get("longitude")):
            logger.warning(f"NPI {row.get('HCP NPI')}: cannot geocode zip {zip5}, defaulting to {TIER_DEFAULT}")
            return TIER_DEFAULT

        miles = haversine_miles(origin_lat, origin_lon, result["latitude"], result["longitude"])
        return miles_to_tier(miles)

    logger.info("Assigning drive-time tiers...")
    df["Tier"] = df.apply(get_tier, axis=1)

    # Log tier distribution
    tier_counts = df["Tier"].value_counts()
    for tier, count in sorted(tier_counts.items()):
        logger.info(f"  {tier}: {count}")

    return df
