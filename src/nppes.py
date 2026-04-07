"""Step 4: NPPES Phone Verification — Query CMS NPI Registry for phone/credential data."""

import json
import logging
import os
import re
import time
from datetime import datetime, timedelta

import requests

logger = logging.getLogger(__name__)

NPPES_API_URL = "https://npiregistry.cms.hhs.gov/api/"
CACHE_DIR = "data/cache"
CACHE_FILE = os.path.join(CACHE_DIR, "nppes_cache.json")
RATE_LIMIT_SLEEP = 0.15  # seconds between requests
CACHE_MAX_AGE_DAYS = 30


def normalize_phone(phone: str | None) -> str | None:
    """Normalize phone to 10-digit string."""
    if not phone:
        return None
    digits = re.sub(r"\D", "", str(phone))
    # Strip leading 1 for US country code
    if len(digits) == 11 and digits.startswith("1"):
        digits = digits[1:]
    return digits if len(digits) == 10 else None


def load_cache() -> dict:
    """Load NPPES cache from disk."""
    if os.path.exists(CACHE_FILE):
        with open(CACHE_FILE) as f:
            return json.load(f)
    return {}


def save_cache(cache: dict) -> None:
    """Save NPPES cache to disk."""
    os.makedirs(CACHE_DIR, exist_ok=True)
    with open(CACHE_FILE, "w") as f:
        json.dump(cache, f, indent=2)


def is_cache_fresh(entry: dict) -> bool:
    """Check if a cache entry is less than 30 days old."""
    ts = entry.get("cached_at")
    if not ts:
        return False
    cached_at = datetime.fromisoformat(ts)
    return datetime.now() - cached_at < timedelta(days=CACHE_MAX_AGE_DAYS)


def query_nppes(npi: str) -> dict | None:
    """Query the NPPES API for a single NPI. Returns parsed result or None."""
    try:
        resp = requests.get(
            NPPES_API_URL,
            params={"version": "2.1", "number": npi},
            timeout=10,
        )
        resp.raise_for_status()
        data = resp.json()

        if data.get("result_count", 0) == 0:
            return None

        result = data["results"][0]
        parsed = {"npi": npi}

        # Extract credential
        basic = result.get("basic", {})
        parsed["credential"] = basic.get("credential")
        parsed["enumeration_status"] = basic.get("status")

        # Extract location phone and fax
        for addr in result.get("addresses", []):
            if addr.get("address_purpose") == "LOCATION":
                parsed["phone"] = addr.get("telephone_number")
                parsed["fax"] = addr.get("fax_number")
                break

        # Extract primary taxonomy
        for tax in result.get("taxonomies", []):
            if tax.get("primary"):
                parsed["taxonomy"] = tax.get("desc")
                break

        return parsed

    except Exception as e:
        logger.error(f"NPPES query failed for NPI {npi}: {e}")
        return None


def verify_phones(df, force: bool = False):
    """Query NPPES for all leads, update phone data and add status columns."""
    import pandas as pd

    df = df.copy()
    cache = load_cache()

    nppes_phones = []
    phone_statuses = []
    nppes_credentials = []
    nppes_faxes = []

    total = len(df)
    cached_count = 0
    queried_count = 0

    for idx, row in df.iterrows():
        npi = str(row.get("HCP NPI", "")).strip()
        if not npi:
            nppes_phones.append(None)
            phone_statuses.append("Missing")
            nppes_credentials.append(row.get("Credential"))
            nppes_faxes.append(None)
            continue

        # Check cache
        if not force and npi in cache and is_cache_fresh(cache[npi]):
            result = cache[npi]
            cached_count += 1
        else:
            result = query_nppes(npi)
            if result:
                result["cached_at"] = datetime.now().isoformat()
                cache[npi] = result
            queried_count += 1
            time.sleep(RATE_LIMIT_SLEEP)

            if queried_count % 100 == 0:
                logger.info(f"  NPPES progress: {queried_count + cached_count}/{total}")
                save_cache(cache)  # periodic save

        if not result:
            nppes_phones.append(row.get("Phone Number"))
            phone_statuses.append("Missing")
            nppes_credentials.append(row.get("Credential"))
            nppes_faxes.append(None)
            continue

        # Phone comparison logic
        original_phone = normalize_phone(row.get("Phone Number"))
        nppes_phone = normalize_phone(result.get("phone"))

        if original_phone and nppes_phone and original_phone == nppes_phone:
            nppes_phones.append(nppes_phone)
            phone_statuses.append("Verified")
        elif not original_phone and nppes_phone:
            nppes_phones.append(nppes_phone)
            phone_statuses.append("Added from NPPES")
        elif original_phone and nppes_phone and original_phone != nppes_phone:
            nppes_phones.append(nppes_phone)
            phone_statuses.append("Updated (NPPES differs)")
        elif not original_phone and not nppes_phone:
            nppes_phones.append(None)
            phone_statuses.append("Missing")
        else:
            # Has original but NPPES has none — keep original
            nppes_phones.append(original_phone)
            phone_statuses.append("Verified")

        # Use NPPES credential if available
        cred = result.get("credential") or row.get("Credential")
        nppes_credentials.append(cred)
        nppes_faxes.append(result.get("fax"))

    # Save final cache
    save_cache(cache)

    df["Verified Phone"] = nppes_phones
    df["Phone Status"] = phone_statuses
    df["Credential"] = nppes_credentials
    df["Fax"] = nppes_faxes

    logger.info(f"NPPES verification complete: {cached_count} cached, {queried_count} queried")
    status_counts = pd.Series(phone_statuses).value_counts()
    for status, count in status_counts.items():
        logger.info(f"  {status}: {count}")

    return df
