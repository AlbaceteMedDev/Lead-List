"""CMS NPPES NPI Registry verification with local caching."""

from __future__ import annotations

import json
import logging
import re
import time
from datetime import datetime, timedelta, timezone
from pathlib import Path
from typing import Optional

import pandas as pd

log = logging.getLogger(__name__)

NPPES_ENDPOINT = "https://npiregistry.cms.hhs.gov/api/"
RATE_LIMIT_SLEEP = 0.15
CACHE_TTL_DAYS = 30

STATUS_VERIFIED = "Verified"
STATUS_ADDED = "Added from NPPES"
STATUS_UPDATED = "Updated (NPPES differs)"
STATUS_MISSING = "Missing"


def normalize_phone(value: object) -> str:
    if value is None:
        return ""
    digits = re.sub(r"\D", "", str(value))
    if len(digits) == 11 and digits.startswith("1"):
        digits = digits[1:]
    return digits if len(digits) == 10 else ""


def _load_cache(cache_path: Path) -> dict:
    if not cache_path.exists():
        return {}
    try:
        with open(cache_path, "r", encoding="utf-8") as f:
            return json.load(f)
    except Exception as exc:
        log.warning("Could not read NPPES cache %s: %s", cache_path, exc)
        return {}


def _save_cache(cache_path: Path, cache: dict) -> None:
    cache_path.parent.mkdir(parents=True, exist_ok=True)
    tmp = cache_path.with_suffix(cache_path.suffix + ".tmp")
    with open(tmp, "w", encoding="utf-8") as f:
        json.dump(cache, f)
    tmp.replace(cache_path)


def _cache_entry_fresh(entry: dict) -> bool:
    ts = entry.get("fetched_at")
    if not ts:
        return False
    try:
        fetched = datetime.fromisoformat(ts)
    except ValueError:
        return False
    if fetched.tzinfo is None:
        fetched = fetched.replace(tzinfo=timezone.utc)
    return (datetime.now(timezone.utc) - fetched) < timedelta(days=CACHE_TTL_DAYS)


def parse_response(payload: dict) -> dict:
    """Extract the fields we need from a NPPES API response."""
    out = {
        "nppes_phone": "",
        "nppes_fax": "",
        "nppes_credential": "",
        "nppes_taxonomy": "",
        "nppes_status": "",
        "nppes_found": False,
    }
    if not payload or not isinstance(payload, dict):
        return out
    results = payload.get("results") or []
    if not results:
        return out
    rec = results[0]
    out["nppes_found"] = True

    basic = rec.get("basic") or {}
    out["nppes_credential"] = (basic.get("credential") or "").strip()
    out["nppes_status"] = (basic.get("status") or "").strip()

    for addr in rec.get("addresses") or []:
        if (addr.get("address_purpose") or "").upper() == "LOCATION":
            out["nppes_phone"] = normalize_phone(addr.get("telephone_number"))
            out["nppes_fax"] = normalize_phone(addr.get("fax_number"))
            break

    for tax in rec.get("taxonomies") or []:
        if tax.get("primary"):
            out["nppes_taxonomy"] = (tax.get("desc") or "").strip()
            break

    return out


def _fetch_one(npi: str, session) -> dict:
    try:
        resp = session.get(NPPES_ENDPOINT, params={"version": "2.1", "number": npi}, timeout=15)
        resp.raise_for_status()
        return parse_response(resp.json())
    except Exception as exc:
        log.warning("NPPES lookup failed for %s: %s", npi, exc)
        return parse_response({})


def fetch_many(npis: list[str], cache_path: Path, force: bool = False, cache_only: bool = False) -> dict[str, dict]:
    """Fetch NPPES records for a list of NPIs, using the cache where possible.

    With ``cache_only=True``, don't hit the network at all — uncached NPIs get
    an empty placeholder. Used in CI to keep the deploy fast while the cache
    is grown offline.
    """
    cache = _load_cache(cache_path)

    session = None
    if not cache_only:
        try:
            import requests  # type: ignore
        except ImportError:
            log.error("requests not installed; skipping NPPES verification")
            return {n: parse_response({}) for n in npis}
        session = requests.Session()

    out: dict[str, dict] = {}
    new_lookups = 0
    for idx, npi in enumerate(npis, 1):
        if not npi:
            continue
        cached = cache.get(npi)
        if cached and not force and _cache_entry_fresh(cached):
            out[npi] = cached["data"]
            continue
        if cache_only:
            out[npi] = parse_response({})
            continue
        data = _fetch_one(npi, session)
        cache[npi] = {"fetched_at": datetime.now(timezone.utc).isoformat(), "data": data}
        out[npi] = data
        new_lookups += 1
        time.sleep(RATE_LIMIT_SLEEP)
        if new_lookups % 50 == 0:
            _save_cache(cache_path, cache)
            log.info("NPPES progress: %d new lookups (index %d/%d)", new_lookups, idx, len(npis))

    if new_lookups:
        _save_cache(cache_path, cache)
    log.info("NPPES done: %d total NPIs, %d new lookups, %d cached", len(npis), new_lookups, len(npis) - new_lookups)
    return out


def reconcile_phone(original: Optional[str], nppes_phone: str, nppes_found: bool) -> tuple[str, str]:
    """Return (verified_phone, phone_status)."""
    orig_norm = normalize_phone(original)
    if not orig_norm and not nppes_found:
        return "", STATUS_MISSING
    if not orig_norm and nppes_phone:
        return nppes_phone, STATUS_ADDED
    if not orig_norm and not nppes_phone:
        return "", STATUS_MISSING
    if orig_norm and not nppes_phone:
        return orig_norm, STATUS_MISSING if not nppes_found else STATUS_VERIFIED
    if orig_norm == nppes_phone:
        return nppes_phone, STATUS_VERIFIED
    return nppes_phone, STATUS_UPDATED


def enrich_frame(df: pd.DataFrame, cache_path: Path, force: bool = False, cache_only: bool = False) -> pd.DataFrame:
    df = df.copy()
    if "HCP NPI" not in df.columns:
        log.warning("No HCP NPI column; skipping NPPES enrichment")
        return df

    npis = [str(n) for n in df["HCP NPI"].tolist()]
    results = fetch_many(npis, cache_path=cache_path, force=force, cache_only=cache_only)

    def _row(npi: str) -> dict:
        return results.get(str(npi), parse_response({}))

    df["NPPES Found"] = df["HCP NPI"].map(lambda n: _row(n)["nppes_found"])
    df["NPPES Phone"] = df["HCP NPI"].map(lambda n: _row(n)["nppes_phone"])
    df["NPPES Fax"] = df["HCP NPI"].map(lambda n: _row(n)["nppes_fax"])
    df["NPPES Credential"] = df["HCP NPI"].map(lambda n: _row(n)["nppes_credential"])
    df["NPPES Taxonomy"] = df["HCP NPI"].map(lambda n: _row(n)["nppes_taxonomy"])
    df["NPPES Status"] = df["HCP NPI"].map(lambda n: _row(n)["nppes_status"])

    original_phone = df["Phone Number"] if "Phone Number" in df.columns else pd.Series([""] * len(df))
    verified, status = [], []
    for orig, npi in zip(original_phone.fillna(""), df["HCP NPI"]):
        r = _row(npi)
        v, s = reconcile_phone(orig, r["nppes_phone"], r["nppes_found"])
        verified.append(v)
        status.append(s)
    df["Verified Phone"] = verified
    df["Phone Status"] = status

    if "Credential" in df.columns:
        df["Credential"] = [
            (nc or oc) for nc, oc in zip(df["NPPES Credential"].fillna(""), df["Credential"].fillna(""))
        ]
    else:
        df["Credential"] = df["NPPES Credential"]

    return df
