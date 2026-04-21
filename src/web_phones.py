"""Web-sourced phone numbers (Google / Healthgrades / practice sites).

Phones come from a JSON file at ``data/cache/web_phones.json`` keyed by NPI::

    {
      "1234567890": {
        "phone": "6319683777",
        "source": "Healthgrades",
        "practice": "Orlin Cohen Medical Specialists",
        "address": "160 E Main St, Bay Shore, NY 11706",
        "fetched_at": "2026-04-21"
      }
    }

The file is grown by running WebSearch on priority leads (usually A+/A
Private Practice) and manually confirming the returned office phone.
For Private Practice leads, the Google-sourced phone supersedes NPPES
in the dashboard because NPPES self-reports can be years out of date.
"""

from __future__ import annotations

import json
import logging
import re
from pathlib import Path

import pandas as pd

log = logging.getLogger(__name__)


def _normalize_phone(value: object) -> str:
    if not value:
        return ""
    digits = re.sub(r"\D", "", str(value))
    if len(digits) == 11 and digits.startswith("1"):
        digits = digits[1:]
    return digits if len(digits) == 10 else ""


def load(cache_path: Path) -> dict[str, dict]:
    if not cache_path.exists():
        return {}
    try:
        return json.loads(cache_path.read_text(encoding="utf-8"))
    except Exception as exc:
        log.warning("Could not read web-phones cache %s: %s", cache_path, exc)
        return {}


def save(cache_path: Path, data: dict[str, dict]) -> None:
    cache_path.parent.mkdir(parents=True, exist_ok=True)
    tmp = cache_path.with_suffix(cache_path.suffix + ".tmp")
    tmp.write_text(json.dumps(data, indent=2, sort_keys=True), encoding="utf-8")
    tmp.replace(cache_path)


def apply(df: pd.DataFrame, cache_path: Path) -> pd.DataFrame:
    """Attach ``Web Phone`` / ``Web Phone Source`` / ``Web Practice`` columns."""
    df = df.copy()
    data = load(cache_path)
    if "HCP NPI" not in df.columns:
        return df

    def _get(npi: str, field: str) -> str:
        entry = data.get(str(npi))
        if not entry:
            return ""
        v = entry.get(field) or ""
        if field == "phone":
            return _normalize_phone(v)
        return str(v)

    df["Web Phone"] = df["HCP NPI"].map(lambda n: _get(n, "phone"))
    df["Web Phone Source"] = df["HCP NPI"].map(lambda n: _get(n, "source"))
    df["Web Practice"] = df["HCP NPI"].map(lambda n: _get(n, "practice"))
    df["Web Address"] = df["HCP NPI"].map(lambda n: _get(n, "address"))
    return df
