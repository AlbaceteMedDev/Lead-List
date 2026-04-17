"""Three-pass email enrichment: classify, infer from practice patterns, flag missing."""

from __future__ import annotations

import logging
import re
from collections import Counter
from typing import Iterable, Optional

import pandas as pd

log = logging.getLogger(__name__)

STATUS_MISSING = "Missing"
STATUS_GENERIC = "Generic Office Email"
STATUS_HOSPITAL = "Hospital System Email"
STATUS_VERIFIED = "Verified (name + practice domain)"
STATUS_PERSONAL_NAME = "Personal Email (name match)"
STATUS_PERSONAL_NO_NAME = "Personal Email (no name match)"
STATUS_PRACTICE_REVIEW = "Practice Email (review recommended)"
STATUS_INFERRED = "Inferred (pattern@domain)"

GENERIC_PREFIXES = {
    "info", "office", "admin", "contact", "billing", "reception",
    "appointments", "scheduling", "frontdesk", "front-desk", "help",
    "support", "mail", "general", "hello",
}

FREE_PROVIDERS = {
    "gmail.com", "yahoo.com", "hotmail.com", "aol.com", "outlook.com",
    "icloud.com", "me.com", "mac.com", "msn.com", "live.com",
    "comcast.net", "verizon.net", "att.net", "sbcglobal.net", "ymail.com",
    "protonmail.com", "proton.me",
}

HOSPITAL_DOMAINS = {
    "atlantichealth.org", "rwjbh.org", "hackensackmeridian.org",
    "hackensackumc.org", "nyu.edu", "nyulangone.org", "mountsinai.org",
    "mssm.edu", "northwell.edu", "hss.edu", "mskcc.org", "nyp.org",
    "cumc.columbia.edu", "weill.cornell.edu", "montefiore.org",
    "ynhh.org", "hhchealth.org", "partners.org", "mgh.harvard.edu",
    "bwh.harvard.edu", "dana-farber.org", "tuftsmedicalcenter.org",
    "pennmedicine.upenn.edu", "jefferson.edu", "temple.edu",
    "mainlinehealth.org", "lvhn.org", "upmc.edu", "ahn.org",
    "jhmi.edu", "jhu.edu", "medstarhealth.org", "christianacare.org",
    "geisinger.edu", "inova.org", "sentara.com", "dukehealth.org",
    "unchealth.unc.edu", "clevelandclinic.org", "uhhospitals.org",
    "mayoclinic.org", "kp.org", "va.gov",
}


def _norm(s: Optional[str]) -> str:
    return (s or "").strip().lower()


def _split_email(email: str) -> Optional[tuple[str, str]]:
    email = _norm(email)
    if not email or "@" not in email:
        return None
    local, _, domain = email.rpartition("@")
    if not local or not domain or "." not in domain:
        return None
    return local, domain


def _name_in_local(local: str, first: str, last: str) -> bool:
    local = local.lower()
    first = (first or "").lower()
    last = (last or "").lower()
    if last and last in local:
        return True
    if first and first in local:
        return True
    if first and last and len(first) >= 1 and (first[0] + last) in local:
        return True
    return False


def classify_email(email: Optional[str], first: str, last: str) -> str:
    if not email or _norm(email) == "":
        return STATUS_MISSING
    parts = _split_email(email)
    if parts is None:
        return STATUS_MISSING
    local, domain = parts
    if local in GENERIC_PREFIXES or any(local.startswith(p + ".") for p in GENERIC_PREFIXES):
        return STATUS_GENERIC
    if domain in HOSPITAL_DOMAINS or any(domain.endswith("." + d) for d in HOSPITAL_DOMAINS):
        return STATUS_HOSPITAL
    if domain in FREE_PROVIDERS:
        return STATUS_PERSONAL_NAME if _name_in_local(local, first, last) else STATUS_PERSONAL_NO_NAME
    if _name_in_local(local, first, last):
        return STATUS_VERIFIED
    return STATUS_PRACTICE_REVIEW


def _detect_pattern(local: str, first: str, last: str) -> Optional[str]:
    """Return a pattern template like 'first.last' given a sample email local part."""
    local = local.lower()
    first = (first or "").lower()
    last = (last or "").lower()
    if not first or not last:
        return None
    if local == f"{first}.{last}":
        return "first.last"
    if local == f"{first[0]}.{last}":
        return "f.last"
    if local == f"{first[0]}{last}":
        return "flast"
    if local == f"{first}{last}":
        return "firstlast"
    if local == f"{first}{last[0]}":
        return "firstl"
    if local == last:
        return "last"
    if local == first:
        return "first"
    return None


def _apply_pattern(pattern: str, first: str, last: str) -> Optional[str]:
    first = (first or "").lower()
    last = (last or "").lower()
    if not first or not last:
        return None
    return {
        "first.last": f"{first}.{last}",
        "f.last": f"{first[0]}.{last}",
        "flast": f"{first[0]}{last}",
        "firstlast": f"{first}{last}",
        "firstl": f"{first}{last[0]}",
        "last": last,
        "first": first,
    }.get(pattern)


def _domain_is_inferrable(domain: str) -> bool:
    if not domain:
        return False
    if domain in FREE_PROVIDERS:
        return False
    if domain in HOSPITAL_DOMAINS or any(domain.endswith("." + d) for d in HOSPITAL_DOMAINS):
        return False
    return True


def _practice_profile(rows: pd.DataFrame) -> Optional[tuple[str, str]]:
    """Given rows for a single practice, return (pattern, domain) if confident."""
    patterns: Counter = Counter()
    domains: Counter = Counter()
    for _, r in rows.iterrows():
        email = _norm(r.get("Email", ""))
        parts = _split_email(email)
        if not parts:
            continue
        local, domain = parts
        if not _domain_is_inferrable(domain):
            continue
        if local in GENERIC_PREFIXES:
            domains[domain] += 1
            continue
        pat = _detect_pattern(local, r.get("First Name", ""), r.get("Last Name", ""))
        if pat:
            patterns[pat] += 1
            domains[domain] += 1
    if not domains:
        return None
    domain = domains.most_common(1)[0][0]
    pattern = patterns.most_common(1)[0][0] if patterns else "first.last"
    return pattern, domain


def enrich_frame(df: pd.DataFrame) -> pd.DataFrame:
    df = df.copy()
    if "Email" not in df.columns:
        df["Email"] = ""
    if "Email Status" not in df.columns:
        df["Email Status"] = ""

    statuses = [
        classify_email(e, f, l)
        for e, f, l in zip(
            df["Email"].fillna(""),
            df.get("First Name", pd.Series([""] * len(df))).fillna(""),
            df.get("Last Name", pd.Series([""] * len(df))).fillna(""),
        )
    ]
    df["Email Status"] = statuses

    # Pass 2: infer missing from practice patterns
    if "Primary Site of Care" in df.columns:
        for site, group in df.groupby("Primary Site of Care"):
            if not site or pd.isna(site):
                continue
            profile = _practice_profile(group)
            if not profile:
                continue
            pattern, domain = profile
            missing_idx = group.index[df.loc[group.index, "Email Status"] == STATUS_MISSING]
            for idx in missing_idx:
                first = _norm(df.at[idx, "First Name"]) if "First Name" in df.columns else ""
                last = _norm(df.at[idx, "Last Name"]) if "Last Name" in df.columns else ""
                local = _apply_pattern(pattern, first, last)
                if not local:
                    continue
                inferred = f"{local}@{domain}"
                df.at[idx, "Email"] = inferred
                df.at[idx, "Email Status"] = STATUS_INFERRED

    return df
