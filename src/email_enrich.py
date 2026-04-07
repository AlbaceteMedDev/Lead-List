"""Step 5: Email Enrichment — Classify, infer from practice patterns, flag missing."""

import json
import logging
import re

import pandas as pd

logger = logging.getLogger(__name__)

DEFAULT_CONFIG = "config/email_templates.json"


def load_email_config(config_path: str = DEFAULT_CONFIG) -> dict:
    """Load email-related config (domains, prefixes)."""
    with open(config_path) as f:
        return json.load(f)


def _get_domain(email) -> str | None:
    """Extract domain from email address."""
    if not email or not isinstance(email, str) or "@" not in email:
        return None
    return email.split("@")[1].lower().strip()


def _get_local(email) -> str | None:
    """Extract local part from email address."""
    if not email or not isinstance(email, str) or "@" not in email:
        return None
    return email.split("@")[0].lower().strip()


def _name_in_local(local: str, first: str | None, last: str | None) -> bool:
    """Check if first or last name appears in the email local part."""
    if not local:
        return False
    local = local.lower()
    if first and len(first) > 1 and first.lower() in local:
        return True
    if last and len(last) > 1 and last.lower() in local:
        return True
    return False


def classify_email(email: str | None, first: str | None, last: str | None,
                   config: dict) -> str:
    """Pass 1: Classify an existing email."""
    if not email or not isinstance(email, str) or not email.strip():
        return "Missing"

    email = email.strip().lower()
    local = _get_local(email)
    domain = _get_domain(email)

    if not local or not domain:
        return "Missing"

    generic_prefixes = config.get("generic_email_prefixes", [])
    free_domains = [d.lower() for d in config.get("free_email_domains", [])]
    hospital_domains = [d.lower() for d in config.get("hospital_email_domains", [])]

    # Check generic office email
    base_local = local.split(".")[0] if "." in local else local
    for prefix in generic_prefixes:
        if local == prefix or local.startswith(prefix + "@") or local.startswith(prefix + "."):
            return "Generic Office Email"
        if base_local == prefix:
            return "Generic Office Email"

    # Check hospital system email
    if domain in hospital_domains:
        return "Hospital System Email"

    is_free = domain in free_domains
    has_name = _name_in_local(local, first, last)

    if not is_free and has_name:
        return "Verified (name + practice domain)"
    if is_free and has_name:
        return "Personal Email (name match)"
    if is_free and not has_name:
        return "Personal Email (no name match)"
    if not is_free and not has_name:
        return "Practice Email (review recommended)"

    return "Practice Email (review recommended)"


def _detect_email_pattern(email: str, first: str, last: str) -> str | None:
    """Detect the email format pattern used."""
    if not email or not first or not last:
        return None
    local = _get_local(email)
    if not local:
        return None

    first_l = first.lower()
    last_l = last.lower()

    # Check patterns in order of specificity
    if local == f"{first_l}.{last_l}":
        return "first.last"
    if local == f"{first_l}{last_l}":
        return "firstlast"
    if local == f"{first_l[0]}{last_l}":
        return "flast"
    if local == f"{first_l[0]}.{last_l}":
        return "f.last"
    if local == last_l:
        return "last"
    if local == f"{first_l}{last_l[0]}":
        return "firstl"
    if local == f"{last_l}{first_l[0]}":
        return "lastf"
    if local == f"{last_l}.{first_l[0]}":
        return "last.f"
    # Partial matches — first initial + last name is most common
    if local.startswith(first_l[0]) and last_l in local:
        return "flast"
    if first_l in local and last_l in local:
        return "firstlast"

    return None


def _generate_email(pattern: str, first, last, domain: str) -> str:
    """Generate an email address from a detected pattern."""
    if not first or not isinstance(first, str) or not last or not isinstance(last, str):
        return None
    first_l = first.lower()
    last_l = last.lower()

    formats = {
        "first.last": f"{first_l}.{last_l}@{domain}",
        "firstlast": f"{first_l}{last_l}@{domain}",
        "flast": f"{first_l[0]}{last_l}@{domain}",
        "f.last": f"{first_l[0]}.{last_l}@{domain}",
        "last": f"{last_l}@{domain}",
        "firstl": f"{first_l}{last_l[0]}@{domain}",
        "lastf": f"{last_l}{first_l[0]}@{domain}",
        "last.f": f"{last_l}.{first_l[0]}@{domain}",
    }
    return formats.get(pattern, f"{first_l[0]}{last_l}@{domain}")


def infer_missing_emails(df: pd.DataFrame, config: dict) -> pd.DataFrame:
    """Pass 2: Infer missing emails from practice colleague patterns."""
    df = df.copy()
    free_domains = set(d.lower() for d in config.get("free_email_domains", []))
    hospital_domains = set(d.lower() for d in config.get("hospital_email_domains", []))

    # Build practice → email patterns map
    practice_patterns = {}
    for _, row in df.iterrows():
        site = row.get("Primary Site of Care")
        email = row.get("Email")
        first = row.get("First Name")
        last = row.get("Last Name")

        if not site or not email or not first or not last:
            continue

        domain = _get_domain(email)
        if not domain or domain in free_domains:
            continue

        pattern = _detect_email_pattern(email, first, last)
        if pattern:
            practice_patterns.setdefault(site, []).append({
                "pattern": pattern,
                "domain": domain,
            })

    # For each missing email, try to infer
    inferred_count = 0
    for idx, row in df.iterrows():
        if row.get("Email Status") != "Missing":
            continue

        site = row.get("Primary Site of Care")
        first = row.get("First Name")
        last = row.get("Last Name")

        if not site or not first or not last:
            continue
        if site not in practice_patterns:
            continue

        # Find most common pattern and domain at this practice
        patterns = practice_patterns[site]
        domain_counts = {}
        pattern_counts = {}
        for p in patterns:
            d = p["domain"]
            if d not in free_domains and d not in hospital_domains:
                domain_counts[d] = domain_counts.get(d, 0) + 1
                pattern_counts[p["pattern"]] = pattern_counts.get(p["pattern"], 0) + 1

        if not domain_counts:
            # Allow hospital domains only if the lead is at that hospital
            for p in patterns:
                d = p["domain"]
                if d in hospital_domains:
                    domain_counts[d] = domain_counts.get(d, 0) + 1
                    pattern_counts[p["pattern"]] = pattern_counts.get(p["pattern"], 0) + 1

        if not domain_counts:
            continue

        best_domain = max(domain_counts, key=domain_counts.get)
        best_pattern = max(pattern_counts, key=pattern_counts.get)

        inferred_email = _generate_email(best_pattern, first, last, best_domain)
        df.at[idx, "Email"] = inferred_email
        df.at[idx, "Email Status"] = f"Inferred ({best_pattern}@{best_domain})"
        inferred_count += 1

    logger.info(f"Email inference: {inferred_count} emails inferred from practice patterns")
    return df


def enrich_emails(df: pd.DataFrame, config_path: str = DEFAULT_CONFIG) -> pd.DataFrame:
    """Run three-pass email enrichment."""
    config = load_email_config(config_path)
    df = df.copy()

    # Pass 1: Classify existing emails
    logger.info("Pass 1: Classifying existing emails...")
    df["Email Status"] = df.apply(
        lambda row: classify_email(
            row.get("Email"),
            row.get("First Name"),
            row.get("Last Name"),
            config,
        ),
        axis=1,
    )

    status_counts = df["Email Status"].value_counts()
    for status, count in status_counts.items():
        logger.info(f"  {status}: {count}")

    # Pass 2: Infer missing from practice patterns
    logger.info("Pass 2: Inferring missing emails from practice patterns...")
    df = infer_missing_emails(df, config)

    # Pass 3: Flag remaining missing
    missing_count = (df["Email Status"] == "Missing").sum()
    logger.info(f"Pass 3: {missing_count} emails still missing — flagged for manual sourcing")

    return df
