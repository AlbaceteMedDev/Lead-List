"""Albacete MedDev lead list processing pipeline entrypoint.

Drop raw AcuityMD CSV exports into ``data/input/`` and run::

    python run.py

This produces an enriched Excel workbook in ``data/output/`` plus an HTML
tracking dashboard. See ``CLAUDE.md`` / ``README.md`` for the full spec.
"""

from __future__ import annotations

import argparse
import logging
import sys
from datetime import datetime
from pathlib import Path
from typing import Iterable

import pandas as pd

from src import classify, dashboard, email_enrich, export, ingest, mac_mapping, nppes, outreach, tier

ROOT = Path(__file__).resolve().parent
DEFAULT_INPUT = ROOT / "data" / "input"
DEFAULT_OUTPUT = ROOT / "data" / "output"
DEFAULT_CACHE = ROOT / "data" / "cache" / "nppes_cache.json"
CONFIG_DIR = ROOT / "config"


def _configure_logging(verbose: bool) -> None:
    level = logging.DEBUG if verbose else logging.INFO
    logging.basicConfig(
        level=level,
        format="%(asctime)s %(levelname)-7s %(name)s | %(message)s",
        datefmt="%H:%M:%S",
    )


def _selected_input_files(input_dir: Path, explicit: Iterable[str] | None) -> list[Path]:
    if not explicit:
        return sorted(input_dir.glob("*.csv"))
    selected = []
    for name in explicit:
        p = Path(name)
        if not p.is_absolute():
            p = input_dir / p
        if not p.exists():
            raise SystemExit(f"Input file not found: {p}")
        selected.append(p)
    return selected


def _ingest_selected(files: list[Path]) -> pd.DataFrame:
    frames = [ingest.read_csv(f) for f in files]
    return ingest.merge_frames(frames)


def parse_args(argv: list[str] | None = None) -> argparse.Namespace:
    p = argparse.ArgumentParser(description="Albacete MedDev lead list pipeline")
    p.add_argument("--input-dir", type=Path, default=DEFAULT_INPUT)
    p.add_argument("--output-dir", type=Path, default=DEFAULT_OUTPUT)
    p.add_argument("--files", nargs="*", help="Specific CSV filenames under --input-dir")
    p.add_argument("--tiers", nargs="*", help="Only emit these tiers in the workbook (e.g. 1 2 3 4)")
    p.add_argument("--origin-lat", type=float, default=tier.NYC_MIDTOWN[0])
    p.add_argument("--origin-lon", type=float, default=tier.NYC_MIDTOWN[1])
    p.add_argument("--force-nppes", action="store_true", help="Re-query NPPES even if cached")
    p.add_argument("--skip-nppes", action="store_true")
    p.add_argument("--skip-emails", action="store_true")
    p.add_argument("--output", default="lead_list_enriched.xlsx")
    p.add_argument("--dashboard", default="dashboard.html")
    p.add_argument("--cache-path", type=Path, default=DEFAULT_CACHE)
    p.add_argument("-v", "--verbose", action="store_true")
    return p.parse_args(argv)


_TIER_LABELS = {
    "1": "Tier 1 (0-30 min)",
    "2": "Tier 2 (30-60 min)",
    "3": "Tier 3 (60-120 min)",
    "4": "Tier 4 (120-180 min)",
    "5": "Tier 5 (180+ min drivable)",
    "6": "Tier 6 (Requires flight)",
    "h": "Hospital-Based",
    "hospital": "Hospital-Based",
    "hospital-based": "Hospital-Based",
}


def _resolve_tier_labels(values: Iterable[str] | None) -> list[str] | None:
    if not values:
        return None
    out = []
    for v in values:
        label = _TIER_LABELS.get(str(v).lower())
        if label and label not in out:
            out.append(label)
    return out or None


def main(argv: list[str] | None = None) -> int:
    args = parse_args(argv)
    _configure_logging(args.verbose)
    log = logging.getLogger("pipeline")

    files = _selected_input_files(args.input_dir, args.files)
    if not files:
        log.error("No CSVs to process in %s", args.input_dir)
        return 1
    log.info("Processing %d file(s)", len(files))

    df = _ingest_selected(files)
    if df.empty:
        log.error("Merged frame is empty; nothing to do")
        return 1

    hospital_keywords = classify.load_keywords(CONFIG_DIR / "hospital_keywords.json")
    mac_cfg = mac_mapping.load_mac_config(CONFIG_DIR / "mac_jurisdictions.json")
    templates = outreach.load_templates(CONFIG_DIR / "email_templates.json")

    df = classify.classify_frame(df, hospital_keywords)
    df = tier.tier_frame(df, origin=(args.origin_lat, args.origin_lon))

    if args.skip_nppes:
        log.info("Skipping NPPES verification (--skip-nppes)")
        for col, default in [
            ("NPPES Found", False), ("NPPES Phone", ""), ("NPPES Fax", ""),
            ("NPPES Credential", ""), ("NPPES Taxonomy", ""), ("NPPES Status", ""),
            ("Verified Phone", df.get("Phone Number", "")),
            ("Phone Status", "Missing"),
        ]:
            if col not in df.columns:
                df[col] = default
    else:
        df = nppes.enrich_frame(df, cache_path=args.cache_path, force=args.force_nppes)

    if args.skip_emails:
        log.info("Skipping email enrichment (--skip-emails)")
        if "Email Status" not in df.columns:
            df["Email Status"] = ""
    else:
        df = email_enrich.enrich_frame(df)

    df = mac_mapping.enrich_frame(df, mac_cfg)
    df = outreach.enrich_frame(df, templates)

    args.output_dir.mkdir(parents=True, exist_ok=True)
    stamp = datetime.utcnow().strftime("%Y%m%d_%H%M")
    xlsx_path = args.output_dir / (Path(args.output).stem + f"_{stamp}.xlsx")
    dash_path = args.output_dir / (Path(args.dashboard).stem + f"_{stamp}.html")

    export.write_workbook(df, xlsx_path, tiers=_resolve_tier_labels(args.tiers))
    dashboard.write_dashboard(df, dash_path)

    log.info("Pipeline complete: %d leads -> %s", len(df), xlsx_path.name)
    log.info("Dashboard: %s", dash_path.name)
    return 0


if __name__ == "__main__":
    sys.exit(main())
