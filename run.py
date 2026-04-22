"""Albacete MedDev lead list pipeline with call/email tracking persistence.

Drop AcuityMD CSVs and (optionally) the current Master_Lead_List_Tracker.xlsx
into ``data/input/``, then::

    python run.py

On every run the pipeline:

1. Reads all AcuityMD CSVs (merge + dedup by NPI).
2. Reads any ``Master_Lead_List_Tracker*.xlsx`` in the input folder and merges
   its call/email activity into ``data/cache/activity.json`` so nothing ever
   gets overwritten.
3. Enriches: practice type, drive-time tier, NPPES phones, email inference,
   MAC + Microlyte, Target Score / Incision Likelihood, personalized drafts.
4. Re-applies saved activity, then emits:
   - ``data/output/lead_list_enriched_<ts>.xlsx`` (Master-Tracker layout)
   - ``data/output/dashboard_<ts>.html`` (activity tracking dashboard)
"""

from __future__ import annotations

import argparse
import logging
import sys
from datetime import datetime
from pathlib import Path
from typing import Iterable

import pandas as pd

from src import (
    classify, dashboard, email_enrich, export, ingest, mac_mapping, mobile, nppes,
    outreach, routing, scoring, tier, tracking, web_phones,
)

ROOT = Path(__file__).resolve().parent
DEFAULT_INPUT = ROOT / "data" / "input"
DEFAULT_OUTPUT = ROOT / "data" / "output"
DEFAULT_NPPES_CACHE = ROOT / "data" / "cache" / "nppes_cache.json"
DEFAULT_ACTIVITY_CACHE = ROOT / "data" / "cache" / "activity.json"
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


def parse_args(argv: list[str] | None = None) -> argparse.Namespace:
    p = argparse.ArgumentParser(description="Albacete MedDev lead list pipeline")
    p.add_argument("--input-dir", type=Path, default=DEFAULT_INPUT)
    p.add_argument("--output-dir", type=Path, default=DEFAULT_OUTPUT)
    p.add_argument("--files", nargs="*", help="Specific CSV filenames under --input-dir")
    p.add_argument("--origin-lat", type=float, default=tier.NYC_MIDTOWN[0])
    p.add_argument("--origin-lon", type=float, default=tier.NYC_MIDTOWN[1])
    p.add_argument("--force-nppes", action="store_true")
    p.add_argument("--skip-nppes", action="store_true")
    p.add_argument("--nppes-cache-only", action="store_true", help="Use existing NPPES cache but skip new network lookups")
    p.add_argument("--skip-emails", action="store_true")
    p.add_argument("--skip-trackers", action="store_true", help="Ignore any Master_Lead_List_Tracker*.xlsx in input dir")
    p.add_argument("--output", default="lead_list_enriched")
    p.add_argument("--dashboard", default="dashboard")
    p.add_argument("--nppes-cache", type=Path, default=DEFAULT_NPPES_CACHE)
    p.add_argument("--activity-cache", type=Path, default=DEFAULT_ACTIVITY_CACHE)
    p.add_argument("-v", "--verbose", action="store_true")
    return p.parse_args(argv)


def main(argv: list[str] | None = None) -> int:
    args = parse_args(argv)
    _configure_logging(args.verbose)
    log = logging.getLogger("pipeline")

    csv_files = _selected_input_files(args.input_dir, args.files)
    if not csv_files:
        log.error("No CSVs to process in %s", args.input_dir)
        return 1
    log.info("Processing %d CSV file(s)", len(csv_files))

    tracker_files: list[Path] = []
    if not args.skip_trackers:
        tracker_files = tracking.find_tracker_files(args.input_dir)
        if tracker_files:
            log.info("Found %d tracker xlsx to merge activity from", len(tracker_files))

    frames = [ingest.read_csv(f) for f in csv_files]
    df = ingest.merge_frames(frames)
    if df.empty:
        log.error("Merged frame is empty; nothing to do")
        return 1

    hospital_keywords = classify.load_keywords(CONFIG_DIR / "hospital_keywords.json")
    mac_cfg = mac_mapping.load_mac_config(CONFIG_DIR / "mac_jurisdictions.json")
    templates = outreach.load_templates(CONFIG_DIR / "email_templates.json")

    df = classify.classify_frame(df, hospital_keywords)
    df = tier.tier_frame(df, origin=(args.origin_lat, args.origin_lon))
    df = routing.enrich_frame(df)

    if args.skip_nppes:
        log.info("Skipping NPPES verification (--skip-nppes)")
        for col in ["NPPES Found", "NPPES Phone", "NPPES Fax", "NPPES Credential", "NPPES Taxonomy", "NPPES Status"]:
            if col not in df.columns:
                df[col] = ""
        df["Verified Phone"] = df.get("Phone Number", pd.Series([""] * len(df))).fillna("")
        df["Phone Status"] = df["Verified Phone"].map(lambda v: "Verified" if v else "Missing")
    else:
        df = nppes.enrich_frame(df, cache_path=args.nppes_cache, force=args.force_nppes, cache_only=args.nppes_cache_only)

    if args.skip_emails:
        log.info("Skipping email enrichment (--skip-emails)")
        if "Email Status" not in df.columns:
            df["Email Status"] = ""
    else:
        df = email_enrich.enrich_frame(df)

    df = mac_mapping.enrich_frame(df, mac_cfg)
    df = outreach.enrich_frame(df, templates)
    df = scoring.enrich_frame(df)

    df = web_phones.apply(df, args.activity_cache.parent / "web_phones.json")
    # Re-run classification now that web-verified practice names are available
    # so any manual hospital/private overrides from web_phones take effect.
    df = classify.classify_frame(df, hospital_keywords)

    activity = tracking.ingest_trackers(tracker_files, args.activity_cache)
    edit_files = tracking.find_edit_files(args.activity_cache.parent)
    if edit_files:
        activity = tracking.apply_edit_files(edit_files, activity)
        tracking.save_cache(args.activity_cache, activity)
    df = tracking.apply_activity(df, activity)
    df = tracking.summarize_last_touch(df)

    args.output_dir.mkdir(parents=True, exist_ok=True)
    stamp = datetime.utcnow().strftime("%Y%m%d_%H%M")
    xlsx_path = args.output_dir / f"{Path(args.output).stem}_{stamp}.xlsx"
    dash_path = args.output_dir / f"{Path(args.dashboard).stem}_{stamp}.html"
    mobile_path = args.output_dir / f"mobile_{stamp}.html"

    export.write_workbook(df, xlsx_path)
    dashboard.write_dashboard(df, dash_path)
    mobile.write_mobile(df, mobile_path)

    log.info("Pipeline complete: %d leads -> %s", len(df), xlsx_path.name)
    log.info("Dashboard: %s", dash_path.name)
    return 0


if __name__ == "__main__":
    sys.exit(main())
