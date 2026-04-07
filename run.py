"""Lead List Processing Pipeline — Main entrypoint."""

import argparse
import logging
import sys
import time
from datetime import datetime

from src.ingest import ingest
from src.classify import classify_practices
from src.tier import assign_tiers
from src.nppes import verify_phones
from src.email_enrich import enrich_emails
from src.mac_mapping import map_mac_jurisdictions
from src.outreach import generate_outreach
from src.export import export_workbook

logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s [%(levelname)s] %(name)s: %(message)s",
    datefmt="%H:%M:%S",
)
logger = logging.getLogger("pipeline")


def parse_args():
    parser = argparse.ArgumentParser(description="Lead List Processing Pipeline")
    parser.add_argument(
        "--files", nargs="+",
        help="Specific CSV filenames in data/input/ to process",
    )
    parser.add_argument(
        "--tiers", nargs="+", type=int,
        help="Only include these tier numbers (e.g., 1 2 3 4)",
    )
    parser.add_argument(
        "--origin-lat", type=float, default=40.7580,
        help="Origin latitude for drive-time tiering (default: NYC Midtown)",
    )
    parser.add_argument(
        "--origin-lon", type=float, default=-73.9855,
        help="Origin longitude for drive-time tiering (default: NYC Midtown)",
    )
    parser.add_argument(
        "--force-nppes", action="store_true",
        help="Force re-query of NPPES API (ignore cache)",
    )
    parser.add_argument(
        "--skip-nppes", action="store_true",
        help="Skip NPPES verification entirely",
    )
    parser.add_argument(
        "--skip-emails", action="store_true",
        help="Skip email generation step",
    )
    parser.add_argument(
        "--output", type=str,
        help="Custom output filename (placed in data/output/)",
    )
    return parser.parse_args()


def main():
    args = parse_args()
    start = time.time()

    logger.info("=" * 60)
    logger.info("Lead List Processing Pipeline")
    logger.info("=" * 60)

    # Step 1: Ingest & Merge
    logger.info("\n>>> Step 1: Ingest & Merge")
    df = ingest(files=args.files)

    # Step 2: Practice Classification
    logger.info("\n>>> Step 2: Practice Classification")
    df = classify_practices(df)

    # Step 3: Drive-Time Tiering
    logger.info("\n>>> Step 3: Drive-Time Tiering")
    df = assign_tiers(df, origin_lat=args.origin_lat, origin_lon=args.origin_lon)

    # Step 4: NPPES Phone Verification
    if not args.skip_nppes:
        logger.info("\n>>> Step 4: NPPES Phone Verification")
        df = verify_phones(df, force=args.force_nppes)
    else:
        logger.info("\n>>> Step 4: NPPES Phone Verification (SKIPPED)")
        df["Verified Phone"] = df.get("Phone Number")
        df["Phone Status"] = "Skipped"

    # Step 5: Email Enrichment
    logger.info("\n>>> Step 5: Email Enrichment")
    df = enrich_emails(df)

    # Step 6: MAC Jurisdiction Mapping
    logger.info("\n>>> Step 6: MAC Jurisdiction Mapping")
    df = map_mac_jurisdictions(df)

    # Step 7: Personalized Email Generation
    if not args.skip_emails:
        logger.info("\n>>> Step 7: Outreach Email Generation")
        df = generate_outreach(df)
    else:
        logger.info("\n>>> Step 7: Outreach Email Generation (SKIPPED)")

    # Step 8: Excel Output
    logger.info("\n>>> Step 8: Excel Output")
    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    filename = args.output or f"lead_list_enriched_{timestamp}.xlsx"
    if not filename.endswith(".xlsx"):
        filename += ".xlsx"
    output_path = f"data/output/{filename}"

    export_workbook(df, output_path=output_path, tiers=args.tiers)

    elapsed = time.time() - start
    logger.info(f"\nPipeline complete in {elapsed:.1f}s")
    logger.info(f"Output: {output_path}")
    return 0


if __name__ == "__main__":
    sys.exit(main())
