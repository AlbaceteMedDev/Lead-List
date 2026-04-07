"""Step 1: Ingest & Merge — Read AcuityMD CSV exports, merge, deduplicate by NPI."""

import glob
import logging
import os
import re

import pandas as pd

logger = logging.getLogger(__name__)

# Common identity columns shared across AcuityMD exports
IDENTITY_COLUMNS = [
    "HCP NPI", "First Name", "Last Name", "Middle Name", "Prefix",
    "Credential", "Specialty", "Phone Number", "Email",
    "Primary Site of Care", "Address 1", "Address 2", "City", "State",
    "Postal Code", "Medical School", "Medical School Graduation Year", "HCP URL",
]


def clean_column_name(col: str) -> str:
    """Strip extra quotes and whitespace from AcuityMD column names."""
    return re.sub(r'^"+|"+$', "", col).strip()


def read_csv_file(filepath: str) -> pd.DataFrame:
    """Read a single CSV file with all columns as strings."""
    logger.info(f"Reading {os.path.basename(filepath)}")
    df = pd.read_csv(filepath, dtype=str, keep_default_na=False)
    df.columns = [clean_column_name(c) for c in df.columns]
    # Replace empty strings with None for consistent null handling
    df = df.replace({"": None})
    logger.info(f"  → {len(df)} rows, {len(df.columns)} columns")
    return df


def merge_dataframes(dfs: list[pd.DataFrame]) -> pd.DataFrame:
    """Outer-join multiple DataFrames on HCP NPI, preserving all columns."""
    if not dfs:
        raise ValueError("No DataFrames to merge")
    if len(dfs) == 1:
        return dfs[0]

    merged = dfs[0]
    for i, df in enumerate(dfs[1:], start=2):
        # Find shared columns (identity) and unique columns (volumes)
        shared_cols = [c for c in merged.columns if c in df.columns]
        if "HCP NPI" not in shared_cols:
            logger.warning(f"CSV #{i} missing 'HCP NPI' — skipping merge")
            continue

        merged = pd.merge(merged, df, on="HCP NPI", how="outer", suffixes=("", f"_dup{i}"))

        # For shared identity columns, fill nulls from the duplicate
        for col in shared_cols:
            if col == "HCP NPI":
                continue
            dup_col = f"{col}_dup{i}"
            if dup_col in merged.columns:
                merged[col] = merged[col].fillna(merged[dup_col])
                merged.drop(columns=[dup_col], inplace=True)

    return merged


def deduplicate(df: pd.DataFrame) -> pd.DataFrame:
    """Deduplicate on HCP NPI, keeping the first non-null value per column."""
    if "HCP NPI" not in df.columns:
        raise ValueError("DataFrame missing 'HCP NPI' column")

    before = len(df)
    df = df.groupby("HCP NPI", as_index=False).first()
    after = len(df)
    if before != after:
        logger.info(f"Deduplicated: {before} → {after} rows ({before - after} duplicates removed)")
    return df


def ingest(input_dir: str = "data/input", files: list[str] | None = None) -> pd.DataFrame:
    """Read all CSVs from input_dir (or specific files), merge, and deduplicate."""
    if files:
        csv_paths = [os.path.join(input_dir, f) for f in files]
    else:
        csv_paths = sorted(glob.glob(os.path.join(input_dir, "*.csv")))

    if not csv_paths:
        raise FileNotFoundError(f"No CSV files found in {input_dir}")

    logger.info(f"Found {len(csv_paths)} CSV file(s)")
    dfs = [read_csv_file(p) for p in csv_paths]

    logger.info("Merging DataFrames...")
    merged = merge_dataframes(dfs)

    logger.info("Deduplicating on HCP NPI...")
    deduped = deduplicate(merged)

    logger.info(f"Ingestion complete: {len(deduped)} unique leads")
    return deduped
