"""Scoring + routing tests."""

import sys
from pathlib import Path

import pandas as pd

sys.path.insert(0, str(Path(__file__).resolve().parent.parent))

from src import routing, scoring


def test_target_tier_buckets():
    assert scoring.target_tier_label(95) == "A+"
    assert scoring.target_tier_label(75) == "A"
    assert scoring.target_tier_label(60) == "B"
    assert scoring.target_tier_label(45) == "C"
    assert scoring.target_tier_label(20) == "D"


def test_lead_priority_mapping():
    assert scoring.lead_priority("A+") == "A"
    assert scoring.lead_priority("A") == "A"
    assert scoring.lead_priority("B") == "B"
    assert scoring.lead_priority("C") == "C"
    assert scoring.lead_priority("D") == "D"


def test_incision_likelihood_high():
    row = pd.Series({"Joint Replacement - Procedure Volume": "500"})
    vol_cols = {"joint_repl": "Joint Replacement - Procedure Volume", "open_ortho": "", "open_spine": "", "hip": "", "knee": ""}
    assert scoring.incision_likelihood(row, vol_cols) == scoring.INCISION_HIGH


def test_incision_likelihood_low():
    row = pd.Series({"Joint Replacement - Procedure Volume": "30"})
    vol_cols = {"joint_repl": "Joint Replacement - Procedure Volume", "open_ortho": "", "open_spine": "", "hip": "", "knee": ""}
    assert scoring.incision_likelihood(row, vol_cols) == scoring.INCISION_LOW


def test_best_approach_high_score_t1_triggers_in_person():
    out = scoring.best_approach(95, "Tier 1 (0-30 min)", "Yes")
    assert "In-Person" in out


def test_best_approach_low_score_is_email_nurture():
    out = scoring.best_approach(30, "Tier 5 (180+ min drivable)", "No")
    assert "Email" in out


def test_enrich_frame_populates_all_fields():
    df = pd.DataFrame([
        {
            "Practice Type": "Private Practice",
            "Tier": "Tier 1 (0-30 min)",
            "Microlyte Eligible": "Yes",
            "Joint Replacement - Procedure Volume": "600",
            "Knee Joint Replacement - Procedure Volume": "400",
            "Hip Joint Replacement - Procedure Volume": "150",
        }
    ])
    out = scoring.enrich_frame(df)
    row = out.iloc[0]
    assert row["Target Score"] > 0
    assert row["Target Tier"] in ("A+", "A", "B", "C", "D")
    assert row["Lead Priority"] in ("A", "B", "C", "D")
    assert row["Lg Incision Likelihood"] in (
        scoring.INCISION_HIGH, scoring.INCISION_MED_HIGH, scoring.INCISION_MED, scoring.INCISION_LOW,
    )
    assert isinstance(row["Why Target?"], str) and len(row["Why Target?"]) > 0
    assert isinstance(row["Best Approach"], str) and len(row["Best Approach"]) > 0


def test_routing_classifies_sources():
    assert routing._classify_source("joint_replacement_hcp_targeting_export.csv") == "JR"
    assert routing._classify_source("spine_surgeon_targets.csv") == "S&N"
    assert routing._classify_source("outisde_of_ortho_+_spine_hcp.csv") == "OOS"


def test_routing_prefers_jr_over_oos():
    line = routing._lead_primary_line("joint_replacement.csv;outisde_of_ortho.csv")
    assert line == "JR"


def test_routing_splits_by_line():
    df = pd.DataFrame([
        {"HCP NPI": "1", "__source_file": "joint_replacement_foo.csv"},
        {"HCP NPI": "2", "__source_file": "spine_bar.csv"},
        {"HCP NPI": "3", "__source_file": "outisde_of_ortho_baz.csv"},
    ])
    out = routing.enrich_frame(df)
    split = routing.split_by_product_line(out)
    assert set(split.keys()) == {"JR", "S&N", "OOS"}
    assert split["JR"].iloc[0]["HCP NPI"] == "1"
