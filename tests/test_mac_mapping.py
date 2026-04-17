"""MAC jurisdiction and Microlyte eligibility tests, including Virginia carve-out."""

import json
import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent.parent))

from src import mac_mapping

ROOT = Path(__file__).resolve().parent.parent
CFG = mac_mapping.load_mac_config(ROOT / "config" / "mac_jurisdictions.json")


def test_new_york_ngs_eligible():
    out = mac_mapping.lookup("NY", CFG)
    assert out["MAC Jurisdiction"] == "NGS"
    assert out["Microlyte Eligible"] == "Yes"


def test_new_jersey_novitas_not_eligible():
    out = mac_mapping.lookup("NJ", CFG)
    assert out["MAC Jurisdiction"] == "Novitas"
    assert out["Microlyte Eligible"] == "No"


def test_pennsylvania_novitas_not_eligible():
    out = mac_mapping.lookup("PA", CFG)
    assert out["Microlyte Eligible"] == "No"


def test_virginia_default_palmetto_eligible():
    out = mac_mapping.lookup("VA", CFG, address1="123 Main St", city="Richmond")
    assert out["MAC Jurisdiction"] == "Palmetto"
    assert out["Microlyte Eligible"] == "Yes"


def test_virginia_arlington_carveout_to_novitas():
    out = mac_mapping.lookup("VA", CFG, address1="100 N Glebe Rd", city="Arlington")
    assert out["MAC Jurisdiction"] == "Novitas"
    assert out["Microlyte Eligible"] == "No"


def test_virginia_fairfax_carveout():
    out = mac_mapping.lookup("VA", CFG, address1="8260 Willow Oaks Corporate Dr", city="Fairfax")
    assert out["Microlyte Eligible"] == "No"


def test_virginia_alexandria_carveout():
    out = mac_mapping.lookup("VA", CFG, address1="1 Alexandria Way", city="Alexandria")
    assert out["Microlyte Eligible"] == "No"


def test_florida_first_coast_not_eligible():
    out = mac_mapping.lookup("FL", CFG)
    assert out["MAC Jurisdiction"] == "First Coast"
    assert out["Microlyte Eligible"] == "No"


def test_unknown_state():
    out = mac_mapping.lookup("XX", CFG)
    assert out["Microlyte Eligible"] == "Unknown"
