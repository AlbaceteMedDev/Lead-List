"""Practice classification edge cases."""

import json
import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent.parent))

from src import classify

ROOT = Path(__file__).resolve().parent.parent
KEYWORDS = json.loads((ROOT / "config" / "hospital_keywords.json").read_text())


def test_university_orthopedics_center_is_private():
    assert classify.classify_site("University Orthopedics Center", KEYWORDS) == classify.PRIVATE


def test_hackensack_meridian_is_hospital():
    assert classify.classify_site("Hackensack Meridian Health", KEYWORDS) == classify.HOSPITAL


def test_hospital_for_special_surgery_is_hospital():
    assert classify.classify_site("Hospital for Special Surgery", KEYWORDS) == classify.HOSPITAL


def test_plain_hospital_is_hospital():
    assert classify.classify_site("Morristown Memorial Hospital", KEYWORDS) == classify.HOSPITAL


def test_private_practice_default():
    assert classify.classify_site("North Jersey Orthopedic Associates", KEYWORDS) == classify.PRIVATE


def test_empty_defaults_private():
    assert classify.classify_site("", KEYWORDS) == classify.PRIVATE
    assert classify.classify_site(None, KEYWORDS) == classify.PRIVATE


def test_va_medical_center_is_hospital():
    assert classify.classify_site("VA Medical Center - East Orange", KEYWORDS) == classify.HOSPITAL


def test_university_of_pennsylvania_is_hospital():
    assert classify.classify_site("University of Pennsylvania Hospital", KEYWORDS) == classify.HOSPITAL
