"""Email classification and inference tests."""

import sys
from pathlib import Path

import pandas as pd

sys.path.insert(0, str(Path(__file__).resolve().parent.parent))

from src import email_enrich


def test_missing_email():
    assert email_enrich.classify_email("", "Jane", "Smith") == email_enrich.STATUS_MISSING
    assert email_enrich.classify_email(None, "Jane", "Smith") == email_enrich.STATUS_MISSING


def test_generic_office():
    assert email_enrich.classify_email("info@orthonj.com", "Jane", "Smith") == email_enrich.STATUS_GENERIC
    assert email_enrich.classify_email("office@orthonj.com", "Jane", "Smith") == email_enrich.STATUS_GENERIC


def test_hospital_domain():
    assert email_enrich.classify_email("jsmith@hss.edu", "Jane", "Smith") == email_enrich.STATUS_HOSPITAL


def test_verified_practice_domain():
    status = email_enrich.classify_email("jsmith@orthonj.com", "Jane", "Smith")
    assert status == email_enrich.STATUS_VERIFIED


def test_personal_with_name():
    status = email_enrich.classify_email("jane.smith@gmail.com", "Jane", "Smith")
    assert status == email_enrich.STATUS_PERSONAL_NAME


def test_personal_no_name():
    status = email_enrich.classify_email("foobar@gmail.com", "Jane", "Smith")
    assert status == email_enrich.STATUS_PERSONAL_NO_NAME


def test_practice_no_name_match_flagged():
    status = email_enrich.classify_email("reception-desk-4@orthonj.com", "Jane", "Smith")
    assert status == email_enrich.STATUS_PRACTICE_REVIEW


def test_inference_from_practice_pattern():
    df = pd.DataFrame([
        {"First Name": "Jane", "Last Name": "Smith", "Email": "jsmith@orthonj.com", "Primary Site of Care": "OrthoNJ"},
        {"First Name": "Bill", "Last Name": "Jones", "Email": "bjones@orthonj.com", "Primary Site of Care": "OrthoNJ"},
        {"First Name": "Mary", "Last Name": "Kay", "Email": "", "Primary Site of Care": "OrthoNJ"},
    ])
    out = email_enrich.enrich_frame(df)
    mk = out[out["Last Name"] == "Kay"].iloc[0]
    assert mk["Email"] == "mkay@orthonj.com"
    assert mk["Email Status"] == email_enrich.STATUS_INFERRED


def test_inference_never_uses_free_domain():
    df = pd.DataFrame([
        {"First Name": "Jane", "Last Name": "Smith", "Email": "jane.smith@gmail.com", "Primary Site of Care": "OrthoNJ"},
        {"First Name": "Mary", "Last Name": "Kay", "Email": "", "Primary Site of Care": "OrthoNJ"},
    ])
    out = email_enrich.enrich_frame(df)
    mk = out[out["Last Name"] == "Kay"].iloc[0]
    assert mk["Email"] == ""
    assert mk["Email Status"] == email_enrich.STATUS_MISSING
