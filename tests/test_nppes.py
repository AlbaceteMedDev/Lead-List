"""NPPES parsing, phone normalization, and reconciliation."""

import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent.parent))

from src import nppes


def test_normalize_phone_strips_nondigits():
    assert nppes.normalize_phone("(212) 555-1234") == "2125551234"


def test_normalize_phone_strips_country_code():
    assert nppes.normalize_phone("+1-212-555-1234") == "2125551234"


def test_normalize_phone_empty():
    assert nppes.normalize_phone(None) == ""
    assert nppes.normalize_phone("") == ""
    assert nppes.normalize_phone("bad") == ""


def test_parse_response_empty():
    out = nppes.parse_response({})
    assert out["nppes_found"] is False
    assert out["nppes_phone"] == ""


def test_parse_response_full():
    payload = {
        "results": [
            {
                "basic": {"credential": "MD", "status": "A"},
                "addresses": [
                    {"address_purpose": "MAILING", "telephone_number": "(000) 000-0000"},
                    {"address_purpose": "LOCATION", "telephone_number": "212-555-1234", "fax_number": "2125559999"},
                ],
                "taxonomies": [
                    {"primary": False, "desc": "Physician Assistant"},
                    {"primary": True, "desc": "Orthopaedic Surgery"},
                ],
            }
        ]
    }
    out = nppes.parse_response(payload)
    assert out["nppes_found"] is True
    assert out["nppes_phone"] == "2125551234"
    assert out["nppes_fax"] == "2125559999"
    assert out["nppes_credential"] == "MD"
    assert out["nppes_taxonomy"] == "Orthopaedic Surgery"
    assert out["nppes_status"] == "A"


def test_reconcile_phone_verified():
    v, s = nppes.reconcile_phone("212-555-1234", "2125551234", nppes_found=True)
    assert (v, s) == ("2125551234", nppes.STATUS_VERIFIED)


def test_reconcile_phone_added():
    v, s = nppes.reconcile_phone("", "2125551234", nppes_found=True)
    assert (v, s) == ("2125551234", nppes.STATUS_ADDED)


def test_reconcile_phone_updated():
    v, s = nppes.reconcile_phone("2125550000", "2125551234", nppes_found=True)
    assert (v, s) == ("2125551234", nppes.STATUS_UPDATED)


def test_reconcile_phone_missing():
    v, s = nppes.reconcile_phone("", "", nppes_found=False)
    assert (v, s) == ("", nppes.STATUS_MISSING)
