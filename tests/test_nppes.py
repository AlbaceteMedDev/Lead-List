"""Tests for NPPES phone verification (src/nppes.py)."""

import json
import os
import tempfile
from unittest.mock import MagicMock, patch

import pandas as pd
import pytest

from src.nppes import (
    normalize_phone,
    load_cache,
    save_cache,
    is_cache_fresh,
    query_nppes,
    verify_phones,
)


class TestNormalizePhone:
    """Test phone number normalization to 10 digits."""

    def test_ten_digits(self):
        assert normalize_phone("2015551234") == "2015551234"

    def test_with_dashes(self):
        assert normalize_phone("201-555-1234") == "2015551234"

    def test_with_parens(self):
        assert normalize_phone("(201) 555-1234") == "2015551234"

    def test_with_country_code(self):
        assert normalize_phone("12015551234") == "2015551234"
        assert normalize_phone("+1-201-555-1234") == "2015551234"

    def test_with_dots(self):
        assert normalize_phone("201.555.1234") == "2015551234"

    def test_none(self):
        assert normalize_phone(None) is None

    def test_empty(self):
        assert normalize_phone("") is None

    def test_too_short(self):
        assert normalize_phone("12345") is None

    def test_too_long(self):
        assert normalize_phone("123456789012345") is None

    def test_numeric_input(self):
        assert normalize_phone(2015551234) == "2015551234"


class TestCache:
    """Test NPPES cache load/save and freshness."""

    def test_load_nonexistent(self):
        """Loading from a path that doesn't exist returns empty dict."""
        with patch("src.nppes.CACHE_FILE", "/tmp/nonexistent_cache.json"):
            result = load_cache()
            assert result == {}

    def test_save_and_load(self):
        with tempfile.TemporaryDirectory() as tmpdir:
            cache_file = os.path.join(tmpdir, "cache.json")
            with patch("src.nppes.CACHE_FILE", cache_file), \
                 patch("src.nppes.CACHE_DIR", tmpdir):
                data = {"1234567890": {"phone": "2015551234", "cached_at": "2026-04-01T00:00:00"}}
                save_cache(data)
                loaded = load_cache()
                assert loaded == data

    def test_cache_fresh(self):
        """Cache entries less than 30 days old are fresh."""
        from datetime import datetime
        entry = {"cached_at": datetime.now().isoformat()}
        assert is_cache_fresh(entry) is True

    def test_cache_stale(self):
        """Cache entries more than 30 days old are stale."""
        entry = {"cached_at": "2025-01-01T00:00:00"}
        assert is_cache_fresh(entry) is False

    def test_cache_no_timestamp(self):
        """Cache entries without timestamps are stale."""
        assert is_cache_fresh({}) is False


class TestQueryNppes:
    """Test NPPES API parsing (mocked)."""

    def _mock_response(self):
        """Create a mock NPPES API response."""
        return {
            "result_count": 1,
            "results": [{
                "basic": {
                    "credential": "M.D.",
                    "status": "A",
                },
                "addresses": [
                    {
                        "address_purpose": "MAILING",
                        "telephone_number": "111-111-1111",
                    },
                    {
                        "address_purpose": "LOCATION",
                        "telephone_number": "201-555-9999",
                        "fax_number": "201-555-8888",
                    },
                ],
                "taxonomies": [
                    {"primary": False, "desc": "Internal Medicine"},
                    {"primary": True, "desc": "Orthopaedic Surgery"},
                ],
            }],
        }

    @patch("src.nppes.requests.get")
    def test_parses_location_phone(self, mock_get):
        mock_resp = MagicMock()
        mock_resp.json.return_value = self._mock_response()
        mock_resp.raise_for_status = MagicMock()
        mock_get.return_value = mock_resp

        result = query_nppes("1234567890")
        assert result is not None
        assert result["phone"] == "201-555-9999"
        assert result["fax"] == "201-555-8888"

    @patch("src.nppes.requests.get")
    def test_parses_credential(self, mock_get):
        mock_resp = MagicMock()
        mock_resp.json.return_value = self._mock_response()
        mock_resp.raise_for_status = MagicMock()
        mock_get.return_value = mock_resp

        result = query_nppes("1234567890")
        assert result["credential"] == "M.D."

    @patch("src.nppes.requests.get")
    def test_parses_primary_taxonomy(self, mock_get):
        mock_resp = MagicMock()
        mock_resp.json.return_value = self._mock_response()
        mock_resp.raise_for_status = MagicMock()
        mock_get.return_value = mock_resp

        result = query_nppes("1234567890")
        assert result["taxonomy"] == "Orthopaedic Surgery"

    @patch("src.nppes.requests.get")
    def test_no_results(self, mock_get):
        mock_resp = MagicMock()
        mock_resp.json.return_value = {"result_count": 0, "results": []}
        mock_resp.raise_for_status = MagicMock()
        mock_get.return_value = mock_resp

        result = query_nppes("0000000000")
        assert result is None

    @patch("src.nppes.requests.get")
    def test_api_error_returns_none(self, mock_get):
        mock_get.side_effect = Exception("Connection refused")
        result = query_nppes("1234567890")
        assert result is None


class TestVerifyPhones:
    """Test phone status logic."""

    @patch("src.nppes.query_nppes")
    @patch("src.nppes.time.sleep")
    def test_phone_verified(self, mock_sleep, mock_query):
        """Original phone matches NPPES phone → Verified."""
        mock_query.return_value = {"npi": "111", "phone": "201-555-1234", "credential": "MD", "fax": None}
        df = pd.DataFrame({
            "HCP NPI": ["111"],
            "Phone Number": ["2015551234"],
            "Credential": ["MD"],
        })
        result = verify_phones(df, force=True)
        assert result.iloc[0]["Phone Status"] == "Verified"
        assert result.iloc[0]["Verified Phone"] == "2015551234"

    @patch("src.nppes.query_nppes")
    @patch("src.nppes.time.sleep")
    def test_phone_added_from_nppes(self, mock_sleep, mock_query):
        """No original phone, NPPES has one → Added from NPPES."""
        mock_query.return_value = {"npi": "222", "phone": "201-555-5678", "credential": "DO", "fax": None}
        df = pd.DataFrame({
            "HCP NPI": ["222"],
            "Phone Number": [None],
            "Credential": ["DO"],
        })
        result = verify_phones(df, force=True)
        assert result.iloc[0]["Phone Status"] == "Added from NPPES"
        assert result.iloc[0]["Verified Phone"] == "2015555678"

    @patch("src.nppes.query_nppes")
    @patch("src.nppes.time.sleep")
    def test_phone_updated(self, mock_sleep, mock_query):
        """Original differs from NPPES → Updated (NPPES differs)."""
        mock_query.return_value = {"npi": "333", "phone": "201-555-9999", "credential": "MD", "fax": None}
        df = pd.DataFrame({
            "HCP NPI": ["333"],
            "Phone Number": ["2015551111"],
            "Credential": ["MD"],
        })
        result = verify_phones(df, force=True)
        assert result.iloc[0]["Phone Status"] == "Updated (NPPES differs)"
        assert result.iloc[0]["Verified Phone"] == "2015559999"

    @patch("src.nppes.query_nppes")
    @patch("src.nppes.time.sleep")
    def test_phone_missing(self, mock_sleep, mock_query):
        """No phone on file, NPI not in NPPES → Missing."""
        mock_query.return_value = None
        df = pd.DataFrame({
            "HCP NPI": ["444"],
            "Phone Number": [None],
            "Credential": [None],
        })
        result = verify_phones(df, force=True)
        assert result.iloc[0]["Phone Status"] == "Missing"
