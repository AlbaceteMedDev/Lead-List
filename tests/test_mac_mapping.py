"""Tests for MAC jurisdiction mapping (src/mac_mapping.py)."""

import pandas as pd
import pytest

from src.mac_mapping import load_mac_config, map_mac_jurisdictions


class TestMacConfig:
    """Test MAC jurisdiction configuration."""

    def test_config_loads(self):
        config = load_mac_config()
        assert isinstance(config, dict)
        assert len(config) > 0

    def test_ngs_states_eligible(self):
        """NGS jurisdiction states (NY, CT, MA, etc.) should be Microlyte eligible."""
        config = load_mac_config()
        ngs_states = ["NY", "CT", "MA", "ME", "NH", "VT", "RI"]
        for state in ngs_states:
            assert state in config, f"Missing state: {state}"
            assert config[state]["mac"] == "NGS"
            assert config[state]["microlyte_eligible"] is True, f"{state} should be eligible"

    def test_novitas_lcd_states_not_eligible(self):
        """Novitas LCD states (NJ, PA, MD, DE, DC) should NOT be eligible."""
        config = load_mac_config()
        novitas_states = ["NJ", "PA", "MD", "DE", "DC"]
        for state in novitas_states:
            assert state in config, f"Missing state: {state}"
            assert config[state]["mac"] == "Novitas"
            assert config[state]["microlyte_eligible"] is False, f"{state} should NOT be eligible"
            assert config[state]["lcd_active"] is True

    def test_cgs_states_not_eligible(self):
        """CGS states (OH, KY) should NOT be eligible."""
        config = load_mac_config()
        for state in ["OH", "KY"]:
            assert config[state]["mac"] == "CGS"
            assert config[state]["microlyte_eligible"] is False

    def test_first_coast_states_not_eligible(self):
        """First Coast states (FL, PR, VI) should NOT be eligible."""
        config = load_mac_config()
        for state in ["FL", "PR", "VI"]:
            assert config[state]["mac"] == "First Coast"
            assert config[state]["microlyte_eligible"] is False

    def test_virginia_requires_county_review(self):
        """Virginia should be flagged for county-level review."""
        config = load_mac_config()
        assert "VA" in config
        assert config["VA"].get("requires_county_review") is True
        # VA default is Palmetto (eligible), but NoVA is Novitas (not eligible)
        assert config["VA"]["mac"] == "Palmetto GBA"
        assert config["VA"]["microlyte_eligible"] is True

    def test_palmetto_states_eligible(self):
        """Palmetto states (WV, NC, SC) should be eligible."""
        config = load_mac_config()
        for state in ["WV", "NC", "SC"]:
            assert config[state]["mac"] == "Palmetto GBA"
            assert config[state]["microlyte_eligible"] is True

    def test_wps_states_eligible(self):
        """WPS states (MI, WI, MN, IN) should be eligible."""
        config = load_mac_config()
        for state in ["MI", "WI", "MN", "IN"]:
            assert config[state]["mac"] == "WPS"
            assert config[state]["microlyte_eligible"] is True


class TestMapMacJurisdictions:
    """Test the map_mac_jurisdictions function."""

    def test_ny_gets_ngs_and_eligible(self):
        df = pd.DataFrame({"HCP NPI": ["111"], "State": ["NY"]})
        result = map_mac_jurisdictions(df)
        assert result.iloc[0]["MAC Jurisdiction"] == "NGS"
        assert result.iloc[0]["Microlyte Eligible"] == "Yes"

    def test_nj_gets_novitas_and_not_eligible(self):
        df = pd.DataFrame({"HCP NPI": ["222"], "State": ["NJ"]})
        result = map_mac_jurisdictions(df)
        assert result.iloc[0]["MAC Jurisdiction"] == "Novitas"
        assert result.iloc[0]["Microlyte Eligible"] == "No"

    def test_ct_eligible(self):
        df = pd.DataFrame({"HCP NPI": ["333"], "State": ["CT"]})
        result = map_mac_jurisdictions(df)
        assert result.iloc[0]["Microlyte Eligible"] == "Yes"

    def test_pa_not_eligible(self):
        df = pd.DataFrame({"HCP NPI": ["444"], "State": ["PA"]})
        result = map_mac_jurisdictions(df)
        assert result.iloc[0]["Microlyte Eligible"] == "No"

    def test_md_not_eligible(self):
        df = pd.DataFrame({"HCP NPI": ["555"], "State": ["MD"]})
        result = map_mac_jurisdictions(df)
        assert result.iloc[0]["Microlyte Eligible"] == "No"

    def test_va_flagged_for_review(self):
        """Virginia leads must be flagged for county-level review."""
        df = pd.DataFrame({"HCP NPI": ["666"], "State": ["VA"]})
        result = map_mac_jurisdictions(df)
        assert "REVIEW" in result.iloc[0]["VA Review Flag"]
        assert "county" in result.iloc[0]["VA Review Flag"].lower()
        # Default VA is eligible (Palmetto)
        assert result.iloc[0]["Microlyte Eligible"] == "Yes"

    def test_unknown_state(self):
        df = pd.DataFrame({"HCP NPI": ["777"], "State": ["XX"]})
        result = map_mac_jurisdictions(df)
        assert result.iloc[0]["MAC Jurisdiction"] == "Unknown"
        assert result.iloc[0]["Microlyte Eligible"] == "No"

    def test_missing_state(self):
        df = pd.DataFrame({"HCP NPI": ["888"], "State": [None]})
        result = map_mac_jurisdictions(df)
        assert result.iloc[0]["MAC Jurisdiction"] == "Unknown"
        assert result.iloc[0]["Microlyte Eligible"] == "No"

    def test_mixed_states(self):
        """Test multiple states in one DataFrame."""
        df = pd.DataFrame({
            "HCP NPI": ["1", "2", "3", "4"],
            "State": ["NY", "NJ", "CT", "PA"],
        })
        result = map_mac_jurisdictions(df)
        assert list(result["Microlyte Eligible"]) == ["Yes", "No", "Yes", "No"]
