"""Tests for drive-time tiering (src/tier.py)."""

import pandas as pd
import pytest

from src.tier import assign_tiers, haversine_miles, miles_to_tier, DEFAULT_LAT, DEFAULT_LON


class TestHaversine:
    """Test haversine distance calculation."""

    def test_same_point(self):
        assert haversine_miles(40.0, -74.0, 40.0, -74.0) == 0.0

    def test_nyc_to_newark(self):
        """NYC Midtown to Newark, NJ — roughly 10 miles."""
        dist = haversine_miles(DEFAULT_LAT, DEFAULT_LON, 40.7357, -74.1724)
        assert 8 < dist < 15

    def test_nyc_to_philly(self):
        """NYC to Philadelphia — roughly 80 miles straight-line."""
        dist = haversine_miles(DEFAULT_LAT, DEFAULT_LON, 39.9526, -75.1652)
        assert 70 < dist < 100

    def test_nyc_to_chicago(self):
        """NYC to Chicago — roughly 713 miles."""
        dist = haversine_miles(DEFAULT_LAT, DEFAULT_LON, 41.8781, -87.6298)
        assert 700 < dist < 750


class TestMilesToTier:
    """Test mile-to-tier conversion at boundaries."""

    def test_tier_1_boundaries(self):
        assert miles_to_tier(0) == "Tier 1 (0-30 min)"
        assert miles_to_tier(10) == "Tier 1 (0-30 min)"
        assert miles_to_tier(20) == "Tier 1 (0-30 min)"

    def test_tier_2_boundaries(self):
        assert miles_to_tier(20.1) == "Tier 2 (30-60 min)"
        assert miles_to_tier(45) == "Tier 2 (30-60 min)"

    def test_tier_3_boundaries(self):
        assert miles_to_tier(45.1) == "Tier 3 (60-120 min)"
        assert miles_to_tier(90) == "Tier 3 (60-120 min)"

    def test_tier_4_boundaries(self):
        assert miles_to_tier(90.1) == "Tier 4 (120-180 min)"
        assert miles_to_tier(140) == "Tier 4 (120-180 min)"

    def test_tier_5_boundaries(self):
        assert miles_to_tier(140.1) == "Tier 5 (180+ drivable)"
        assert miles_to_tier(350) == "Tier 5 (180+ drivable)"

    def test_tier_6_boundaries(self):
        assert miles_to_tier(350.1) == "Tier 6 (Requires flight)"
        assert miles_to_tier(1000) == "Tier 6 (Requires flight)"


class TestAssignTiers:
    """Test full tier assignment on a DataFrame."""

    def test_hospital_based_gets_hospital_tier(self):
        """Hospital-Based leads always get Hospital-Based tier regardless of location."""
        df = pd.DataFrame({
            "HCP NPI": ["1234567890"],
            "Practice Type": ["Hospital-Based"],
            "Postal Code": ["07410"],  # NJ zip, would be Tier 1
        })
        result = assign_tiers(df)
        assert result.iloc[0]["Tier"] == "Hospital-Based"

    def test_private_practice_nyc_zip_tier_1(self):
        """NYC-area zip code should be Tier 1."""
        df = pd.DataFrame({
            "HCP NPI": ["1234567890"],
            "Practice Type": ["Private Practice"],
            "Postal Code": ["10001"],  # Manhattan
        })
        result = assign_tiers(df)
        assert result.iloc[0]["Tier"] == "Tier 1 (0-30 min)"

    def test_missing_zip_defaults_to_tier_5(self):
        """Missing postal code defaults to Tier 5."""
        df = pd.DataFrame({
            "HCP NPI": ["1234567890"],
            "Practice Type": ["Private Practice"],
            "Postal Code": [None],
        })
        result = assign_tiers(df)
        assert result.iloc[0]["Tier"] == "Tier 5 (180+ drivable)"

    def test_custom_origin(self):
        """Tier changes based on origin point."""
        df = pd.DataFrame({
            "HCP NPI": ["1234567890"],
            "Practice Type": ["Private Practice"],
            "Postal Code": ["90210"],  # Beverly Hills, CA
        })
        # From NYC — should be Tier 6 (flight)
        result_nyc = assign_tiers(df)
        assert "Tier 6" in result_nyc.iloc[0]["Tier"] or "Tier 5" in result_nyc.iloc[0]["Tier"]

        # From LA — should be Tier 1 or 2
        result_la = assign_tiers(df, origin_lat=34.0522, origin_lon=-118.2437)
        assert "Tier 1" in result_la.iloc[0]["Tier"] or "Tier 2" in result_la.iloc[0]["Tier"]
