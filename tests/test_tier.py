"""Drive-time tier boundary tests."""

import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent.parent))

from src import tier


def test_haversine_zero():
    assert tier.haversine_miles(40.0, -74.0, 40.0, -74.0) == 0.0


def test_haversine_known_distance():
    nyc = (40.7580, -73.9855)
    philly = (39.9526, -75.1652)
    miles = tier.haversine_miles(*nyc, *philly)
    assert 75 < miles < 90


def test_tier_boundaries():
    assert tier.miles_to_tier(0) == "Tier 1 (0-30 min)"
    assert tier.miles_to_tier(19.9) == "Tier 1 (0-30 min)"
    assert tier.miles_to_tier(20) == "Tier 2 (30-60 min)"
    assert tier.miles_to_tier(44.9) == "Tier 2 (30-60 min)"
    assert tier.miles_to_tier(45) == "Tier 3 (60-120 min)"
    assert tier.miles_to_tier(89.9) == "Tier 3 (60-120 min)"
    assert tier.miles_to_tier(90) == "Tier 4 (120-180 min)"
    assert tier.miles_to_tier(139.9) == "Tier 4 (120-180 min)"
    assert tier.miles_to_tier(140) == "Tier 5 (180+ min drivable)"
    assert tier.miles_to_tier(349.9) == "Tier 5 (180+ min drivable)"
    assert tier.miles_to_tier(350) == "Tier 6 (Requires flight)"
    assert tier.miles_to_tier(5000) == "Tier 6 (Requires flight)"


def test_tier_unknown_defaults_to_five():
    assert tier.miles_to_tier(None) == "Tier 5 (180+ min drivable)"
    assert tier.miles_to_tier(float("nan")) == "Tier 5 (180+ min drivable)"
