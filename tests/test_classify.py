"""Tests for practice classification (src/classify.py)."""

import pandas as pd
import pytest

from src.classify import classify_practices, is_hospital_based, load_keywords

KEYWORDS = load_keywords()


class TestIsHospitalBased:
    """Test the is_hospital_based function against known edge cases."""

    def test_university_orthopedics_center_is_private(self):
        """Critical edge case: 'University Orthopedics Center' is a PRIVATE practice."""
        assert is_hospital_based("University Orthopedics Center", KEYWORDS) is False

    def test_university_orthopaedic_associates_is_private(self):
        """University Orthopaedic Associates is also private."""
        assert is_hospital_based("University Orthopaedic Associates", KEYWORDS) is False

    def test_university_of_is_hospital(self):
        """'University of X Medical Center' should be hospital."""
        assert is_hospital_based("University of Pennsylvania Medical Center", KEYWORDS) is True

    def test_university_hospital_is_hospital(self):
        assert is_hospital_based("University Hospital Newark", KEYWORDS) is True

    def test_named_systems(self):
        """Test all major named health systems are classified as hospital."""
        hospital_names = [
            "Atlantic Health System - Morristown",
            "Hackensack Meridian Health - JFK",
            "Northwell Health Ortho",
            "NYU Langone Orthopedic Center",
            "Hospital for Special Surgery",
            "Mount Sinai West",
            "Robert Wood Johnson University Hospital",
            "UPMC Shadyside",
            "Yale New Haven Hospital",
            "Hartford HealthCare Bone & Joint",
            "Stamford Health",
            "Greenwich Hospital",
            "Geisinger Medical Center",
        ]
        for name in hospital_names:
            assert is_hospital_based(name, KEYWORDS) is True, f"Expected hospital: {name}"

    def test_generic_patterns(self):
        """Test generic hospital patterns."""
        assert is_hospital_based("Valley Hospital", KEYWORDS) is True
        assert is_hospital_based("Springfield Medical Center", KEYWORDS) is True
        assert is_hospital_based("Regional Health System", KEYWORDS) is True
        assert is_hospital_based("VA Medical Center", KEYWORDS) is True
        assert is_hospital_based("Veterans Affairs Clinic", KEYWORDS) is True

    def test_private_practices(self):
        """Known private practices should NOT be classified as hospital."""
        private_names = [
            "Garden State Orthopaedic Associates",
            "Summit Medical Group",
            "Orlin Cohen Orthopedic Group",
            "Advanced Orthopedics & Sports Medicine",
            "Rothman Orthopaedics",
            "Princeton Orthopaedic Associates",
            "Tri-County Orthopedics",
            "OrthoNJ",
        ]
        for name in private_names:
            assert is_hospital_based(name, KEYWORDS) is False, f"Expected private: {name}"

    def test_none_and_empty(self):
        """None and empty strings should default to private."""
        assert is_hospital_based(None, KEYWORDS) is False
        assert is_hospital_based("", KEYWORDS) is False

    def test_case_insensitive(self):
        """Matching should be case-insensitive."""
        assert is_hospital_based("ATLANTIC HEALTH SYSTEM", KEYWORDS) is True
        assert is_hospital_based("atlantic health system", KEYWORDS) is True


class TestClassifyPractices:
    """Test the full classify_practices function."""

    def test_adds_practice_type_column(self):
        df = pd.DataFrame({
            "HCP NPI": ["1234567890", "0987654321"],
            "Primary Site of Care": [
                "Garden State Orthopaedic Associates",
                "NYU Langone Health",
            ],
        })
        result = classify_practices(df)
        assert "Practice Type" in result.columns
        assert result.iloc[0]["Practice Type"] == "Private Practice"
        assert result.iloc[1]["Practice Type"] == "Hospital-Based"

    def test_default_to_private_practice(self):
        """Unrecognized names default to Private Practice."""
        df = pd.DataFrame({
            "HCP NPI": ["1111111111"],
            "Primary Site of Care": ["Some Random Clinic Name"],
        })
        result = classify_practices(df)
        assert result.iloc[0]["Practice Type"] == "Private Practice"

    def test_does_not_modify_original(self):
        """classify_practices should return a copy."""
        df = pd.DataFrame({
            "HCP NPI": ["1234567890"],
            "Primary Site of Care": ["Test Clinic"],
        })
        result = classify_practices(df)
        assert "Practice Type" not in df.columns
        assert "Practice Type" in result.columns
