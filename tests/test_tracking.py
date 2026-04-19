"""Activity persistence tests."""

import json
import sys
from pathlib import Path

import pandas as pd

sys.path.insert(0, str(Path(__file__).resolve().parent.parent))

from src import tracking


def test_merge_entry_keeps_latest_non_empty():
    existing = {"Lead Status": "New", "Call 1 Date": "2026-04-01", "Call 1 Outcome": "No Answer"}
    incoming = {"Lead Status": "Connected", "Call 1 Date": "", "Call 1 Outcome": "Voicemail"}
    out = tracking._merge_entry(existing, incoming)
    assert out["Lead Status"] == "Connected"
    assert out["Call 1 Date"] == "2026-04-01"
    assert out["Call 1 Outcome"] == "Voicemail"


def test_merge_entry_clears_roundless_outcome():
    existing = {"Call 1 Date": "", "Call 1 Outcome": "stale"}
    incoming = {}
    out = tracking._merge_entry(existing, incoming)
    assert "Call 1 Outcome" not in out


def test_apply_activity_joins_by_npi():
    df = pd.DataFrame([
        {"HCP NPI": "1234567890", "First Name": "Jane", "Last Name": "Smith"},
        {"HCP NPI": "0987654321", "First Name": "Bill", "Last Name": "Jones"},
    ])
    activity = {
        "1234567890": {"Lead Status": "Interested", "Call 1 Date": "2026-04-10", "Call 1 Outcome": "Connected - Interested"},
    }
    out = tracking.apply_activity(df, activity)
    jane = out[out["Last Name"] == "Smith"].iloc[0]
    bill = out[out["Last Name"] == "Jones"].iloc[0]
    assert jane["Lead Status"] == "Interested"
    assert jane["Call 1 Date"] == "2026-04-10"
    assert bill["Lead Status"] == "New"
    assert bill["Call 1 Date"] == ""


def test_load_save_cache_roundtrip(tmp_path):
    cache_path = tmp_path / "activity.json"
    data = {"1234567890": {"Lead Status": "Meeting Booked"}}
    tracking.save_cache(cache_path, data)
    assert tracking.load_cache(cache_path) == data


def test_summarize_last_touch():
    df = pd.DataFrame([
        {"HCP NPI": "1", "Call 1 Date": "2026-04-01", "Call 2 Date": "2026-04-10", "Email 1 Date": "2026-04-05"},
        {"HCP NPI": "2", "Call 1 Date": "", "Email 1 Date": ""},
    ])
    out = tracking.summarize_last_touch(df)
    assert out.iloc[0]["Last Touch Date"] == "2026-04-10"
    assert out.iloc[0]["Touch Count"] == 3
    assert out.iloc[1]["Last Touch Date"] == ""
    assert out.iloc[1]["Touch Count"] == 0
