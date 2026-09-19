"""Regression: the trailing 'SUBMITTED: [ ]' line is editorial state, not body.

The rewrite capture used to run to the end of the block, so every rewrite
carried '\n\nSUBMITTED: [ ]' and failed the 7-sentence validator (8 sentences
plus a blank line), and rewrite == current could never be true. The published
'457 rewrites, 0 valid' counted 131 blank rewrites and 102 unchanged copies as
rewrites, and validated none.
"""
import importlib.util
import os
import sys
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
SCRIPT = ROOT / "scripts" / "build_submission_cockpit.py"


def _load():
    spec = importlib.util.spec_from_file_location("build_submission_cockpit", SCRIPT)
    mod = importlib.util.module_from_spec(spec)
    spec.loader.exec_module(mod)
    return mod


SEVEN = (
    "Does X beat Y in adults with Z? We used 12 trials with 4,321 participants from three registries. "
    "We pooled risk ratios with a random-effects model. The pooled risk ratio was 0.80 (95% CI 0.70 to 0.91). "
    "Leave-one-out analysis kept the estimate between 0.78 and 0.83. This suggests a modest benefit in this population. "
    "The evidence cannot address long-term harms because follow-up was short."
)


def _block(index, current, rewrite, submitted=" "):
    return (
        "=" * 60 + "\n"
        f"[{index}/469] Sample Project\n"
        "PATH: C:\\Projects\\sample\\\n"
        "TITLE: Sample\n"
        "TYPE: MA | ESTIMAND: RR\n"
        "DATA: 12 trials\n"
        f"CURRENT BODY ({len(current.split())} words):\n{current}\n"
        f"YOUR REWRITE (at most 156 words, 7 sentences):\n{rewrite}\n\n"
        f"SUBMITTED: [{submitted}]\n"
    )


def test_submitted_line_is_not_part_of_the_body(tmp_path):
    mod = _load()
    wb = tmp_path / "wb.txt"
    wb.write_text(_block(1, "Old body.", SEVEN), encoding="utf-8")
    (entry,) = mod.parse_external_rewrite_workbook(wb)
    assert entry["rewrite"] == SEVEN
    assert "SUBMITTED" not in entry["rewrite"]
    assert entry["rewritePresent"] is True
    assert entry["rewriteValid"] is True, entry["rewriteIssues"]
    assert entry["rewriteSentenceCount"] == 7
    assert entry["submitted"] is False


def test_blank_rewrite_is_not_present(tmp_path):
    mod = _load()
    wb = tmp_path / "wb.txt"
    wb.write_text(_block(1, "Old body.", ""), encoding="utf-8")
    (entry,) = mod.parse_external_rewrite_workbook(wb)
    assert entry["rewrite"] == ""
    assert entry["rewritePresent"] is False
    assert entry["rewriteValid"] is False


def test_unchanged_copy_is_matches_current_not_present(tmp_path):
    mod = _load()
    wb = tmp_path / "wb.txt"
    wb.write_text(_block(1, SEVEN, SEVEN), encoding="utf-8")
    (entry,) = mod.parse_external_rewrite_workbook(wb)
    assert entry["rewriteMatchesCurrent"] is True
    assert entry["rewritePresent"] is False


def test_submitted_flag_is_read(tmp_path):
    mod = _load()
    wb = tmp_path / "wb.txt"
    wb.write_text(_block(1, "Old body.", SEVEN, submitted="x"), encoding="utf-8")
    (entry,) = mod.parse_external_rewrite_workbook(wb)
    assert entry["submitted"] is True
    assert entry["rewriteValid"] is True


def test_bundled_snapshot_validates_a_nonzero_number_of_rewrites():
    """The validator must be able to pass on real data. 0 of N is the tell."""
    mod = _load()
    entries = mod.parse_external_rewrite_workbook(ROOT / "data-source" / "rewrite-workbook.snapshot.txt")
    assert len(entries) == 469
    present = [e for e in entries if e["rewritePresent"]]
    valid = [e for e in present if e["rewriteValid"]]
    assert len(present) == 236
    assert sum(e["rewriteMatchesCurrent"] for e in entries) == 102
    assert len(valid) == 199
    assert sum(e["submitted"] for e in entries) == 0
