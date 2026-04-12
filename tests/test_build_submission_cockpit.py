import importlib.util
import os
import time
from pathlib import Path


MODULE_PATH = Path(__file__).resolve().parents[1] / "scripts" / "build_submission_cockpit.py"
SPEC = importlib.util.spec_from_file_location("build_submission_cockpit", MODULE_PATH)
MODULE = importlib.util.module_from_spec(SPEC)
assert SPEC and SPEC.loader
SPEC.loader.exec_module(MODULE)


def make_tracker_row():
    row = {column: "" for column in MODULE.TRACKER_COLUMNS}
    row.update(
        {
            "project_id": "demo-1",
            "project_name": "Demo Project",
            "slug": "demo-project",
            "project_path": r"C:\Demo\Project",
            "tier": "Tier 1",
            "resolved_status": "ACTIVE",
            "target_journal": "BMJ",
            "portfolio_priority": "high",
            "ops_readiness_score": "80",
            "citation_readiness_score": "70",
            "workflow_readiness_score": "60",
            "recommended_license": "MIT",
            "source_paper_signal": "yes",
            "source_protocol_signal": "no",
            "existing_pages_url": "",
            "deposit_state": "Not indexed",
            "rewrite_workbook_source": r"C:\E156\rewrite-workbook.txt",
            "rewrite_present": "yes",
            "rewrite_valid": "yes",
            "rewrite_word_count": "150",
            "rewrite_sentence_count": "7",
            "rewrite_matches_current": "yes",
            "rewrite_status": "rewritten",
            "publish_now": "yes",
            "e156_body_status": "ready",
            "protocol_status": "ready",
            "dashboard_status": "ready",
            "code_release_status": "ready",
            "mit_license_status": "ready",
            "github_status": "missing",
            "pages_status": "missing",
            "pdf_status": "pending",
            "synthesis_status": "not-configured",
            "source_article_path": r"C:\Demo\paper.md",
            "github_repo_url": "",
            "github_pages_url": "",
            "paper_pdf_path": "",
            "synthesis_article_url": "",
            "notes": "",
            "last_updated": "2026-04-12T00:00:00+00:00",
        }
    )
    return row


def empty_inputs():
    return (
        {"portfolio": []},
        {"projects": []},
        {"projects": []},
        {"projects": []},
        {"items": []},
    )


def test_build_records_omits_per_record_generated_at():
    records = MODULE.build_records([make_tracker_row()], *empty_inputs())
    assert len(records) == 1
    assert "generatedAt" not in records[0]


def test_generated_at_from_paths_uses_latest_mtime(tmp_path):
    older = tmp_path / "older.txt"
    newer = tmp_path / "newer.txt"
    older.write_text("a", encoding="utf-8")
    newer.write_text("b", encoding="utf-8")

    now = int(time.time())
    os.utime(older, (now - 10, now - 10))
    os.utime(newer, (now, now))

    assert MODULE.generated_at_from_paths([older, newer]) == MODULE.isoformat_utc_timestamp(float(now))


def test_write_json_uses_lf_newlines(tmp_path):
    output = tmp_path / "payload.json"
    MODULE.write_json(output, {"alpha": 1})
    data = output.read_bytes()
    assert b"\r\n" not in data
    assert data.endswith(b"\n")


def test_windows_to_local_path_prefers_existing_native_path():
    if os.name != "nt":
        return
    native_path = str(MODULE_PATH).replace("/", "\\")
    resolved = MODULE.windows_to_local_path(native_path)
    assert resolved == Path(native_path)
