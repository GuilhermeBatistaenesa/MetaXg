from datetime import datetime

import os
from notification import build_email_payload
from outcomes import OUTCOME_VERIFIED_SUCCESS, compute_totals


def test_notification_subject_and_attachments():
    started_at = datetime(2026, 1, 28, 10, 0, 0)
    run_context = {
        "execution_id": "exec-999",
        "object_name": "MetaX",
        "started_at": started_at.isoformat(),
        "finished_at": started_at.isoformat(),
        "duration_sec": 0,
        "run_status": "CONSISTENT",
        "report_path": None,
        "manifest_path": "C:\\temp\\manifest.json",
        "email_status": None,
        "email_error": None,
        "public_write_ok": True,
        "public_write_error": None,
        "environment": {"cwd": "C:\\repo"},
    }
    people = [{"nome": "Ana", "cpf": "123", "outcome": OUTCOME_VERIFIED_SUCCESS, "errors": {}, "no_photo": False}]
    totals = compute_totals(people, detected=1)
    manifest = {"run_context": run_context, "totals": totals, "people": people}

    try:
        with open("temp_manifest_final.json", "w", encoding="utf-8") as f:
            f.write("{}")

        subject, html_body, attachments = build_email_payload(
            manifest,
            report_path="C:\\temp\\report.txt",
            manifest_path="temp_manifest_final.json",
            partial_manifest_path="C:\\temp\\manifest_partial.json",
            attachment_paths=["C:\\temp\\report.txt", "temp_manifest_final.json"],
        )
    finally:
        if os.path.exists("temp_manifest_final.json"):
            os.remove("temp_manifest_final.json")

    assert "CONSISTENT" in subject
    assert "exec-999" in subject
    assert "2026-01-28" in subject
    assert "Execution ID" in html_body
    assert "Relatorio" in html_body
    assert "Manifest" in html_body
    assert "C:\\temp\\report.txt" in attachments
    assert "temp_manifest_final.json" in attachments
