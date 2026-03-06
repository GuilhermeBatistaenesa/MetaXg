from datetime import datetime
import os
import tempfile
from notification import build_email_payload
from outcomes import OUTCOME_VERIFIED_SUCCESS, compute_totals


def test_notification_subject_and_attachments():
    started_at = datetime(2026, 1, 28, 10, 0, 0)
    temp_dir = tempfile.mkdtemp()
    report_path = os.path.join(temp_dir, "report.txt")
    manifest_path = os.path.join(temp_dir, "manifest.json")
    partial_manifest_path = os.path.join(temp_dir, "manifest_partial.json")
    run_context = {
        "execution_id": "exec-999",
        "object_name": "MetaX",
        "started_at": started_at.isoformat(),
        "finished_at": started_at.isoformat(),
        "duration_sec": 0,
        "run_status": "CONSISTENT",
        "report_path": None,
        "manifest_path": manifest_path,
        "email_status": None,
        "email_error": None,
        "public_write_ok": True,
        "public_write_error": None,
        "environment": {"cwd": temp_dir},
    }
    people = [{"nome": "Ana", "cpf": "123", "outcome": OUTCOME_VERIFIED_SUCCESS, "errors": {}, "no_photo": False}]
    totals = compute_totals(people, detected=1)
    manifest = {"run_context": run_context, "totals": totals, "people": people}

    try:
        with open(os.path.join(temp_dir, "temp_manifest_final.json"), "w", encoding="utf-8") as f:
            f.write("{}")

        subject, html_body, attachments = build_email_payload(
            manifest,
            report_path=report_path,
            manifest_path=os.path.join(temp_dir, "temp_manifest_final.json"),
            partial_manifest_path=partial_manifest_path,
            attachment_paths=[report_path, os.path.join(temp_dir, "temp_manifest_final.json")],
        )
    finally:
        manifest_final = os.path.join(temp_dir, "temp_manifest_final.json")
        if os.path.exists(manifest_final):
            os.remove(manifest_final)

    assert "CONSISTENT" in subject
    assert "exec-999" in subject
    assert "2026-01-28" in subject
    assert "Execution ID" in html_body
    assert "Relatorio" in html_body
    assert "Manifest" in html_body
    assert report_path in attachments
    assert os.path.join(temp_dir, "temp_manifest_final.json") in attachments
