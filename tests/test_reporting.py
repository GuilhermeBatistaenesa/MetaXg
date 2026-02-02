import tempfile
from datetime import datetime

from output_manager import OutputManager, KIND_RELATORIOS
from reporting import gerar_relatorio_txt
from outcomes import (
    OUTCOME_FAILED_ACTION,
    OUTCOME_VERIFIED_SUCCESS,
    compute_totals,
)


def test_report_contains_sections_and_execution_id():
    started_at = datetime(2026, 1, 28, 9, 0, 0)
    run_context = {
        "execution_id": "exec-123",
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
    people = [
        {"nome": "Ana", "cpf": "123", "outcome": OUTCOME_FAILED_ACTION, "errors": {"action_error": "x"}, "no_photo": False},
        {"nome": "Bia", "cpf": "456", "outcome": OUTCOME_VERIFIED_SUCCESS, "errors": {}, "no_photo": False},
    ]
    totals = compute_totals(people, detected=2)
    manifest = {"run_context": run_context, "totals": totals, "people": people}

    with tempfile.TemporaryDirectory() as tmpdir:
        om = OutputManager(
            execution_id="exec-123",
            object_name="MetaX",
            public_base_dir="",
            local_root=tmpdir,
            started_at=started_at,
        )
        path = gerar_relatorio_txt(manifest, om)
        assert path
        with open(path, "r", encoding="utf-8") as f:
            content = f.read()

    assert "Execution ID: exec-123" in content
    assert "LISTA DE FALHAS NA ACAO" in content
    assert "LISTA DE SUCESSO (VERIFICADO)" in content
    assert "Report Path:" in content
    assert "Manifest Path:" in content
