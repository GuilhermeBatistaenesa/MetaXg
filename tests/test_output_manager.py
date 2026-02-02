import tempfile
from datetime import datetime

from output_manager import OutputManager, KIND_LOGS


def test_append_text_does_not_duplicate():
    with tempfile.TemporaryDirectory() as tmpdir:
        om = OutputManager(
            execution_id="exec-1",
            object_name="MetaX",
            public_base_dir="",
            local_root=tmpdir,
            started_at=datetime(2026, 1, 28, 10, 0, 0),
        )
        filename = "test.log"
        om.append_text(KIND_LOGS, filename, "linha1\n", write_public=False)
        om.append_text(KIND_LOGS, filename, "linha2\n", write_public=False)

        path = om.get_local_path(KIND_LOGS, filename)
        with open(path, "r", encoding="utf-8") as f:
            content = f.read()

    assert content == "linha1\nlinha2\n"
