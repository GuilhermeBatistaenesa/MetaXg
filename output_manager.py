import json
import os
import tempfile
from datetime import datetime

KIND_LOGS = "LOGS"
KIND_RELATORIOS = "RELATORIOS"
KIND_JSON = "JSON"
KIND_SCREENSHOTS = "SCREENSHOTS"


class OutputManager:
    def __init__(self, execution_id: str, object_name: str, public_base_dir: str, local_root: str, started_at: datetime = None):
        self.execution_id = execution_id
        self.object_name = object_name or "MetaX"
        self.public_base_dir = public_base_dir or ""
        self.local_root = local_root
        self.started_at = started_at or datetime.now()
        self.public_write_ok = True
        self.public_write_error = None

        if not self.public_base_dir:
            self._mark_public_error("PUBLIC_BASE_DIR nao configurado")

    def _mark_public_error(self, message: str):
        self.public_write_ok = False
        if not self.public_write_error:
            self.public_write_error = message
        print(f"[OUTPUT] Falha ao escrever na pasta publica: {message}")

    def _local_base_dir(self, kind: str) -> str:
        if kind == KIND_LOGS:
            return os.path.join(self.local_root, "logs")
        if kind == KIND_RELATORIOS:
            return os.path.join(self.local_root, "relatorios")
        if kind == KIND_JSON:
            return os.path.join(self.local_root, "json")
        if kind == KIND_SCREENSHOTS:
            return os.path.join(self.local_root, "logs", "screenshots")
        raise ValueError(f"Tipo de output invalido: {kind}")

    def _public_base_dir(self, kind: str) -> str:
        if kind == KIND_LOGS:
            return "07_LOGS"
        if kind == KIND_RELATORIOS:
            return "08_RELATORIOS"
        if kind == KIND_JSON:
            return "09_JSON"
        if kind == KIND_SCREENSHOTS:
            return "10_SCREENSHOTS"
        raise ValueError(f"Tipo de output invalido: {kind}")

    def _public_alias_base_dir(self, kind: str) -> str:
        if kind == KIND_LOGS:
            return "logs"
        if kind == KIND_RELATORIOS:
            return "relatorios"
        if kind == KIND_JSON:
            return "json"
        if kind == KIND_SCREENSHOTS:
            return os.path.join("logs", "screenshots")
        raise ValueError(f"Tipo de output invalido: {kind}")

    def _dated_dir(self, base_dir: str, when: datetime) -> str:
        year = when.strftime("%Y")
        month = when.strftime("%m")
        return os.path.join(base_dir, self.object_name, year, month)

    def _ensure_dir(self, path: str):
        os.makedirs(path, exist_ok=True)

    def _atomic_write_bytes(self, path: str, data: bytes):
        dir_name = os.path.dirname(path)
        self._ensure_dir(dir_name)
        fd, tmp_path = tempfile.mkstemp(prefix=".tmp_", dir=dir_name)
        try:
            with os.fdopen(fd, "wb") as f:
                f.write(data)
            os.replace(tmp_path, path)
        finally:
            try:
                if os.path.exists(tmp_path):
                    os.remove(tmp_path)
            except Exception:
                pass

    def _atomic_write_text(self, path: str, content: str):
        self._atomic_write_bytes(path, content.encode("utf-8"))

    def _write_public_text(self, kind: str, filename: str, content: str, when: datetime):
        if not self.public_base_dir:
            return
        self._write_public_text_for_base(self._public_base_dir, kind, filename, content, when, mark_error=True)
        self._write_public_text_for_base(self._public_alias_base_dir, kind, filename, content, when, mark_error=False)

    def _write_public_text_for_base(
        self,
        base_fn,
        kind: str,
        filename: str,
        content: str,
        when: datetime,
        mark_error: bool,
    ):
        try:
            base = base_fn(kind)
            dest_dir = self._dated_dir(os.path.join(self.public_base_dir, base), when)
            dest_path = os.path.join(dest_dir, filename)
            self._atomic_write_text(dest_path, content)
        except Exception as e:
            if mark_error:
                self._mark_public_error(str(e))
            else:
                print(f"[OUTPUT] Falha ao escrever alias publico: {e}")

    def _append_public_text(self, kind: str, filename: str, content: str, when: datetime):
        if not self.public_base_dir:
            return
        self._append_public_text_for_base(self._public_base_dir, kind, filename, content, when, mark_error=True)
        self._append_public_text_for_base(self._public_alias_base_dir, kind, filename, content, when, mark_error=False)

    def _append_public_text_for_base(
        self,
        base_fn,
        kind: str,
        filename: str,
        content: str,
        when: datetime,
        mark_error: bool,
    ):
        try:
            base = base_fn(kind)
            dest_dir = self._dated_dir(os.path.join(self.public_base_dir, base), when)
            self._ensure_dir(dest_dir)
            dest_path = os.path.join(dest_dir, filename)
            with open(dest_path, "a", encoding="utf-8") as f:
                f.write(content)
        except Exception as e:
            if mark_error:
                self._mark_public_error(str(e))
            else:
                print(f"[OUTPUT] Falha ao escrever alias publico: {e}")

    def _write_public_bytes(self, kind: str, filename: str, data: bytes, when: datetime):
        if not self.public_base_dir:
            return
        self._write_public_bytes_for_base(self._public_base_dir, kind, filename, data, when, mark_error=True)
        self._write_public_bytes_for_base(self._public_alias_base_dir, kind, filename, data, when, mark_error=False)

    def _write_public_bytes_for_base(
        self,
        base_fn,
        kind: str,
        filename: str,
        data: bytes,
        when: datetime,
        mark_error: bool,
    ):
        try:
            base = base_fn(kind)
            dest_dir = self._dated_dir(os.path.join(self.public_base_dir, base), when)
            dest_path = os.path.join(dest_dir, filename)
            self._atomic_write_bytes(dest_path, data)
        except Exception as e:
            if mark_error:
                self._mark_public_error(str(e))
            else:
                print(f"[OUTPUT] Falha ao escrever alias publico: {e}")

    def write_text(self, kind: str, filename: str, content: str, when: datetime = None) -> str:
        when = when or self.started_at
        local_dir = self._local_base_dir(kind)
        local_path = os.path.join(local_dir, filename)
        self._atomic_write_text(local_path, content)
        self._write_public_text(kind, filename, content, when)
        return local_path

    def append_text(self, kind: str, filename: str, content: str, when: datetime = None, write_public: bool = True) -> str:
        when = when or self.started_at
        local_dir = self._local_base_dir(kind)
        local_path = os.path.join(local_dir, filename)
        self._ensure_dir(local_dir)
        with open(local_path, "a", encoding="utf-8") as f:
            f.write(content)
        if write_public:
            self._append_public_text(kind, filename, content, when)
        return local_path

    def append_public_text_only(self, kind: str, filename: str, content: str, when: datetime = None):
        when = when or self.started_at
        self._append_public_text(kind, filename, content, when)

    def get_local_path(self, kind: str, filename: str) -> str:
        local_dir = self._local_base_dir(kind)
        return os.path.join(local_dir, filename)

    def write_json(self, kind: str, filename: str, data: dict, when: datetime = None) -> str:
        content = json.dumps(data, ensure_ascii=False, indent=2)
        return self.write_text(kind, filename, content, when=when)

    def save_screenshot_bytes(self, filename: str, data: bytes, when: datetime = None) -> str:
        when = when or self.started_at
        local_dir = self._local_base_dir(KIND_SCREENSHOTS)
        local_path = os.path.join(local_dir, filename)
        self._atomic_write_bytes(local_path, data)
        self._write_public_bytes(KIND_SCREENSHOTS, filename, data, when)
        return local_path
