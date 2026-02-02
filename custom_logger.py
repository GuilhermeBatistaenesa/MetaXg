import json
import time
from datetime import datetime
from typing import Optional

from output_manager import KIND_LOGS, OutputManager

class CustomLogger:
    _instance = None

    def __new__(cls, *args, **kwargs):
        if not cls._instance:
            cls._instance = super(CustomLogger, cls).__new__(cls, *args, **kwargs)
            cls._instance._initialized = False
        return cls._instance

    def __init__(self):
        if getattr(self, "_initialized", False):
            return

        self.output_manager: Optional[OutputManager] = None
        self.log_filename = None
        self.execution_id = None
        self.log_level = "INFO"
        self.run_status = "RUNNING"
        self.buffer = []
        self.public_buffer = []
        self.public_buffer_max_lines = 50
        self.public_flush_seconds = 5
        self.last_public_flush_time = None
        self._initialized = True

    def configure(self, output_manager: OutputManager, execution_id: str, started_at: datetime, log_level: str = "INFO"):
        self.output_manager = output_manager
        self.execution_id = execution_id
        self.log_level = (log_level or "INFO").upper()
        timestamp = started_at.strftime("%Y-%m-%d_%H-%M-%S")
        self.log_filename = f"execution_{timestamp}__{execution_id}.jsonl"
        print(f"[LOGGER] Logs serao salvos em: {self.log_filename}")
        self.last_public_flush_time = time.time()

        if self.buffer:
            self._flush_buffer()

    def set_run_status(self, run_status: str):
        if run_status:
            self.run_status = run_status

    def log(self, level, message, details=None):
        levels = {"DEBUG": 10, "INFO": 20, "WARN": 30, "ERROR": 40}
        current = levels.get(self.log_level, 20)
        incoming = levels.get(level.upper(), 20)
        if incoming < current:
            return
        entry = {
            "timestamp": datetime.now().isoformat(),
            "execution_id": self.execution_id,
            "run_status": self.run_status,
            "level": level.upper(),
            "event": message,
            "message": message,
            "details": details or {}
        }

        line = f"[{level.upper()}] {message}"
        if details:
            try:
                details_text = json.dumps(details, ensure_ascii=False)
            except Exception:
                details_text = str(details)
            line = f"{line} | details={details_text}"
        print(line)

        self._append_entry(entry)

    def _append_entry(self, entry: dict):
        line = json.dumps(entry, ensure_ascii=False) + "\n"
        if not self.output_manager or not self.log_filename:
            self.buffer.append(line)
            return

        try:
            self.output_manager.append_text(KIND_LOGS, self.log_filename, line, write_public=False)
            self.public_buffer.append(line)
            if self._should_flush_public():
                self._flush_public()
        except Exception as e:
            print(f"[LOGGER ERROR] Failed to write log: {e}")

    def _flush_buffer(self):
        if not self.output_manager or not self.log_filename:
            return
        content = "".join(self.buffer)
        self.buffer = []
        try:
            self.output_manager.append_text(KIND_LOGS, self.log_filename, content, write_public=False)
            self.output_manager.append_public_text_only(KIND_LOGS, self.log_filename, content)
        except Exception as e:
            print(f"[LOGGER ERROR] Failed to flush logs: {e}")

    def _should_flush_public(self) -> bool:
        if not self.public_buffer:
            return False
        if len(self.public_buffer) >= self.public_buffer_max_lines:
            return True
        if self.last_public_flush_time is None:
            return True
        return (time.time() - self.last_public_flush_time) >= self.public_flush_seconds

    def _flush_public(self):
        if not self.output_manager or not self.log_filename or not self.public_buffer:
            return
        content = "".join(self.public_buffer)
        self.public_buffer = []
        try:
            self.output_manager.append_public_text_only(KIND_LOGS, self.log_filename, content)
            self.last_public_flush_time = time.time()
        except Exception as e:
            print(f"[LOGGER ERROR] Failed to flush public logs: {e}")

    def flush(self):
        self._flush_public()

    def info(self, message, details=None):
        self.log("INFO", message, details)

    def warn(self, message, details=None):
        self.log("WARN", message, details)

    def error(self, message, details=None):
        self.log("ERROR", message, details)

# Singleton instance
logger = CustomLogger()
