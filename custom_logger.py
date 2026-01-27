import json
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
        self.buffer = []
        self._initialized = True

    def configure(self, output_manager: OutputManager, execution_id: str, started_at: datetime, log_level: str = "INFO"):
        self.output_manager = output_manager
        self.log_level = (log_level or "INFO").upper()
        timestamp = started_at.strftime("%Y-%m-%d_%H-%M-%S")
        self.log_filename = f"execution_{timestamp}__{execution_id}.log"
        print(f"[LOGGER] Logs serao salvos em: {self.log_filename}")

        if self.buffer:
            self._flush_buffer()

    def log(self, level, message, details=None):
        levels = {"DEBUG": 10, "INFO": 20, "WARN": 30, "ERROR": 40}
        current = levels.get(self.log_level, 20)
        incoming = levels.get(level.upper(), 20)
        if incoming < current:
            return
        entry = {
            "timestamp": datetime.now().isoformat(),
            "level": level.upper(),
            "message": message,
            "details": details or {}
        }

        # 1. Print to Terminal
        print(f"[{level.upper()}] {message}")
        if details:
            # Only print simple details to terminal to avoid clutter, or skipped
            pass 

        self._append_entry(entry)

    def _append_entry(self, entry: dict):
        line = json.dumps(entry, ensure_ascii=False) + "\n"
        if not self.output_manager or not self.log_filename:
            self.buffer.append(line)
            return

        try:
            self.output_manager.append_text(KIND_LOGS, self.log_filename, line)
        except Exception as e:
            print(f"[LOGGER ERROR] Failed to write log: {e}")

    def _flush_buffer(self):
        if not self.output_manager or not self.log_filename:
            return
        content = "".join(self.buffer)
        self.buffer = []
        try:
            self.output_manager.write_text(KIND_LOGS, self.log_filename, content)
        except Exception as e:
            print(f"[LOGGER ERROR] Failed to flush logs: {e}")

    def info(self, message, details=None):
        self.log("INFO", message, details)

    def warn(self, message, details=None):
        self.log("WARN", message, details)

    def error(self, message, details=None):
        self.log("ERROR", message, details)

# Singleton instance
logger = CustomLogger()
