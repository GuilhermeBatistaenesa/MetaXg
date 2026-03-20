import inspect
import json
import os
import time
from datetime import datetime
from typing import Optional

from output_manager import KIND_LOGS, OutputManager


TERMINAL_LEVEL_MAP = {
    "DEBUG": "INFO ",
    "INFO": "INFO ",
    "OK": "OK   ",
    "WARN": "WARN ",
    "ERROR": "ERRO ",
    "FATAL": "FATAL",
    "RESUM": "RESUM",
}


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
        self.robot_name = "MetaXg"
        self.robot_version = "0.0.0"
        self.environment_name = "Producao"
        self.current_step = "Inicializacao"
        self.current_stage_index = None
        self.current_stage_total = None
        self.current_stage_name = None
        self.buffer = []
        self.public_buffer = []
        self.public_buffer_max_lines = 50
        self.public_flush_seconds = 5
        self.last_public_flush_time = None
        self._initialized = True

    def configure(
        self,
        output_manager: OutputManager,
        execution_id: str,
        started_at: datetime,
        log_level: str = "INFO",
        robot_name: str = "MetaXg",
        robot_version: str = "0.0.0",
        environment_name: str = "Producao",
    ):
        self.output_manager = output_manager
        self.execution_id = execution_id
        self.log_level = (log_level or "INFO").upper()
        self.robot_name = robot_name or "MetaXg"
        self.robot_version = robot_version or "0.0.0"
        self.environment_name = environment_name or "Producao"
        timestamp = started_at.strftime("%Y-%m-%d_%H-%M-%S")
        self.log_filename = f"execution_{timestamp}__{execution_id}.jsonl"
        self.last_public_flush_time = time.time()
        self._print_header(started_at)

        if self.buffer:
            self._flush_buffer()

    def _print_header(self, started_at: datetime):
        print("============================================================", flush=True)
        print("ENESA | AUTOMACAO CORPORATIVA", flush=True)
        print(f"Robo: {self.robot_name}", flush=True)
        print(f"Versao: {self.robot_version}", flush=True)
        print(f"Execucao: {self.execution_id}", flush=True)
        print(f"Ambiente: {self.environment_name}", flush=True)
        print(f"Inicio: {started_at.strftime('%Y-%m-%d %H:%M:%S')}", flush=True)
        print("============================================================", flush=True)

    def set_run_status(self, run_status: str):
        if run_status:
            self.run_status = run_status

    def stage(self, index: int, total: int, name: str):
        self.current_stage_index = index
        self.current_stage_total = total
        self.current_stage_name = name
        self.current_step = name
        print("", flush=True)
        print(f"[ETAPA {index}/{total}] {name}", flush=True)
        print("------------------------------------------------------------", flush=True)
        self._append_entry(
            self._build_entry(
                level="INFO",
                message=f"Etapa iniciada: {name}",
                details={"stage_index": index, "stage_total": total, "stage_name": name},
                event_type="stage_start",
                source_file="custom_logger.py",
            )
        )

    def finish_summary(
        self,
        started_at: datetime,
        finished_at: datetime,
        status: str,
        report_path: str | None,
        totals: dict | None = None,
    ):
        totals = totals or {}
        duration = finished_at - started_at
        print("", flush=True)
        print("============================================================", flush=True)
        print("RESUMO FINAL DA EXECUCAO", flush=True)
        print(f"Robo: {self.robot_name}", flush=True)
        print(f"Execucao: {self.execution_id}", flush=True)
        print(f"Status: {status}", flush=True)
        if totals:
            print(f"Itens recebidos: {totals.get('detected', 0)}", flush=True)
            print(f"Itens processados: {totals.get('people_total', 0)}", flush=True)
            erros = (
                totals.get("by_outcome", {}).get("FAILED_ACTION", 0)
                + totals.get("by_outcome", {}).get("FAILED_VERIFICATION", 0)
                + totals.get("by_outcome", {}).get("SAVED_NOT_VERIFIED", 0)
            )
            print(f"Itens com erro: {erros}", flush=True)
        if report_path:
            print(f"Relatorio: {report_path}", flush=True)
        print(f"Fim: {finished_at.strftime('%Y-%m-%d %H:%M:%S')}", flush=True)
        print(f"Duracao: {self._format_duration(duration.total_seconds())}", flush=True)
        print("============================================================", flush=True)

    def _format_duration(self, total_seconds: float) -> str:
        total_seconds = int(total_seconds or 0)
        hours = total_seconds // 3600
        minutes = (total_seconds % 3600) // 60
        seconds = total_seconds % 60
        return f"{hours:02d}:{minutes:02d}:{seconds:02d}"

    def log(self, level, message, details=None, event_type: str | None = None):
        levels = {"DEBUG": 10, "INFO": 20, "OK": 20, "RESUM": 20, "WARN": 30, "ERROR": 40, "FATAL": 50}
        current = levels.get(self.log_level, 20)
        incoming = levels.get(level.upper(), 20)
        if incoming < current:
            return

        source_file = self._resolve_source_file()
        entry = self._build_entry(
            level=level,
            message=message,
            details=details,
            event_type=event_type,
            source_file=source_file,
        )
        self._print_terminal_line(level.upper(), message)
        self._append_entry(entry)

    def _resolve_source_file(self) -> str:
        try:
            for frame in inspect.stack()[2:]:
                filename = os.path.basename(frame.filename)
                if filename != "custom_logger.py":
                    return filename
        except Exception:
            pass
        return "custom_logger.py"

    def _build_entry(self, level, message, details=None, event_type: str | None = None, source_file: str | None = None):
        return {
            "timestamp_utc": datetime.utcnow().replace(microsecond=0).isoformat() + "Z",
            "timestamp": datetime.now().isoformat(),
            "robot_name": self.robot_name,
            "robot_version": self.robot_version,
            "execution_id": self.execution_id,
            "run_status": self.run_status,
            "step": self.current_step,
            "event_type": event_type or self._infer_event_type(level, message),
            "severity": self._normalize_severity(level),
            "level": level.upper(),
            "message": message,
            "source_file": source_file or "custom_logger.py",
            "correlation_keys": self._build_correlation_keys(details),
            "details": details or {},
        }

    def _infer_event_type(self, level: str, message: str) -> str:
        level = level.upper()
        if level in {"ERROR", "FATAL"}:
            return "failure"
        if level == "WARN":
            return "warning"
        if level == "OK":
            return "success"
        if level == "RESUM":
            return "summary"
        msg = (message or "").lower()
        if "iniciando" in msg or "acessando" in msg or "processando" in msg or "gerando" in msg:
            return "action"
        return "info"

    def _normalize_severity(self, level: str) -> str:
        level = level.upper()
        if level == "ERROR":
            return "ERRO"
        if level in {"OK", "RESUM"}:
            return level
        return level

    def _build_correlation_keys(self, details) -> dict:
        details = details or {}
        keys = {}
        for key in ("cpf", "funcionario", "nome", "contrato", "contrato_chave", "path", "run_id"):
            if key in details and details.get(key):
                keys[key] = details.get(key)
        return keys

    def _print_terminal_line(self, level: str, message: str):
        visual = TERMINAL_LEVEL_MAP.get(level.upper(), "INFO ")
        print(f"[{visual}] {message}", flush=True)

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
            print(f"[ERRO ] Falha ao gravar log tecnico: {e}", flush=True)

    def _flush_buffer(self):
        if not self.output_manager or not self.log_filename:
            return
        content = "".join(self.buffer)
        self.buffer = []
        try:
            self.output_manager.append_text(KIND_LOGS, self.log_filename, content, write_public=False)
            self.output_manager.append_public_text_only(KIND_LOGS, self.log_filename, content)
        except Exception as e:
            print(f"[ERRO ] Falha ao descarregar buffer de logs: {e}", flush=True)

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
            print(f"[ERRO ] Falha ao publicar log tecnico: {e}", flush=True)

    def flush(self):
        self._flush_public()

    def info(self, message, details=None):
        self.log("INFO", message, details)

    def ok(self, message, details=None):
        self.log("OK", message, details, event_type="success")

    def warn(self, message, details=None):
        self.log("WARN", message, details)

    def error(self, message, details=None):
        self.log("ERROR", message, details)

    def fatal(self, message, details=None):
        self.log("FATAL", message, details, event_type="failure")

    def resum(self, message, details=None):
        self.log("RESUM", message, details, event_type="summary")


logger = CustomLogger()
