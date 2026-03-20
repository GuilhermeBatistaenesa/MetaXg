import os
import sys
from dotenv import load_dotenv

# Resolve base dir for .env (exe dir when frozen, repo root otherwise)
if getattr(sys, "frozen", False):
    BASE_DIR = os.path.dirname(sys.executable)
else:
    BASE_DIR = os.path.dirname(os.path.abspath(__file__))

# Load environment variables from .env file (if present)
load_dotenv(dotenv_path=os.path.join(BASE_DIR, ".env"))

ROOT_DIR = BASE_DIR
DEFAULT_PUBLIC_BASE_DIR = r"P:\ProcessoMetaX"

# Sharepoint Configuration
SHAREPOINT_SITE_URL = os.getenv("SHAREPOINT_SITE_URL")
SHAREPOINT_CLIENT_ID = os.getenv("SHAREPOINT_CLIENT_ID")
SHAREPOINT_CLIENT_SECRET = os.getenv("SHAREPOINT_CLIENT_SECRET")

# SQL Server Configuration
DB_DRIVER = os.getenv("DB_DRIVER", "SQL Server")
DB_SERVER = os.getenv("DB_SERVER")
DB_NAME = os.getenv("DB_NAME")
DB_USER = os.getenv("DB_USER")
DB_PASSWORD = os.getenv("DB_PASSWORD")

# MetaX Portal Configuration
METAX_LOGIN = os.getenv("METAX_LOGIN")
METAX_PASSWORD = os.getenv("METAX_PASSWORD")
METAX_URL_LOGIN = os.getenv("METAX_URL_LOGIN", "https://portal.metax.ind.br/SegLogin/")
METAX_CONTRATO_MECANICA_VALUE = os.getenv("METAX_CONTRATO_MECANICA_VALUE")
METAX_CONTRATO_MECANICA_LABEL = os.getenv("METAX_CONTRATO_MECANICA_LABEL")
METAX_CONTRATO_ELETROMECANICA_VALUE = os.getenv("METAX_CONTRATO_ELETROMECANICA_VALUE")
METAX_CONTRATO_ELETROMECANICA_LABEL = os.getenv("METAX_CONTRATO_ELETROMECANICA_LABEL")
METAX_CONTRATO_DEFAULT_VALUE = os.getenv("METAX_CONTRATO_DEFAULT_VALUE")
METAX_CONTRATO_DEFAULT_LABEL = os.getenv("METAX_CONTRATO_DEFAULT_LABEL")

def _default_public_base() -> str:
    env_base = os.getenv("PUBLIC_BASE_DIR")
    if env_base:
        return env_base
    if os.path.exists(DEFAULT_PUBLIC_BASE_DIR):
        return DEFAULT_PUBLIC_BASE_DIR
    return ROOT_DIR

PUBLIC_BASE_DIR = _default_public_base()
OBJECT_NAME = os.getenv("OBJECT_NAME", "MetaXg")
PUBLIC_INPUTS_DIR = os.getenv("PUBLIC_INPUTS_DIR", os.path.join(PUBLIC_BASE_DIR, "entrada"))
PUBLIC_CODE_DIR = os.getenv("PUBLIC_CODE_DIR", os.path.join(r"Z:\T.I", "MetaXg"))
PUBLIC_PROCESSADOS_DIR = os.getenv("PUBLIC_PROCESSADOS_DIR", os.path.join(PUBLIC_BASE_DIR, "fotos", "processados"))
PUBLIC_ERROS_DIR = os.getenv("PUBLIC_ERROS_DIR", os.path.join(PUBLIC_BASE_DIR, "fotos", "erros"))
PUBLIC_LOGS_DIR = os.getenv("PUBLIC_LOGS_DIR", os.path.join(PUBLIC_BASE_DIR, "logs"))
PUBLIC_RELATORIOS_DIR = os.getenv("PUBLIC_RELATORIOS_DIR", os.path.join(PUBLIC_BASE_DIR, "relatorios"))
PUBLIC_JSON_DIR = os.getenv("PUBLIC_JSON_DIR", os.path.join(PUBLIC_BASE_DIR, "json"))
PUBLIC_RELEASES_DIR = os.getenv("PUBLIC_RELEASES_DIR", os.path.join(PUBLIC_BASE_DIR, "releases"))
PUBLIC_SCREENSHOTS_DIR = os.getenv("PUBLIC_SCREENSHOTS_DIR", os.path.join(PUBLIC_BASE_DIR, "screenshots"))

# Paths
FOTOS_EM_PROCESSAMENTO_DIR = os.getenv(
    "FOTOS_EM_PROCESSAMENTO_DIR",
    os.getenv("PASTA_FOTOS", os.path.join(PUBLIC_BASE_DIR, "fotos", "em_processamento")),
)
FOTOS_PROCESSADOS_DIR = os.getenv("FOTOS_PROCESSADOS_DIR", PUBLIC_PROCESSADOS_DIR)
FOTOS_ERROS_DIR = os.getenv("FOTOS_ERROS_DIR", PUBLIC_ERROS_DIR)
FOTOS_BUSCA_DIRS = [p for p in [FOTOS_EM_PROCESSAMENTO_DIR, FOTOS_PROCESSADOS_DIR, FOTOS_ERROS_DIR] if p]
PASTA_FOTOS = FOTOS_EM_PROCESSAMENTO_DIR
LOG_DIR = "logs"
SCREENSHOT_DIR = os.path.join(LOG_DIR, "screenshots")

# Configuração de Dias Retroativos
try:
    DIAS_RETROATIVOS = int(os.getenv("DIAS_RETROATIVOS", "0"))
except ValueError:
    DIAS_RETROATIVOS = 0

def _normalize_email_recipients(raw_value: str | None) -> str:
    if not raw_value:
        return ""

    normalized = raw_value.replace(";", ",")
    recipients = []
    for item in normalized.split(","):
        email = item.strip().strip('"').strip("'")
        if email:
            recipients.append(email)
    return "; ".join(recipients)


EMAIL_NOTIFICACAO = _normalize_email_recipients(os.getenv("EMAIL_NOTIFICACAO", ""))
EMAIL_REMETENTE = os.getenv("EMAIL_REMETENTE", "")
EMAIL_SMTP_HOST = os.getenv("EMAIL_SMTP_HOST", "")
try:
    EMAIL_SMTP_PORT = int(os.getenv("EMAIL_SMTP_PORT", "587"))
except ValueError:
    EMAIL_SMTP_PORT = 587
EMAIL_SMTP_USER = os.getenv("EMAIL_SMTP_USER", "")
EMAIL_SMTP_PASSWORD = os.getenv("EMAIL_SMTP_PASSWORD", "")
EMAIL_SMTP_USE_TLS = os.getenv("EMAIL_SMTP_USE_TLS", "1").strip().lower() not in {"0", "false", "no", "off"}
EMAIL_SMTP_USE_SSL = os.getenv("EMAIL_SMTP_USE_SSL", "0").strip().lower() in {"1", "true", "yes", "on"}
try:
    EMAIL_MAX_ATTACHMENTS_MB = int(os.getenv("EMAIL_MAX_ATTACHMENTS_MB", "12"))
except ValueError:
    EMAIL_MAX_ATTACHMENTS_MB = 12

# Ensure directories exist
os.makedirs(LOG_DIR, exist_ok=True)
os.makedirs(SCREENSHOT_DIR, exist_ok=True)

RELATORIOS_DIR = os.path.join(ROOT_DIR, "relatorios")
JSON_DIR = os.path.join(ROOT_DIR, "json")
os.makedirs(RELATORIOS_DIR, exist_ok=True)
os.makedirs(JSON_DIR, exist_ok=True)
