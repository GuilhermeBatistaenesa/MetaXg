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

PUBLIC_BASE_DIR = os.getenv("PUBLIC_BASE_DIR", r"P:\ProcessoMetaX")
OBJECT_NAME = os.getenv("OBJECT_NAME", "MetaXg")
PUBLIC_INPUTS_DIR = os.getenv("PUBLIC_INPUTS_DIR", os.path.join(PUBLIC_BASE_DIR, "em processamento"))
PUBLIC_CODE_DIR = os.getenv("PUBLIC_CODE_DIR", os.path.join(PUBLIC_BASE_DIR, "Codigo"))
PUBLIC_PROCESSADOS_DIR = os.getenv("PUBLIC_PROCESSADOS_DIR", os.path.join(PUBLIC_BASE_DIR, "processados"))
PUBLIC_ERROS_DIR = os.getenv("PUBLIC_ERROS_DIR", os.path.join(PUBLIC_BASE_DIR, "erros"))
PUBLIC_LOGS_DIR = os.getenv("PUBLIC_LOGS_DIR", os.path.join(PUBLIC_BASE_DIR, "logs"))
PUBLIC_RELATORIOS_DIR = os.getenv("PUBLIC_RELATORIOS_DIR", os.path.join(PUBLIC_BASE_DIR, "relatorios"))
PUBLIC_JSON_DIR = os.getenv("PUBLIC_JSON_DIR", os.path.join(PUBLIC_BASE_DIR, "json"))
PUBLIC_RELEASES_DIR = os.getenv("PUBLIC_RELEASES_DIR", os.path.join(PUBLIC_BASE_DIR, "releases"))

# Paths
FOTOS_EM_PROCESSAMENTO_DIR = os.getenv(
    "FOTOS_EM_PROCESSAMENTO_DIR",
    os.getenv("PASTA_FOTOS", os.path.join(PUBLIC_BASE_DIR, "em processamento")),
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

EMAIL_NOTIFICACAO = os.getenv("EMAIL_NOTIFICACAO", "")

# Ensure directories exist
os.makedirs(LOG_DIR, exist_ok=True)
os.makedirs(SCREENSHOT_DIR, exist_ok=True)

RELATORIOS_DIR = os.path.join(ROOT_DIR, "relatorios")
JSON_DIR = os.path.join(ROOT_DIR, "json")
os.makedirs(RELATORIOS_DIR, exist_ok=True)
os.makedirs(JSON_DIR, exist_ok=True)
