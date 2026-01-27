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

PUBLIC_BASE_DIR = os.getenv("PUBLIC_BASE_DIR", r"P:\GuilhermeCostaProenca")
OBJECT_NAME = os.getenv("OBJECT_NAME", "MetaX")

# Paths
PASTA_FOTOS = os.getenv("PASTA_FOTOS", r"P:\MetaX\fotos_funcionarios")
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
