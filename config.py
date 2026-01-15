import os
from dotenv import load_dotenv

# Load environment variables from .env file
load_dotenv()

# Sharepoint Configuration
SHAREPOINT_SITE_URL = os.getenv("SHAREPOINT_SITE_URL")
SHAREPOINT_CLIENT_ID = os.getenv("SHAREPOINT_CLIENT_ID")
SHAREPOINT_CLIENT_SECRET = os.getenv("SHAREPOINT_CLIENT_SECRET")

# SQL Server Configuration
DB_DRIVER = os.getenv("DB_DRIVER", "ODBC Driver 17 for SQL Server")
DB_SERVER = os.getenv("DB_SERVER")
DB_NAME = os.getenv("DB_NAME")
DB_USER = os.getenv("DB_USER")
DB_PASSWORD = os.getenv("DB_PASSWORD")

# MetaX Portal Configuration
METAX_LOGIN = os.getenv("METAX_LOGIN")
METAX_PASSWORD = os.getenv("METAX_PASSWORD")
METAX_URL_LOGIN = os.getenv("METAX_URL_LOGIN", "https://portal.metax.ind.br/SegLogin/")

# Paths
PASTA_FOTOS = os.getenv("PASTA_FOTOS", r"P:\MetaX\fotos_funcionarios")
LOG_DIR = "logs"
SCREENSHOT_DIR = os.path.join(LOG_DIR, "screenshots")

# Ensure directories exist
os.makedirs(LOG_DIR, exist_ok=True)
os.makedirs(SCREENSHOT_DIR, exist_ok=True)
