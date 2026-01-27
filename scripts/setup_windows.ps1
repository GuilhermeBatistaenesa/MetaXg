$ErrorActionPreference = "Stop"

Write-Host "[setup] criando venv..."
python -m venv .venv

Write-Host "[setup] ativando venv..."
. .\.venv\Scripts\Activate.ps1

Write-Host "[setup] instalando dependencias..."
pip install -r requirements.txt

Write-Host "[setup] instalando browsers do Playwright..."
python -m playwright install chromium

Write-Host "[setup] smoke test Playwright..."
@"
from playwright.sync_api import sync_playwright

with sync_playwright() as p:
    browser = p.chromium.launch(headless=True)
    browser.close()
print("Smoke test OK")
"@ | python -

Write-Host "[setup] concluido."
