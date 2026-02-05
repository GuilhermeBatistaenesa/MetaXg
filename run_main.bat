@echo off
setlocal
set "VENV_DIR=.venv"
set "VENV_PY=%VENV_DIR%\\Scripts\\python.exe"
set "NEED_INSTALL=0"

where python >nul 2>nul
if errorlevel 1 (
  echo [run_main] Python nao encontrado no PATH.
  echo [run_main] Instale o Python 3.10+ e marque a opcao "Add to PATH".
  pause
  exit /b 1
)

if not exist "%VENV_PY%" (
  echo [run_main] criando venv...
  python -m venv "%VENV_DIR%"
  set "NEED_INSTALL=1"
)

if not exist "%VENV_DIR%\\.metax_ready" (
  set "NEED_INSTALL=1"
)

if "%NEED_INSTALL%"=="1" (
  echo [run_main] instalando dependencias...
  "%VENV_PY%" -m pip install --upgrade pip
  "%VENV_PY%" -m pip install -r requirements.txt
  echo [run_main] instalando browser do Playwright...
  "%VENV_PY%" -m playwright install chromium
  echo ready>"%VENV_DIR%\\.metax_ready"
)

echo [run_main] executando...
"%VENV_PY%" main.py %*
