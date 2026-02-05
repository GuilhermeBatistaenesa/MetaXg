@echo off
setlocal
set "VENV_DIR=.venv"
set "VENV_PY=%VENV_DIR%\Scripts\python.exe"
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

if not exist "%VENV_DIR%\Lib\site-packages\pip\__main__.py" (
  echo [run_main] reparando pip...
  "%VENV_PY%" -m ensurepip --upgrade
)

if not exist "%VENV_DIR%\.metax_ready" (
  set "NEED_INSTALL=1"
)

if "%NEED_INSTALL%"=="0" (
  "%VENV_PY%" -m pip show openpyxl >nul 2>nul
  if errorlevel 1 (
    set "NEED_INSTALL=1"
  )
)

if "%METAX_SKIP_INSTALL%"=="1" (
  set "NEED_INSTALL=0"
)

if "%NEED_INSTALL%"=="1" (
  echo [run_main] instalando dependencias...
  "%VENV_PY%" -m pip install -r requirements.txt
  if errorlevel 1 (
    echo [run_main] falha ao instalar dependencias.
    pause
    exit /b 1
  )
  echo [run_main] instalando browser do Playwright...
  "%VENV_PY%" -m playwright install chromium
  if errorlevel 1 (
    echo [run_main] falha ao instalar browser do Playwright.
    pause
    exit /b 1
  )
  echo ready>"%VENV_DIR%\.metax_ready"
)

if "%METAX_HOLD%"=="" (
  set "METAX_HOLD=1"
)

echo [run_main] executando...
call "%VENV_PY%" main.py %*
set "EXITCODE=%ERRORLEVEL%"
if not "%EXITCODE%"=="0" (
  echo [run_main] erro codigo %EXITCODE%
  pause
  exit /b %EXITCODE%
)

if "%METAX_HOLD%"=="1" (
  pause
)
exit /b 0
