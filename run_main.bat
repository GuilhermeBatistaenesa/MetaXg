@echo off
setlocal EnableExtensions
title MetaXg - Executar

rem ============================================================
rem MetaXg runner - robusto para qualquer PC
rem - Detecta Python real >= 3.10 com venv
rem - Venv por usuario (LOCALAPPDATA) para evitar conflitos
rem - Instala dependencias automaticamente
rem ============================================================

set "SCRIPT_DIR=%~dp0"
pushd "%SCRIPT_DIR%"

set "DIAG=%RUN_MAIN_DIAG%"
set "NO_PAUSE=%RUN_MAIN_NO_PAUSE%"
set "RECREATE_VENV=%RUN_MAIN_RECREATE_VENV%"
if not defined METAX_HOLD set "METAX_HOLD=1"
if defined RUN_MAIN_TRACE echo on

if not defined LOCALAPPDATA (
  if defined USERPROFILE (
    set "LOCALAPPDATA=%USERPROFILE%\AppData\Local"
  ) else (
    set "LOCALAPPDATA=%TEMP%"
  )
)

set "VENV_DIR=%RUN_MAIN_VENV_DIR%"
if not defined VENV_DIR set "VENV_DIR=%LOCALAPPDATA%\MetaXg\.venv"

set "EXIT_CODE=0"
set "BASE_PY_CMD="
set "PYTHON_EXE="

if defined DIAG (
  echo ===== RUN_MAIN_DIAG =====
  echo Current dir : "%CD%"
  echo Script dir  : "%SCRIPT_DIR%"
  echo Venv dir    : "%VENV_DIR%"
  echo LOCALAPPDATA: "%LOCALAPPDATA%"
  echo USERPROFILE : "%USERPROFILE%"
  if defined RECREATE_VENV echo Recreate venv: %RECREATE_VENV%
  where py >nul 2>&1
  if "%ERRORLEVEL%"=="0" (echo py launcher: OK) else (echo py launcher: NOT FOUND)
  where python >nul 2>&1
  if "%ERRORLEVEL%"=="0" (echo python in PATH: OK) else (echo python in PATH: NOT FOUND)
  echo =========================
)

rem ============================================================
rem 1) Find base Python (>=3.10 AND has venv)
rem ============================================================
call :find_base_python

if not defined BASE_PY_CMD (
  echo(
  echo Python nao encontrado. Requer versao 3.10+ e modulo venv.
  echo Instale Python completo. Nao use Windows Store stub ou embeddable.
  echo(
  set "EXIT_CODE=1"
  goto :end
)

if defined RECREATE_VENV (
  echo Forcando recriacao do venv...
  call :delete_venv
)

rem ============================================================
rem 2) Create venv if missing
rem ============================================================
if not exist "%VENV_DIR%\Scripts\python.exe" (
  echo Criando venv em "%VENV_DIR%"...
  call %BASE_PY_CMD% -m venv "%VENV_DIR%"
  if not exist "%VENV_DIR%\Scripts\python.exe" (
    echo(
    echo Falha ao criar venv.
    echo.
    set "EXIT_CODE=1"
    goto :end
  )
)

set "PYTHON_EXE=%VENV_DIR%\Scripts\python.exe"

rem ============================================================
rem 3) Validate venv python; if broken, recreate
rem ============================================================
"%PYTHON_EXE%" -c "import sys" >nul 2>&1
if not "%ERRORLEVEL%"=="0" (
  echo Venv quebrada. Recriando...
  call :delete_venv
  call %BASE_PY_CMD% -m venv "%VENV_DIR%"
  if not exist "%VENV_DIR%\Scripts\python.exe" (
    echo.
    echo Falha ao criar venv.
    echo.
    set "EXIT_CODE=1"
    goto :end
  )
  set "PYTHON_EXE=%VENV_DIR%\Scripts\python.exe"
)

rem ============================================================
rem 4) Ensure pip exists
rem ============================================================
"%PYTHON_EXE%" -m ensurepip --upgrade >nul 2>&1

rem ============================================================
rem 5) Install deps (always, unless METAX_SKIP_INSTALL=1)
rem ============================================================
if not exist "%SCRIPT_DIR%requirements.txt" (
  echo(
  echo requirements.txt nao encontrado.
  echo Current dir: "%CD%"
  echo(
  set "EXIT_CODE=1"
  goto :end
)

if /i "%METAX_SKIP_INSTALL%"=="1" (
  echo(Instalacao de dependencias ignorada ^(METAX_SKIP_INSTALL=1^).
  goto :after_install
)

echo Instalando dependencias (requirements.txt)...
"%PYTHON_EXE%" -m pip install --upgrade pip >nul 2>&1
"%PYTHON_EXE%" -m pip install -r "%SCRIPT_DIR%requirements.txt"
if not "%ERRORLEVEL%"=="0" (
  echo.
  echo Falha ao instalar dependencias.
  echo.
  set "EXIT_CODE=1"
  goto :end
)

rem ============================================================
rem 6) Playwright (Chromium)
rem ============================================================
call :ensure_playwright

:after_install

rem ============================================================
rem 7) Run main
rem ============================================================
if not exist "%SCRIPT_DIR%main.py" (
  echo.
  echo main.py nao encontrado.
  echo.
  set "EXIT_CODE=1"
  goto :end
)

set "PYTHONUNBUFFERED=1"
set "PYTHONIOENCODING=utf-8"

echo(
echo Running: main.py
"%PYTHON_EXE%" -u "%SCRIPT_DIR%main.py" %*
set "EXIT_CODE=%ERRORLEVEL%"

if not "%EXIT_CODE%"=="0" (
  echo(
  echo main.py exited with code %EXIT_CODE%
)

goto :end


rem ============================================================
rem Helpers
rem ============================================================

:delete_venv
if defined VENV_DIR (
  if exist "%VENV_DIR%" (
    cmd /c "rmdir /s /q \"%VENV_DIR%\"" >nul 2>&1
  )
)
exit /b 0


:ensure_playwright
set "PW_DIR=%LOCALAPPDATA%\ms-playwright"
if exist "%PW_DIR%" (
  dir /b "%PW_DIR%\chromium-*" >nul 2>&1
  if "%ERRORLEVEL%"=="0" exit /b 0
)
echo Instalando Playwright (Chromium)...
"%PYTHON_EXE%" -m playwright install chromium
exit /b 0


:find_base_python
set "BASE_PY_CMD="

rem 1) Try python.exe from PATH (where python)
for /f "delims=" %%P in ('where python 2^>nul') do (
  call :check_python_exe "%%P"
  if defined BASE_PY_CMD exit /b 0
)

rem 2) Try py launcher (preferred versions)
where py >nul 2>&1
if "%ERRORLEVEL%"=="0" (
  call :check_py_launcher "3.12"
  if defined BASE_PY_CMD exit /b 0
  call :check_py_launcher "3.11"
  if defined BASE_PY_CMD exit /b 0
  call :check_py_launcher "3.10"
  if defined BASE_PY_CMD exit /b 0
  call :check_py_launcher "3"
  if defined BASE_PY_CMD exit /b 0
)

exit /b 0


:check_python_exe
set "CANDIDATE=%~1"
"%CANDIDATE%" -c "import sys, venv; exit(0 if sys.version_info[:2]>=(3,10) else 1)" >nul 2>&1
if "%ERRORLEVEL%"=="0" (
  set "BASE_PY_CMD="%CANDIDATE%""
)
exit /b 0


:check_py_launcher
set "PYVER=%~1"
py -%PYVER% -c "import sys, venv; exit(0 if sys.version_info[:2]>=(3,10) else 1)" >nul 2>&1
if "%ERRORLEVEL%"=="0" (
  set "BASE_PY_CMD=py -%PYVER%"
)
exit /b 0


:end
echo(
if /i "%NO_PAUSE%"=="1" (
  rem no pause
) else (
  if "%METAX_HOLD%"=="0" (
    rem no pause
  ) else (
    pause
  )
)
popd
exit /b %EXIT_CODE%
