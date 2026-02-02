@echo off
setlocal
powershell -ExecutionPolicy Bypass -File "%~dp0scripts\\build_zip.ps1" %*
