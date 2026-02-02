@echo off
setlocal
powershell -ExecutionPolicy Bypass -File "%~dp0scripts\\build_windows.ps1" %*
