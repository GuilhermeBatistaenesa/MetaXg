@echo off
setlocal
if exist .venv\Scripts\python.exe (
  set "PYTHON=.venv\Scripts\python.exe"
) else (
  set "PYTHON=python"
)
%PYTHON% main.py %*
