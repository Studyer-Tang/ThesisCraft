@echo off
cd /d "%~dp0"
if exist "dist\ThesisCraft.v4.0.4.exe" (
  start "" "dist\ThesisCraft.v4.0.4.exe"
) else (
  start "" ".venv\Scripts\pythonw.exe" "wfp.py"
)
