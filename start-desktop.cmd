@echo off
cd /d "%~dp0"
if exist "dist\ThesisCraft.v4.1.0.exe" (
  start "" "dist\ThesisCraft.v4.1.0.exe" --general
) else (
  start "" ".venv\Scripts\pythonw.exe" "wfp.py" --general
)
