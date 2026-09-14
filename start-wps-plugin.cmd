@echo off
cd /d "%~dp0"
if exist "dist\ThesisCraft.v4.0.3.exe" (
  start "" "dist\ThesisCraft.v4.0.3.exe" --office wps
) else (
  start "" ".venv\Scripts\pythonw.exe" "wfp.py" --office wps
)
