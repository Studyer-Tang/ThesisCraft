@echo off
cd /d "%~dp0"
if exist "dist\ThesisCraft.v4.0.2.exe" (
  start "" "dist\ThesisCraft.v4.0.2.exe" --office wps
) else (
  start "" ".venv\Scripts\pythonw.exe" "wfp.py" --office wps
)
