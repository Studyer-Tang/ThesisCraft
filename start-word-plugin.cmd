@echo off
cd /d "%~dp0"
if exist "dist\ThesisCraft.v4.0.3.exe" (
  start "" "dist\ThesisCraft.v4.0.3.exe" --office word
) else (
  start "" ".venv\Scripts\pythonw.exe" "wfp.py" --office word
)
