@echo off
REM Absatzplanung: startet die App nur auf diesem Rechner (localhost)
cd /d "%~dp0"
python -c "import streamlit, plotly, openpyxl" 2>nul || (
  echo Erster Start: installiere benoetigte Pakete ...
  python -m pip install -r requirements.txt
)
python -m streamlit run planung_app.py --server.address localhost --server.headless false
pause
