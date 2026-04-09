@echo off
setlocal
cd /d "%~dp0"
start "" "http://localhost:8501"
".\.venv\Scripts\python.exe" -m streamlit run ".\dashboard_app.py" --server.headless true --server.port 8501
