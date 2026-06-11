@echo off
cd /d "c:\Users\soyou\OneDrive\바탕 화면\vs"
start "" /B .venv\Scripts\streamlit.exe run dashboard_app.py
echo 대시보드가 백그라운드에서 실행되었습니다.
echo 브라우저에서 http://localhost:8501 로 접속하세요.
pause
