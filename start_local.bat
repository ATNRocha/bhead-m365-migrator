@echo off
setlocal
cd /d "%~dp0"
where py >nul 2>nul
if %errorlevel%==0 (set "PY=py") else (set "PY=python")
start "" http://127.0.0.1:8501
%PY% -m streamlit run app_web.py --server.address=127.0.0.1 --server.port=8501
