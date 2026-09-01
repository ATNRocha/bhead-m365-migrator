@echo off
setlocal
cd /d "%~dp0"
where py >nul 2>nul
if %errorlevel%==0 (set "PY=py") else (set "PY=python")
echo ==========================================================
echo BHead M365 Migrator Web - REDE LOCAL
echo ==========================================================
echo.
echo O servidor ficara ativo enquanto esta janela estiver aberta.
echo Nas outras estacoes, acesse: http://IP-DESTE-PC:8501
echo Para descobrir o IP, execute ipconfig e veja o IPv4.
echo.
start "" http://127.0.0.1:8501
%PY% -m streamlit run app_web.py --server.address=0.0.0.0 --server.port=8501
