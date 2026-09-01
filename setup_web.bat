@echo off
setlocal
cd /d "%~dp0"
where py >nul 2>nul
if %errorlevel%==0 (
  set "PY=py"
) else (
  set "PY=python"
)
echo Instalando dependencias do BHead M365 Migrator Web...
%PY% -m pip install -r requirements.txt
if errorlevel 1 (
  echo.
  echo FALHA ao instalar dependencias.
  pause
  exit /b 1
)
echo.
echo Ambiente Web pronto.
pause
