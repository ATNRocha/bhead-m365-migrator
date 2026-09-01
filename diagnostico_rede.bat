@echo off
echo IPv4 desta maquina:
ipconfig | findstr /I "IPv4"
echo.
echo Porta 8501:
netstat -ano | findstr ":8501"
echo.
pause
