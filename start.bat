@echo off
chcp 65001 > nul
title AutomacaoML - Interface Web

:: Mata qualquer instância anterior na porta 8000
for /f "tokens=5" %%a in ('netstat -aon ^| findstr ":8000 " 2^>nul') do (
    taskkill /F /PID %%a > nul 2>&1
)
timeout /t 1 > nul

echo.
echo  Instalando dependencias...
pip install -r requirements.txt -q

echo  Iniciando servidor...
echo  Acesse: http://localhost:8000
echo.

start "" http://localhost:8000
python -m uvicorn backend.server:app --host 127.0.0.1 --port 8000

pause
