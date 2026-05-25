@echo off
chcp 65001 > nul
title MaxGroup RH
cd /d "%~dp0"

echo.
echo  ╔══════════════════════════════════════╗
echo  ║        MaxGroup — Sistema RH         ║
echo  ╚══════════════════════════════════════╝
echo.

if not exist "venv\Scripts\python.exe" (
    echo  [ERRO] Ambiente virtual não encontrado.
    echo  Execute novamente o setup ou contate o suporte.
    pause
    exit /b 1
)

echo  Iniciando servidor...
start "" cmd /c "timeout /t 3 /nobreak > nul && start http://localhost:5000"
echo  Acesse: http://localhost:5000
echo  Para encerrar: pressione Ctrl+C
echo.
venv\Scripts\python.exe app.py
pause
