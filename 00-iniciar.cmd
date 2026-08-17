@echo off
title Sistema Bancada - Servidor Local SQLite
cd /d "%~dp0"
call venv\Scripts\activate.bat
cls
echo =======================================================================
echo                        S I S T E M A   B A N C A D A
echo =======================================================================
echo.

:: Detectar IP Local IPv4
set "LOCAL_IP="
for /f "tokens=4" %%a in ('route print ^| findstr "\<0.0.0.0\>"') do (
    if not defined LOCAL_IP set "LOCAL_IP=%%a"
)
if "%LOCAL_IP%"=="" (
    for /f "tokens=2 delims=:" %%a in ('ipconfig ^| findstr /c:"IPv4" /c:"IP Address"') do (
        if not defined LOCAL_IP (
            for /f "tokens=1" %%b in ("%%a") do set "LOCAL_IP=%%b"
        )
    )
)
if "%LOCAL_IP%"=="" set "LOCAL_IP=localhost"

set "PORT=8501"
set "FULL_URL=http://%LOCAL_IP%:%PORT%/"

if not exist "src\js" mkdir "src\js"
echo window.BANCADA_LOCAL_IP = '%LOCAL_IP%'; > src\js\server-info.js

echo  [OK] Banco de dados SQLite centralizado ativo em chamados.db
echo.
echo  -----------------------------------------------------------------------
echo   ACESSO NO COMPUTADOR : http://localhost:%PORT%/
echo   ACESSO NO CELULAR / TABLET: %FULL_URL%
echo  -----------------------------------------------------------------------
echo.
echo  Aponte a camera do Celular para o QR Code abaixo:
echo.

curl.exe -A "curl" -s "https://qrenco.de/%FULL_URL%" 2>nul

echo.
echo =======================================================================
echo  Pressione CTRL+C para encerrar o servidor.
echo =======================================================================
echo.

streamlit run dashboard.py --server.port 8501 --logger.level=error

pause
