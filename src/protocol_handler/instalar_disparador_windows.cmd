@echo off
setlocal
title Instalador do Launcher Sistema Bancada

echo ============================================================
echo   INSTALANDO DISPARADOR DIRETO DO SISTEMA BANCADA (WINDOWS)
echo ============================================================
echo.

set "TARGET_DIR=%USERPROFILE%\.bancada"
if not exist "%TARGET_DIR%" mkdir "%TARGET_DIR%"

echo [1/3] Copiando script launcher para %TARGET_DIR%...
copy /y "%~dp0bancada-launcher.ps1" "%TARGET_DIR%\bancada-launcher.ps1" >nul

echo [2/3] Registrando protocolo bancada:// no Windows Registry...
reg add "HKCU\Software\Classes\bancada" /ve /d "URL:Sistema Bancada Protocol" /f >nul
reg add "HKCU\Software\Classes\bancada" /v "URL Protocol" /d "" /f >nul
reg add "HKCU\Software\Classes\bancada\shell" /f >nul
reg add "HKCU\Software\Classes\bancada\shell\open" /f >nul
reg add "HKCU\Software\Classes\bancada\shell\open\command" /ve /d "powershell.exe -NoExit -ExecutionPolicy Bypass -File \"%USERPROFILE%\.bancada\bancada-launcher.ps1\" \"%%1\"" /f >nul

if %ERRORLEVEL% equ 0 (
    echo [OK] Protocolo 'bancada://' registrado com sucesso no HKCU!
) else (
    echo [AVISO] Ocorreu um problema ao registrar no HKCU.
)

echo [3/3] Pronto! O launcher seleciona automaticamente o PowerShell 7 (pwsh.exe) ou 5.1 conforme sua escolha.
echo.
echo ============================================================
echo  Instalacao concluida com sucesso! Pressione qualquer tecla para sair.
echo ============================================================
pause >nul
exit /b 0
