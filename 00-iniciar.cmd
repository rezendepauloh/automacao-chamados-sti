@echo off
setlocal enabledelayedexpansion
title Sistema Bancada - CLI Unificado
cd /d "%~dp0"

:: Verificar parametro CLI
if /i "%~1"=="--start" goto :start_system
if /i "%~1"=="-s" goto :start_system
if /i "%~1"=="--build" goto :start_build
if /i "%~1"=="-b" goto :start_build
if /i "%~1"=="--config-senhas" goto :config_senhas
if /i "%~1"=="--senhas" goto :config_senhas
if /i "%~1"=="-p" goto :config_senhas
if /i "%~1"=="--orquestrador" goto :run_orquestrador
if /i "%~1"=="-o" goto :run_orquestrador
if /i "%~1"=="--rebuild" goto :rebuild_docker
if /i "%~1"=="-r" goto :rebuild_docker
if /i "%~1"=="--down" goto :stop_system
if /i "%~1"=="--stop" goto :stop_system
if /i "%~1"=="-d" goto :stop_system
if /i "%~1"=="--help" goto :show_help
if /i "%~1"=="-h" goto :show_help

:show_menu
cls
echo ================================================================
echo         SISTEMA BANCADA -- AUTOMACAO DE CHAMADOS STI            
echo ================================================================
echo.
echo   Escolha uma opcao:
echo   1 - Iniciar Sistema Bancada (Streamlit Dashboard)
echo   2 - Configurar Senhas (Cofre Keyring)
echo   3 - Executar Orquestrador de Sincronizacao
echo   4 - Reconstruir Docker Compose (--no-cache)
echo   5 - Parar sistema (docker compose down)
echo   0 - Sair
echo.
echo ================================================================
set /p "OPCAO=Opcao [0-5]: "

if "%OPCAO%"=="1" goto :start_system
if "%OPCAO%"=="2" goto :config_senhas
if "%OPCAO%"=="3" goto :run_orquestrador
if "%OPCAO%"=="4" goto :rebuild_docker
if "%OPCAO%"=="5" goto :stop_system
if "%OPCAO%"=="0" exit /b 0

echo Opcao invalida.
timeout /t 1 >nul
goto :show_menu

:setup_docker_and_ip
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

set "PORT=8502"
if exist ".env" (
    for /f "usebackq tokens=1,* delims==" %%i in (".env") do (
        set "KEY=%%i"
        set "VAL=%%j"
        if not "!KEY!"=="" (
            for /f "tokens=* delims= " %%k in ("!KEY!") do set "KEY=%%k"
            if "!KEY!"=="STREAMLIT_PORT" (
                for /f "tokens=* delims= " %%v in ("!VAL!") do set "PORT=%%v"
            )
        )
    )
)
if "%PORT%"=="" set "PORT=8502"
set "PORT=%PORT: =%"

if not exist "src\js" mkdir "src\js"
echo window.BANCADA_LOCAL_IP = '%LOCAL_IP%'; > src\js\server-info.js

where docker >nul 2>nul
if %ERRORLEVEL% equ 0 (
    set "DOCKER_CMD=docker compose"
    set "IS_WSL=0"
) else (
    where wsl >nul 2>nul
    if %ERRORLEVEL% equ 0 (
        set "DOCKER_CMD=wsl.exe docker compose"
        set "IS_WSL=1"
    ) else (
        echo [ERRO] Nem o Docker para Windows nem o WSL foram encontrados.
        echo Instale o Docker Desktop ou WSL para continuar.
        pause
        exit /b 1
    )
)
exit /b 0

:start_system
call :setup_docker_and_ip
cls
echo Iniciando Sistema Bancada (Porta: %PORT% ^| IP: %LOCAL_IP%)...
%DOCKER_CMD% up -d web
timeout /t 2 /nobreak >nul
start http://localhost:%PORT%/
goto :stream_logs

:start_build
call :setup_docker_and_ip
cls
echo Reconstruindo imagem e iniciando Sistema Bancada...
%DOCKER_CMD% build web
%DOCKER_CMD% up -d web
timeout /t 2 /nobreak >nul
start http://localhost:%PORT%/
goto :stream_logs

:stream_logs
echo ------------------------------------------------------------------------
echo  [C] Limpar Tela  ^|  [R] Reiniciar Web  ^|  [B] Navegador  ^|  [Q] Encerrar
echo ------------------------------------------------------------------------
%DOCKER_CMD% logs -f --tail=100 web
echo.
echo Encerrando containers do Sistema Bancada...
%DOCKER_CMD% down
exit /b 0

:config_senhas
call :setup_docker_and_ip
cls
echo Abrindo assistente de senhas no container...
%DOCKER_CMD% run --rm web python src/salvar_senha.py
echo.
echo Assistente finalizado!
pause
goto :show_menu

:run_orquestrador
call :setup_docker_and_ip
cls
echo Executando orquestrador de sincronizacao...
%DOCKER_CMD% run --rm web python orquestrador.py
echo.
echo Orquestrador finalizado!
pause
goto :show_menu

:rebuild_docker
call :setup_docker_and_ip
cls
echo Reconstruindo imagens Docker Compose (--no-cache)...
%DOCKER_CMD% build --no-cache web
echo.
echo Rebuild concluido!
pause
goto :show_menu

:stop_system
call :setup_docker_and_ip
echo Encerrando todos os containers do Sistema Bancada...
%DOCKER_CMD% down
echo Containers encerrados com sucesso!
pause
exit /b 0

:show_help
echo Uso: 00-iniciar.cmd [OPCAO]
echo.
echo Opcoes:
echo   --start, -s              Inicia o Sistema Bancada (Streamlit Dashboard)
echo   --build, -b              Reconstroi a imagem Docker e inicia
echo   --config-senhas, -p      Abre o configurador de senhas do cofre
echo   --orquestrador, -o       Executa o orquestrador de sincronizacao
echo   --rebuild, -r            Reconstroi a imagem Docker (--no-cache)
echo   --down, -d               Para os containers do sistema
echo   --help, -h               Exibe esta ajuda
echo   (sem argumentos)         Abre o menu interativo
exit /b 0
