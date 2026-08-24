@echo off
title Sistema Bancada - Docker WSL
cd /d "%~dp0"

:: 1. Verifica se o comando WSL está disponível no Windows
where wsl >nul 2>nul
if %ERRORLEVEL% equ 0 (
    echo =======================================================================
    echo         S I S T E M A   B A N C A D A   (D O C K E R   W S L)
    echo =======================================================================
    echo.
    echo  [INFO] Redirecionando execucao para o Docker no WSL...
    echo.
    wsl.exe -e bash -lic "cd \"$(wslpath '%~dp0')\" 2>/dev/null || cd /home/paulo/PythonProjects/automacao-chamados-sti; chmod +x 00-iniciar.sh; ./00-iniciar.sh %*"
    goto end
)

:: 2. Fallback: Mensagem de erro caso WSL não esteja instalado
echo.
echo [ERRO] O WSL nao foi encontrado neste sistema.
echo O Sistema Bancada e executado 100%% em containers Docker no WSL.
echo Instale o WSL e o Docker para continuar.
echo.

:end
pause

