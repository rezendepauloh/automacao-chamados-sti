#!/usr/bin/env bash
# =======================================================================
#               SISTEMA BANCADA — CLI UNIFICADO
# =======================================================================

cd "$(dirname "$0")" || exit 1

# Paleta de Cores ANSI
C_RESET="\033[0m"
C_BOLD="\033[1m"
C_CYAN="\033[36m"
C_GREEN="\033[32m"
C_YELLOW="\033[33m"
C_RED="\033[31m"
C_MAGENTA="\033[35m"
C_GRAY="\033[90m"

detect_ip() {
    LOCAL_IP=""
    # 1. Se estiver no WSL, tenta obter o IP do adaptador de rede Windows real
    if [ -f /proc/sys/fs/binfmt_misc/WSLInterop ] || [ -n "$WSL_DISTRO_NAME" ]; then
        if command -v powershell.exe >/dev/null 2>&1; then
            WIN_IP=$(powershell.exe -NoProfile -Command "(Get-NetIPAddress -AddressFamily IPv4 | Where-Object { (\$_.IPAddress -like '192.168.*' -or \$_.IPAddress -like '10.*') -and \$_.IPAddress -notlike '192.168.56.*' -and \$_.InterfaceAlias -notlike '*Virtual*' -and \$_.InterfaceAlias -notlike '*vEthernet*' } | Select-Object -ExpandProperty IPAddress -First 1)" 2>/dev/null | tr -d '\r\n')
            if [ -n "$WIN_IP" ]; then
                LOCAL_IP="$WIN_IP"
            fi
        fi
    fi

    # 2. Se não detectou ou está em Linux nativo, usa hostname -I
    if [ -z "$LOCAL_IP" ]; then
        if command -v hostname >/dev/null 2>&1; then
            LOCAL_IP=$(hostname -I 2>/dev/null | awk '{print $1}')
        fi

        if [ -z "$LOCAL_IP" ] || [ "$LOCAL_IP" = "127.0.0.1" ]; then
            if command -v ip >/dev/null 2>&1; then
                LOCAL_IP=$(ip route get 1.1.1.1 2>/dev/null | grep -oP 'src \K\S+')
            fi
        fi
    fi

    # 3. Fallback
    if [ -z "$LOCAL_IP" ]; then
        LOCAL_IP="localhost"
    fi

    mkdir -p src/js
    echo "window.BANCADA_LOCAL_IP = '${LOCAL_IP}';" > src/js/server-info.js
    export HOST_IP="${LOCAL_IP}"
}

load_env_file() {
    if [ -f .env ]; then
        local env_port
        env_port=$(grep -E '^[[:space:]]*STREAMLIT_PORT[[:space:]]*=' .env | tail -n 1 | cut -d '=' -f 2 | tr -d ' "\r\n')
        if [ -n "$env_port" ]; then
            PORT="$env_port"
        fi
    fi
    PORT="${PORT:-8502}"
}

open_browser() {
    local port="${1:-$PORT}"
    local url="http://localhost:${port}/"
    if [ -f /proc/sys/fs/binfmt_misc/WSLInterop ] || [ -n "$WSL_DISTRO_NAME" ]; then
        if command -v cmd.exe >/dev/null 2>&1; then
            cmd.exe /c start "$url" >/dev/null 2>&1 &
            return
        elif command -v powershell.exe >/dev/null 2>&1; then
            powershell.exe -NoProfile -Command "Start-Process '$url'" >/dev/null 2>&1 &
            return
        fi
    fi

    if command -v xdg-open >/dev/null 2>&1; then
        xdg-open "$url" >/dev/null 2>&1 &
    elif command -v gnome-open >/dev/null 2>&1; then
        gnome-open "$url" >/dev/null 2>&1 &
    fi
}

cleanup() {
    trap - INT TERM EXIT
    # Restaura cursor caso tenha sido ocultado
    tput cnorm 2>/dev/null
    # Desativa scrolling region caso ativa
    local rows
    rows=$(tput lines 2>/dev/null || echo 24)
    printf "\033[1;%dr" "$rows" 2>/dev/null
    printf "\033[%d;1H\n" "$rows" 2>/dev/null

    echo ""
    echo -e "${C_YELLOW}Encerrando containers do Sistema Bancada...${C_RESET}"
    if command -v docker >/dev/null 2>&1; then
        docker compose down
    fi
    echo -e "${C_GREEN}[OK] Servidor e containers finalizados com sucesso.${C_RESET}"
    exit 0
}

# Desenha a barra fixa no rodapé da janela do terminal
render_bottom_toolbar() {
    local rows
    rows=$(tput lines 2>/dev/null || echo 24)
    local cols
    cols=$(tput cols 2>/dev/null || echo 80)
    local sep_len=$(( cols - 2 ))
    [ $sep_len -lt 10 ] && sep_len=70

    # Salva posição do cursor e desce para as duas últimas linhas
    tput sc 2>/dev/null
    printf "\033[%d;1H" $(( rows - 1 ))
    echo -ne "${C_GRAY}─${C_RESET}"
    printf "${C_GRAY}%%0.s─${C_RESET}" $(seq 1 $sep_len)
    printf "\033[K\n"
    printf " ${C_BOLD}[c]${C_RESET} ${C_YELLOW}Limpar Logs${C_RESET} | ${C_BOLD}[r]${C_RESET} ${C_CYAN}Reiniciar Web${C_RESET} | ${C_BOLD}[b]${C_RESET} ${C_GREEN}Navegador${C_RESET} | ${C_BOLD}[q]${C_RESET} ${C_RED}Encerrar${C_RESET}\033[K"
    tput rc 2>/dev/null
}

setup_terminal_split() {
    local rows
    rows=$(tput lines 2>/dev/null || echo 24)
    # Define a região de rolagem (scrolling region) da linha 1 até (rows - 2)
    # Assim, os logs nunca sobrescrevem as duas linhas de baixo!
    printf "\033[1;%dr" $(( rows - 2 )) 2>/dev/null
    render_bottom_toolbar
}

stream_interactive_logs() {
    local LOG_PID=""

    # Configura a tela dividida com barra fixa no rodapé
    setup_terminal_split

    # Trata redimensionamento de janela (SIGWINCH)
    trap 'setup_terminal_split' WINCH

    # Inicia streaming dos logs do container dentro da área de rolagem
    docker compose logs -f --tail=100 web &
    LOG_PID=$!

    # Loop de escuta de comandos de teclado em tempo real
    while kill -0 "$LOG_PID" 2>/dev/null; do
        if read -r -s -n 1 -t 1 key; then
            case "$key" in
                c|C)
                    clear
                    setup_terminal_split
                    ;;
                r|R)
                    kill "$LOG_PID" 2>/dev/null
                    clear
                    echo -e "${C_YELLOW}Reiniciando serviço web...${C_RESET}"
                    docker compose restart web
                    clear
                    setup_terminal_split
                    docker compose logs -f --tail=50 web &
                    LOG_PID=$!
                    ;;
                b|B)
                    open_browser "$PORT"
                    ;;
                q|Q)
                    kill "$LOG_PID" 2>/dev/null
                    cleanup
                    ;;
            esac
        fi
    done

    wait "$LOG_PID" 2>/dev/null
}

start_system() {
    local force_build="${1:-false}"
    load_env_file
    detect_ip

    mkdir -p ~/.local/share/python_keyring
    touch chamados.db

    clear
    echo -e "${C_CYAN}Iniciando Sistema Bancada (Porta: ${PORT} | IP: ${HOST_IP})...${C_RESET}"

    if [ "$force_build" = true ]; then
        echo -e "${C_YELLOW}Construindo imagem Docker...${C_RESET}"
        docker compose build web
    fi

    # Inicia o container diretamente sem auto-build desnecessário
    if ! docker compose up -d web; then
        echo ""
        echo -e "${C_RED}[ERRO] Falha ao iniciar o container web do Sistema Bancada.${C_RESET}"
        exit 1
    fi

    open_browser "$PORT"
    trap cleanup INT TERM EXIT

    stream_interactive_logs
}

config_senhas() {
    load_env_file
    detect_ip
    clear
    echo -e "${C_CYAN}Abrindo assistente de senhas no container interativo...${C_RESET}"
    docker compose run --rm web python src/salvar_senha.py
    echo ""
    echo -e "${C_GREEN}Assistente finalizado!${C_RESET}"
    read -p "Pressione ENTER para voltar ao menu..." dummy
}

run_orquestrador() {
    load_env_file
    detect_ip
    clear
    echo -e "${C_CYAN}Executando orquestrador de chamados e sincronização...${C_RESET}"
    docker compose run --rm web python orquestrador.py
    echo ""
    echo -e "${C_GREEN}Execução do orquestrador finalizada!${C_RESET}"
    read -p "Pressione ENTER para voltar ao menu..." dummy
}

rebuild_docker() {
    clear
    echo -e "${C_MAGENTA}Reconstruindo imagens Docker Compose (--no-cache)...${C_RESET}"
    docker compose build --no-cache web
    echo ""
    echo -e "${C_GREEN}Rebuild concluído com sucesso!${C_RESET}"
    read -p "Pressione ENTER para voltar ao menu..." dummy
}

stop_system() {
    echo -e "${C_YELLOW}Encerrando todos os containers do Sistema Bancada...${C_RESET}"
    docker compose down
    echo -e "${C_GREEN}Containers encerrados com sucesso!${C_RESET}"
}

show_menu() {
    clear
    echo -e "${C_CYAN}${C_BOLD}╔══════════════════════════════════════════════════════════════╗${C_RESET}"
    echo -e "${C_CYAN}${C_BOLD}║         SISTEMA BANCADA — AUTOMAÇÃO DE CHAMADOS STI          ║${C_RESET}"
    echo -e "${C_CYAN}${C_BOLD}║                                                              ║${C_RESET}"
    echo -e "${C_CYAN}${C_BOLD}║${C_RESET}  ${C_BOLD}Escolha uma opção:${C_RESET}                                          ${C_CYAN}${C_BOLD}║${C_RESET}"
    echo -e "${C_CYAN}${C_BOLD}║${C_RESET}  ${C_GREEN}1${C_RESET} - Iniciar Sistema Bancada (Streamlit Dashboard)           ${C_CYAN}${C_BOLD}║${C_RESET}"
    echo -e "${C_CYAN}${C_BOLD}║${C_RESET}  ${C_GREEN}2${C_RESET} - Configurar Senhas (Cofre Keyring)                       ${C_CYAN}${C_BOLD}║${C_RESET}"
    echo -e "${C_CYAN}${C_BOLD}║${C_RESET}  ${C_GREEN}3${C_RESET} - Executar Orquestrador de Sincronização                  ${C_CYAN}${C_BOLD}║${C_RESET}"
    echo -e "${C_CYAN}${C_BOLD}║${C_RESET}  ${C_GREEN}4${C_RESET} - Reconstruir Docker Compose (--no-cache)                 ${C_CYAN}${C_BOLD}║${C_RESET}"
    echo -e "${C_CYAN}${C_BOLD}║${C_RESET}  ${C_GREEN}5${C_RESET} - Parar sistema (docker compose down)                     ${C_CYAN}${C_BOLD}║${C_RESET}"
    echo -e "${C_CYAN}${C_BOLD}║${C_RESET}  ${C_RED}0${C_RESET} - Sair                                                    ${C_CYAN}${C_BOLD}║${C_RESET}"
    echo -e "${C_CYAN}${C_BOLD}╚══════════════════════════════════════════════════════════════╝${C_RESET}"
    echo ""
    read -p "Opção [0-5]: " opcao
    case "$opcao" in
        1) start_system false ;;
        2) config_senhas; show_menu ;;
        3) run_orquestrador; show_menu ;;
        4) rebuild_docker; show_menu ;;
        5) stop_system ;;
        0) exit 0 ;;
        *) echo -e "${C_RED}Opção inválida.${C_RESET}"; sleep 1; show_menu ;;
    esac
}

case "$1" in
    --start|-s)
        start_system false
        ;;
    --build|-b)
        start_system true
        ;;
    --config-senhas|--senhas|-p)
        config_senhas
        ;;
    --orquestrador|-o)
        run_orquestrador
        ;;
    --rebuild|-r)
        rebuild_docker
        ;;
    --down|--stop|-d)
        stop_system
        ;;
    --help|-h)
        echo "Uso: ./00-iniciar.sh [OPÇÃO]"
        echo ""
        echo "Opções:"
        echo "  --start, -s              Inicia o Sistema Bancada (Dashboard)"
        echo "  --build, -b              Reconstrói a imagem e inicia"
        echo "  --config-senhas, -p      Abre o configurador de senhas do cofre"
        echo "  --orquestrador, -o       Executa a rotina do orquestrador"
        echo "  --rebuild, -r            Reconstrói a imagem Docker (--no-cache)"
        echo "  --down, -d               Para os containers do sistema"
        echo "  --help, -h               Exibe esta ajuda"
        echo "  (sem argumentos)         Abre o menu interativo"
        ;;
    *)
        show_menu
        ;;
esac
