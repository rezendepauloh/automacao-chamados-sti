# -*- coding: utf-8 -*-
"""
Script de Inicialização do Sistema Bancada
Detecta IP da rede, gera arquivo de configuração JS, exibe banner colorido com QR Code e inicia o Streamlit.
"""
import os
import sys
import signal
import socket
import subprocess
from pathlib import Path

# Adiciona raiz ao sys.path
ROOT_DIR = Path(__file__).parent.resolve()
sys.path.insert(0, str(ROOT_DIR))
sys.path.insert(0, str(ROOT_DIR / "src"))

# Cores ANSI
RESET = "\033[0m"
BOLD = "\033[1m"
DIM = "\033[2m"
YELLOW = "\033[33m"
CYAN = "\033[36m"
GREEN = "\033[32m"
RED = "\033[31m"
WHITE = "\033[37m"
DARK_GRAY = "\033[90m"

PORT = int(os.getenv("STREAMLIT_SERVER_PORT", "8501"))

def get_local_ip() -> str:
    """Detecta o IP local IPv4 acessível na rede corporativa/Wi-Fi."""
    # 1. Verifica variável de ambiente HOST_IP (repassada pelo Docker/Compose)
    env_ip = os.getenv("HOST_IP", "").strip()
    if env_ip and env_ip != "localhost" and env_ip != "127.0.0.1":
        return env_ip
    
    # 2. Detecção dinâmica via socket
    try:
        s = socket.socket(socket.AF_INET, socket.SOCK_DGRAM)
        s.connect(("1.1.1.1", 80))
        ip = s.getsockname()[0]
        s.close()
        if ip and not ip.startswith("127."):
            return ip
    except Exception:
        pass
    
    return "localhost"

def setup_environment(local_ip: str) -> None:
    """Configura arquivos de suporte pré-inicialização."""
    # Garante criação da pasta js e do arquivo server-info.js
    js_dir = ROOT_DIR / "src" / "js"
    js_dir.mkdir(parents=True, exist_ok=True)
    server_info_file = js_dir / "server-info.js"
    server_info_file.write_text(f"window.BANCADA_LOCAL_IP = '{local_ip}';\n", encoding="utf-8")
    
    # Garante arquivo do banco de dados SQLite local
    db_file = ROOT_DIR / "chamados.db"
    if not db_file.exists():
        db_file.touch()

def print_banner(local_ip: str, port: int) -> None:
    """Exibe o banner estilizado com status e QR Code em ASCII."""
    full_url = f"http://{local_ip}:{port}/"
    local_url = f"http://localhost:{port}/"
    db_mode = os.getenv("DB_TYPE", "sqlite").upper()

    print("")
    print(f"{CYAN}╔══════════════════════════════════════════════════════════════════════╗{RESET}")
    print(f"{CYAN}║                    S I S T E M A   B A N C A D A                     ║{RESET}")
    print(f"{CYAN}║             Automação de Chamados STI — Streamlit Dashboard          ║{RESET}")
    print(f"{CYAN}╚══════════════════════════════════════════════════════════════════════╝{RESET}")
    print(f" {CYAN}●{RESET} {BOLD}Banco de Dados{RESET}       : {GREEN}{db_mode} (chamados.db centralizado){RESET}")
    print(f" {CYAN}●{RESET} {BOLD}Porta do Servidor{RESET}    : {WHITE}{port}{RESET}")
    print(f" {CYAN}●{RESET} {BOLD}Acesso no Computador{RESET} : {CYAN}{local_url}{RESET}")
    print(f" {CYAN}●{RESET} {BOLD}Acesso Celular/Tablet{RESET}: {GREEN}{full_url}{RESET}")
    print(f"{DARK_GRAY}{'─' * 72}{RESET}")

    # Geração de QR Code ASCII
    try:
        import qrcode
        qr = qrcode.QRCode(border=1)
        qr.add_data(full_url)
        qr.make(fit=True)
        print(f" {CYAN}●{RESET} {BOLD}Aponte a câmera do Celular para o QR Code abaixo:{RESET}\n")
        qr.print_ascii(invert=True)
    except Exception:
        pass

    print(f"{DARK_GRAY}{'─' * 72}{RESET}")
    print(f" {DIM}Pressione CTRL+C para encerrar o servidor.{RESET}")
    print("")

def main() -> None:
    local_ip = get_local_ip()
    setup_environment(local_ip)
    print_banner(local_ip, PORT)

    # Executa o Streamlit
    cmd = [
        sys.executable,
        "-m",
        "streamlit",
        "run",
        "dashboard.py",
        "--server.port",
        str(PORT),
        "--server.address",
        "0.0.0.0",
        "--server.headless=true",
        "--logger.level=error"
    ]

    process = None
    try:
        process = subprocess.Popen(cmd)
        process.wait()
    except (KeyboardInterrupt, SystemExit):
        print(f"\n{CYAN}══════════════════════════════════════════════════════════════════════{RESET}")
        print(f" {YELLOW}[!] Encerrando servidor Streamlit...{RESET}")
        print(f"{CYAN}══════════════════════════════════════════════════════════════════════{RESET}")
        if process:
            process.terminate()
            try:
                process.wait(timeout=5)
            except subprocess.TimeoutExpired:
                process.kill()
        print(f" {GREEN}[OK] Servidor finalizado com sucesso.{RESET}\n")
        sys.exit(0)

if __name__ == "__main__":
    main()
