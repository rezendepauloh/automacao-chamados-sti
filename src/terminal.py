# -*- coding: utf-8 -*-
"""
Utilitário centralizado de cores ANSI e formatação de logs para o terminal.
Garante suporte a cores ANSI e codificação UTF-8 (incluindo emojis) em Windows e Linux.
"""
import os
import sys

# 1. Inicialização do Terminal (Suporte ANSI + UTF-8)
if sys.platform == 'win32':
    os.system('')  # Ativa suporte a sequências ANSI no prompt do Windows
    if hasattr(sys.stdout, 'reconfigure'):
        sys.stdout.reconfigure(encoding='utf-8', errors='replace')
    if hasattr(sys.stderr, 'reconfigure'):
        sys.stderr.reconfigure(encoding='utf-8', errors='replace')

# 2. Códigos ANSI de Formatação e Cores
RESET = "\033[0m"
BOLD = "\033[1m"
DIM = "\033[2m"
YELLOW = "\033[33m"
CYAN = "\033[36m"
GREEN = "\033[32m"
RED = "\033[31m"
WHITE = "\033[37m"
DARK_GRAY = "\033[90m"

# 3. Função `log(msg, symbol="ℹ️", color=WHITE)` Padronizada
def log(msg: str, symbol: str = "ℹ️", color: str = WHITE):
    """
    Exibe mensagem formatada no terminal com símbolo e cor ANSI,
    protegida contra erros de codificação Unicode.
    """
    try:
        print(f"[{symbol}] {color}{msg}{RESET}")
    except UnicodeEncodeError:
        safe_msg = str(msg).encode('ascii', errors='replace').decode('ascii')
        safe_sym = str(symbol).encode('ascii', errors='replace').decode('ascii')
        print(f"[{safe_sym}] {color}{safe_msg}{RESET}")

# 4. Utilitários Visuais de Cabeçalhos e Caixas
def print_header(title: str, color: str = CYAN, width: int = 65):
    """Exibe um cabeçalho estilizado com caixa no terminal."""
    border = "═" * (width - 2)
    padding = width - 4
    centered_title = title.center(padding)
    print(f"\n{color}╔{border}╗{RESET}")
    print(f"{color}║ {BOLD}{centered_title}{RESET}{color} ║{RESET}")
    print(f"{color}╚{border}╝{RESET}\n")

def print_section(title: str, color: str = YELLOW):
    """Exibe um divisor de seção simples no terminal."""
    print(f"\n{color}──── {BOLD}{title}{RESET} {color}{'─' * (50 - len(title))}{RESET}")

def print_box_row(content: str, color: str = WHITE, width: int = 65):
    """Exibe uma linha dentro de moldura visual."""
    padding = width - 4
    print(f"{color}║{RESET} {content:<{padding}} {color}║{RESET}")
