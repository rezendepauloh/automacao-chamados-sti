import sqlite3
from pathlib import Path

DB_PATH = Path("chamados.db")

def get_connection():
    """Retorna uma conexão com o banco de dados SQLite."""
    return sqlite3.connect(DB_PATH)
