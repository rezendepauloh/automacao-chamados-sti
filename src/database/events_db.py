from datetime import datetime
import pandas as pd
from .connection import get_connection

def setup_eventos_manuais_table():
    """Cria a tabela de eventos manuais no banco de dados SQLite se não existir."""
    conn = get_connection()
    cursor = conn.cursor()
    cursor.execute("""
    CREATE TABLE IF NOT EXISTS eventos_manuais (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        titulo TEXT NOT NULL,
        data_inicio TEXT NOT NULL,
        data_fim TEXT,
        descricao TEXT,
        autor TEXT,
        data_criacao TEXT
    )
    """)
    conn.commit()
    conn.close()

def save_evento_manual(titulo: str, data_inicio: str, data_fim: str = "", descricao: str = "", autor: str = "Bancada STI"):
    """Insere um novo evento manual na tabela eventos_manuais."""
    setup_eventos_manuais_table()
    conn = get_connection()
    cursor = conn.cursor()
    now_str = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    cursor.execute("""
    INSERT INTO eventos_manuais (titulo, data_inicio, data_fim, descricao, autor, data_criacao)
    VALUES (?, ?, ?, ?, ?, ?)
    """, (titulo, data_inicio, data_fim, descricao, autor, now_str))
    conn.commit()
    conn.close()

def get_eventos_manuais() -> pd.DataFrame:
    """Retorna um DataFrame com todos os eventos manuais cadastrados no SQLite."""
    setup_eventos_manuais_table()
    conn = get_connection()
    try:
        df = pd.read_sql_query("SELECT * FROM eventos_manuais ORDER BY id DESC", conn)
        conn.close()
        return df
    except Exception:
        conn.close()
        return pd.DataFrame()
