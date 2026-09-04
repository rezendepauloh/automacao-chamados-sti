import pandas as pd
from .connection import get_connection

def setup_notifications_table():
    """Cria a tabela de notificações se não existir."""
    conn = get_connection()
    cursor = conn.cursor()
    cursor.execute("""
    CREATE TABLE IF NOT EXISTS notificacoes (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        tipo TEXT NOT NULL,
        titulo TEXT NOT NULL,
        mensagem TEXT NOT NULL,
        data_evento TEXT,
        link_pagina TEXT,
        lida INTEGER DEFAULT 0,
        data_criacao DATETIME DEFAULT CURRENT_TIMESTAMP,
        UNIQUE(tipo, titulo, data_evento) ON CONFLICT IGNORE
    )
    """)
    conn.commit()
    conn.close()

def add_notification(tipo: str, titulo: str, mensagem: str, data_evento: str = "", link_pagina: str = "") -> bool:
    """Adiciona uma notificação no banco de dados. Retorna True se foi inserida uma nova notificação."""
    setup_notifications_table()
    conn = get_connection()
    cursor = conn.cursor()
    try:
        cursor.execute("""
        INSERT INTO notificacoes (tipo, titulo, mensagem, data_evento, link_pagina, lida)
        VALUES (?, ?, ?, ?, ?, 0)
        """, (tipo, titulo, mensagem, data_evento, link_pagina))
        inserted = cursor.rowcount > 0
        conn.commit()
        return inserted
    except Exception:
        return False
    finally:
        conn.close()

def get_notifications(only_unread: bool = False, limit: int = 100) -> pd.DataFrame:
    """Retorna um DataFrame pandas com as notificações do banco."""
    setup_notifications_table()
    conn = get_connection()
    query = "SELECT * FROM notificacoes"
    if only_unread:
        query += " WHERE lida = 0"
    query += " ORDER BY id DESC LIMIT ?"
    df = pd.read_sql_query(query, conn, params=(limit,))
    conn.close()
    return df

def get_unread_notifications_count() -> int:
    """Retorna a quantidade de notificações não lidas."""
    setup_notifications_table()
    conn = get_connection()
    cursor = conn.cursor()
    cursor.execute("SELECT COUNT(*) FROM notificacoes WHERE lida = 0")
    count = cursor.fetchone()[0]
    conn.close()
    return count

def mark_notification_as_read(notif_id: int):
    """Marca uma notificação específica como lida."""
    setup_notifications_table()
    conn = get_connection()
    cursor = conn.cursor()
    cursor.execute("UPDATE notificacoes SET lida = 1 WHERE id = ?", (notif_id,))
    conn.commit()
    conn.close()

def mark_notification_as_unread(notif_id: int):
    """Marca uma notificação específica como não lida."""
    setup_notifications_table()
    conn = get_connection()
    cursor = conn.cursor()
    cursor.execute("UPDATE notificacoes SET lida = 0 WHERE id = ?", (notif_id,))
    conn.commit()
    conn.close()

def mark_all_notifications_as_read():
    """Marca todas as notificações como lidas."""
    setup_notifications_table()
    conn = get_connection()
    cursor = conn.cursor()
    cursor.execute("UPDATE notificacoes SET lida = 1 WHERE lida = 0")
    conn.commit()
    conn.close()
