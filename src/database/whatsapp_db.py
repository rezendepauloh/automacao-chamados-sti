import os
import pandas as pd
from datetime import datetime
from src.database.connection import get_connection, DB_TYPE

DEFAULT_BANCADA_RECIPIENTS = [
    {"nome": "Reginaldo da Silva Bandeira", "telefone": "5567991455446", "ativo": 1},
    {"nome": "Luiz Leonardo Villalba", "telefone": "5567996477799", "ativo": 1},
    {"nome": "Paulo Henrique Gonçalves Rezende", "telefone": "5567992471379", "ativo": 1},
]

def setup_whatsapp_tables():
    """Cria tabelas de destinatários e histórico de disparos do WhatsApp."""
    conn = get_connection()
    cursor = conn.cursor()
    
    if DB_TYPE in ["postgres", "postgresql"]:
        cursor.execute("""
        CREATE TABLE IF NOT EXISTS whatsapp_destinatarios (
            id SERIAL PRIMARY KEY,
            nome VARCHAR(150) NOT NULL,
            telefone VARCHAR(30) UNIQUE NOT NULL,
            ativo INTEGER DEFAULT 1,
            created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP
        );
        """)
        cursor.execute("""
        CREATE TABLE IF NOT EXISTS whatsapp_disparos_log (
            id SERIAL PRIMARY KEY,
            tipo_evento VARCHAR(50) NOT NULL,
            evento_id VARCHAR(100) NOT NULL,
            data_evento VARCHAR(30),
            destinatario VARCHAR(30) NOT NULL,
            mensagem TEXT,
            status VARCHAR(30) DEFAULT 'enviado',
            response_payload TEXT,
            created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP,
            CONSTRAINT uq_whatsapp_disparo UNIQUE (tipo_evento, evento_id, destinatario)
        );
        """)
    else:
        cursor.execute("""
        CREATE TABLE IF NOT EXISTS whatsapp_destinatarios (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            nome TEXT NOT NULL,
            telefone TEXT UNIQUE NOT NULL,
            ativo INTEGER DEFAULT 1,
            created_at DATETIME DEFAULT CURRENT_TIMESTAMP
        );
        """)
        cursor.execute("""
        CREATE TABLE IF NOT EXISTS whatsapp_disparos_log (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            tipo_evento TEXT NOT NULL,
            evento_id TEXT NOT NULL,
            data_evento TEXT,
            destinatario TEXT NOT NULL,
            mensagem TEXT,
            status TEXT DEFAULT 'enviado',
            response_payload TEXT,
            created_at DATETIME DEFAULT CURRENT_TIMESTAMP,
            UNIQUE(tipo_evento, evento_id, destinatario) ON CONFLICT IGNORE
        );
        """)
    conn.commit()
    cursor.close()
    conn.close()

    # Seed destinatários padrão caso não existam
    seed_whatsapp_destinatarios()

def seed_whatsapp_destinatarios():
    """Garante que os 3 membros da bancada estejam cadastrados."""
    conn = get_connection()
    cursor = conn.cursor()
    placeholder = "%s" if DB_TYPE in ["postgres", "postgresql"] else "?"
    
    for item in DEFAULT_BANCADA_RECIPIENTS:
        try:
            if DB_TYPE in ["postgres", "postgresql"]:
                cursor.execute(
                    """
                    INSERT INTO whatsapp_destinatarios (nome, telefone, ativo)
                    VALUES (%s, %s, %s)
                    ON CONFLICT (telefone) DO NOTHING;
                    """,
                    (item["nome"], item["telefone"], item["ativo"])
                )
            else:
                cursor.execute(
                    """
                    INSERT OR IGNORE INTO whatsapp_destinatarios (nome, telefone, ativo)
                    VALUES (?, ?, ?);
                    """,
                    (item["nome"], item["telefone"], item["ativo"])
                )
        except Exception:
            pass
            
    conn.commit()
    cursor.close()
    conn.close()

def get_whatsapp_destinatarios(only_active: bool = True) -> pd.DataFrame:
    """Retorna lista de destinatários cadastrados para alertas de WhatsApp."""
    setup_whatsapp_tables()
    conn = get_connection()
    query = "SELECT id, nome, telefone, ativo, created_at FROM whatsapp_destinatarios"
    if only_active:
        query += " WHERE ativo = 1"
    query += " ORDER BY nome ASC"
    
    df = pd.read_sql_query(query, conn)
    conn.close()
    return df

def add_whatsapp_destinatario(nome: str, telefone: str) -> bool:
    """Adiciona ou reativa um destinatário."""
    setup_whatsapp_tables()
    clean_phone = "".join(filter(str.isdigit, telefone))
    if not clean_phone.startswith("55"):
        clean_phone = "55" + clean_phone

    conn = get_connection()
    cursor = conn.cursor()
    try:
        if DB_TYPE in ["postgres", "postgresql"]:
            cursor.execute("""
            INSERT INTO whatsapp_destinatarios (nome, telefone, ativo)
            VALUES (%s, %s, 1)
            ON CONFLICT (telefone) DO UPDATE SET nome = EXCLUDED.nome, ativo = 1;
            """, (nome.strip(), clean_phone))
        else:
            cursor.execute("""
            INSERT INTO whatsapp_destinatarios (nome, telefone, ativo)
            VALUES (?, ?, 1)
            ON CONFLICT(telefone) DO UPDATE SET nome = excluded.nome, ativo = 1;
            """, (nome.strip(), clean_phone))
        conn.commit()
        return True
    except Exception:
        return False
    finally:
        cursor.close()
        conn.close()

def toggle_whatsapp_destinatario_status(destinatario_id: int, ativo: bool):
    """Ativa ou desativa um destinatário."""
    setup_whatsapp_tables()
    conn = get_connection()
    cursor = conn.cursor()
    val = 1 if ativo else 0
    if DB_TYPE in ["postgres", "postgresql"]:
        cursor.execute("UPDATE whatsapp_destinatarios SET ativo = %s WHERE id = %s", (val, destinatario_id))
    else:
        cursor.execute("UPDATE whatsapp_destinatarios SET ativo = ? WHERE id = ?", (val, destinatario_id))
    conn.commit()
    cursor.close()
    conn.close()

def has_whatsapp_been_sent(tipo_evento: str, evento_id: str, destinatario: str) -> bool:
    """Verifica se determinado evento já foi enviado para esse destinatário."""
    setup_whatsapp_tables()
    clean_phone = "".join(filter(str.isdigit, destinatario))
    conn = get_connection()
    cursor = conn.cursor()
    
    if DB_TYPE in ["postgres", "postgresql"]:
        cursor.execute("""
        SELECT 1 FROM whatsapp_disparos_log
        WHERE tipo_evento = %s AND evento_id = %s AND destinatario = %s
        LIMIT 1;
        """, (tipo_evento, str(evento_id), clean_phone))
    else:
        cursor.execute("""
        SELECT 1 FROM whatsapp_disparos_log
        WHERE tipo_evento = ? AND evento_id = ? AND destinatario = ?
        LIMIT 1;
        """, (tipo_evento, str(evento_id), clean_phone))
        
    row = cursor.fetchone()
    cursor.close()
    conn.close()
    return bool(row)

def log_whatsapp_dispatch(tipo_evento: str, evento_id: str, data_evento: str, destinatario: str, mensagem: str, status: str = "enviado", response_payload: str = ""):
    """Registra disparo de mensagem do WhatsApp."""
    setup_whatsapp_tables()
    clean_phone = "".join(filter(str.isdigit, destinatario))
    conn = get_connection()
    cursor = conn.cursor()
    
    try:
        if DB_TYPE in ["postgres", "postgresql"]:
            cursor.execute("""
            INSERT INTO whatsapp_disparos_log (tipo_evento, evento_id, data_evento, destinatario, mensagem, status, response_payload)
            VALUES (%s, %s, %s, %s, %s, %s, %s)
            ON CONFLICT (tipo_evento, evento_id, destinatario) DO UPDATE SET
                status = EXCLUDED.status,
                response_payload = EXCLUDED.response_payload,
                created_at = CURRENT_TIMESTAMP;
            """, (tipo_evento, str(evento_id), data_evento, clean_phone, mensagem, status, response_payload))
        else:
            cursor.execute("""
            INSERT OR REPLACE INTO whatsapp_disparos_log (tipo_evento, evento_id, data_evento, destinatario, mensagem, status, response_payload)
            VALUES (?, ?, ?, ?, ?, ?, ?);
            """, (tipo_evento, str(evento_id), data_evento, clean_phone, mensagem, status, response_payload))
        conn.commit()
    except Exception as e:
        print(f"Erro ao registrar log de disparo WhatsApp: {e}")
    finally:
        cursor.close()
        conn.close()

def get_whatsapp_disparos_log(limit: int = 50) -> pd.DataFrame:
    """Retorna os últimos disparos realizados via WhatsApp."""
    setup_whatsapp_tables()
    conn = get_connection()
    query = """
    SELECT id, tipo_evento, evento_id, data_evento, destinatario, mensagem, status, created_at
    FROM whatsapp_disparos_log
    ORDER BY id DESC
    LIMIT ?
    """
    if DB_TYPE in ["postgres", "postgresql"]:
        query = query.replace("?", "%s")
        
    df = pd.read_sql_query(query, conn, params=(limit,))
    conn.close()
    return df
