import pandas as pd
from datetime import datetime
from .connection import get_connection

def setup_ramais_table():
    """Cria a tabela ramais_mpms se não existir."""
    conn = get_connection()
    cursor = conn.cursor()
    cursor.execute("""
    CREATE TABLE IF NOT EXISTS ramais_mpms (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        localidade TEXT,
        setor_nome TEXT,
        telefone_ramal TEXT,
        tipo TEXT,
        data_atualizacao TEXT
    )
    """)
    conn.commit()
    conn.close()

def save_ramais_to_db(df: pd.DataFrame):
    """Limpa a tabela ramais_mpms e insere os dados do DataFrame recebido."""
    setup_ramais_table()
    conn = get_connection()
    cursor = conn.cursor()
    cursor.execute("DELETE FROM ramais_mpms")
    conn.commit()
    
    if not df.empty:
        df_to_save = df.copy()
        if "data_atualizacao" not in df_to_save.columns:
            df_to_save["data_atualizacao"] = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
            
        cols = ["localidade", "setor_nome", "telefone_ramal", "tipo", "data_atualizacao"]
        cols_present = [c for c in cols if c in df_to_save.columns]
        df_to_save[cols_present].to_sql("ramais_mpms", conn, if_exists="append", index=False)
        conn.commit()
    conn.close()

def get_ramais_df() -> pd.DataFrame:
    """Retorna os dados da tabela ramais_mpms em um DataFrame."""
    setup_ramais_table()
    conn = get_connection()
    try:
        df = pd.read_sql_query("SELECT id, localidade, setor_nome, telefone_ramal, tipo, data_atualizacao FROM ramais_mpms", conn)
    except Exception:
        df = pd.DataFrame(columns=["id", "localidade", "setor_nome", "telefone_ramal", "tipo", "data_atualizacao"])
    conn.close()
    return df


def setup_ramais_config_table():
    """Cria a tabela ramais_config se não existir."""
    conn = get_connection()
    cursor = conn.cursor()
    cursor.execute("""
    CREATE TABLE IF NOT EXISTS ramais_config (
        chave TEXT PRIMARY KEY,
        valor TEXT
    )
    """)
    conn.commit()
    conn.close()


def get_ramais_config() -> dict:
    """Retorna os links configurados para os PDFs de ramais."""
    setup_ramais_config_table()
    conn = get_connection()
    cursor = conn.cursor()
    cursor.execute("SELECT chave, valor FROM ramais_config")
    rows = cursor.fetchall()
    conn.close()
    config = {k: v for k, v in rows}
    return {
        "Interior": config.get("url_interior", "https://www.mpms.mp.br/anexo/MTMzMDYxNDI3NTAwODYzMjkwNDNmYmI5MGYwYjU2ZGE5ZWI5M2ZmN2EwMTQxLTA0MQ"),
        "Capital / PGJ": config.get("url_capital", "https://www.mpms.mp.br/anexo/MTMzMDYxNDE2ODMwOGI3MjcxZWQ2YzhkYjYyODkwOGFlMDRjNTUzYWFmY2ZhLTA0MQ")
    }


def save_ramais_config(url_interior: str, url_capital: str):
    """Salva os links atualizados dos PDFs de ramais."""
    setup_ramais_config_table()
    conn = get_connection()
    cursor = conn.cursor()
    cursor.execute("INSERT OR REPLACE INTO ramais_config (chave, valor) VALUES ('url_interior', ?)", (url_interior.strip(),))
    cursor.execute("INSERT OR REPLACE INTO ramais_config (chave, valor) VALUES ('url_capital', ?)", (url_capital.strip(),))
    conn.commit()
    conn.close()
