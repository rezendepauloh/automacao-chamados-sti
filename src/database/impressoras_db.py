import pandas as pd
from datetime import datetime
from .connection import get_connection

def setup_impressoras_table():
    """Cria a tabela de impressoras/dispositivos do PaperCut se não existir."""
    conn = get_connection()
    cursor = conn.cursor()
    cursor.execute("""
    CREATE TABLE IF NOT EXISTS impressoras (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        nome TEXT UNIQUE,
        servidor TEXT,
        tipo TEXT,
        modelo TEXT,
        localizacao TEXT,
        ip_host TEXT,
        status TEXT,
        total_paginas INTEGER DEFAULT 0,
        filas_relacionadas TEXT,
        detalhes_extra TEXT,
        data_atualizacao TEXT
    )
    """)
    conn.commit()
    conn.close()

def save_impressoras_to_db(df: pd.DataFrame):
    """
    Salva ou atualiza os dados de impressoras/dispositivos do PaperCut no banco de dados.
    """
    setup_impressoras_table()
    if df is None or df.empty:
        return

    conn = get_connection()
    cursor = conn.cursor()
    now_str = datetime.now().strftime("%Y-%m-%d %H:%M:%S")

    cursor.execute("DELETE FROM impressoras")

    for _, row in df.iterrows():
        nome = str(row.get('nome', '')).strip()
        if not nome or nome.lower() in ['nan', 'none', '']:
            continue

        servidor = str(row.get('servidor', '')).strip()
        tipo = str(row.get('tipo', 'Impressora')).strip()
        modelo = str(row.get('modelo', '')).strip()
        localizacao = str(row.get('localizacao', '')).strip()
        ip_host = str(row.get('ip_host', '')).strip()
        status = str(row.get('status', 'OK')).strip()
        
        try:
            total_paginas = int(row.get('total_paginas', 0))
        except (ValueError, TypeError):
            total_paginas = 0

        filas_relacionadas = str(row.get('filas_relacionadas', '')).strip()
        detalhes_extra = str(row.get('detalhes_extra', '')).strip()

        cursor.execute("""
        INSERT INTO impressoras (
            nome, servidor, tipo, modelo, localizacao, ip_host, status, 
            total_paginas, filas_relacionadas, detalhes_extra, data_atualizacao
        ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
        ON CONFLICT(nome) DO UPDATE SET
            servidor=excluded.servidor,
            tipo=excluded.tipo,
            modelo=excluded.modelo,
            localizacao=excluded.localizacao,
            ip_host=excluded.ip_host,
            status=excluded.status,
            total_paginas=excluded.total_paginas,
            filas_relacionadas=excluded.filas_relacionadas,
            detalhes_extra=excluded.detalhes_extra,
            data_atualizacao=excluded.data_atualizacao
        """, (
            nome, servidor, tipo, modelo, localizacao, ip_host, status,
            total_paginas, filas_relacionadas, detalhes_extra, now_str
        ))

    conn.commit()
    conn.close()

def get_impressoras_df() -> pd.DataFrame:
    """
    Retorna o DataFrame contendo todas as impressoras/dispositivos do PaperCut cadastrados no banco.
    """
    setup_impressoras_table()
    conn = get_connection()
    df = pd.read_sql_query("SELECT * FROM impressoras ORDER BY nome ASC", conn)
    conn.close()
    return df
