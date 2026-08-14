from datetime import datetime
import pandas as pd
from .connection import get_connection

def setup_unidades_tables():
    """Cria as tabelas de unidades unificadas e unidades manuais no SQLite se não existirem."""
    conn = get_connection()
    cursor = conn.cursor()
    cursor.execute("""
    CREATE TABLE IF NOT EXISTS unidades_manuais (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        cidade TEXT,
        tipo TEXT,
        setor TEXT,
        sigla TEXT,
        titular TEXT,
        unidade_predio TEXT,
        telefone TEXT,
        url TEXT,
        data_atualizacao TEXT
    )
    """)
    cursor.execute("""
    CREATE TABLE IF NOT EXISTS unidades (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        cidade TEXT,
        tipo TEXT,
        setor TEXT,
        sigla TEXT,
        titular TEXT,
        unidade_predio TEXT,
        telefone TEXT,
        url TEXT,
        origem TEXT DEFAULT 'web',
        data_atualizacao TEXT
    )
    """)
    conn.commit()
    conn.close()

def setup_unidades_manuais_table():
    """Cria a tabela de unidades manuais no SQLite se não existir."""
    setup_unidades_tables()

def get_unidades_manuais() -> pd.DataFrame:
    """Retorna todas as unidades manuais cadastradas no SQLite."""
    setup_unidades_tables()
    conn = get_connection()
    try:
        df = pd.read_sql_query("SELECT * FROM unidades_manuais ORDER BY id ASC", conn)
        conn.close()
        return df
    except Exception:
        conn.close()
        return pd.DataFrame()

def save_unidades_to_db(df: pd.DataFrame):
    """Salva a lista completa unificada de unidades (Web + Manuais) na tabela 'unidades' do SQLite."""
    setup_unidades_tables()
    if df is None or df.empty:
        return

    conn = get_connection()
    cursor = conn.cursor()
    now_str = datetime.now().strftime("%Y-%m-%d %H:%M:%S")

    cursor.execute("DELETE FROM unidades")

    for _, row in df.iterrows():
        cidade = str(row.get('Cidade', row.get('cidade', ''))).strip()
        tipo = str(row.get('Tipo', row.get('tipo', ''))).strip()
        setor = str(row.get('Setor', row.get('setor', ''))).strip()
        sigla = str(row.get('Sigla', row.get('sigla', ''))).strip()
        titular = str(row.get('Titular', row.get('titular', ''))).strip()
        u_predio = str(row.get('Unidade (Prédio)', row.get('unidade_predio', ''))).strip()
        telefone = str(row.get('Telefone', row.get('telefone', ''))).strip()
        url = str(row.get('URL', row.get('url', ''))).strip()

        if not setor and not cidade:
            continue

        cursor.execute("""
        INSERT INTO unidades (cidade, tipo, setor, sigla, titular, unidade_predio, telefone, url, data_atualizacao)
        VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?)
        """, (cidade, tipo, setor, sigla, titular, u_predio, telefone, url, now_str))

    conn.commit()
    conn.close()

def get_unidades_df() -> pd.DataFrame:
    """Retorna o DataFrame de todas as unidades (Promotorias, Procuradorias e Setores Manuais) salvas no SQLite com identificador de origem."""
    setup_unidades_tables()
    conn = get_connection()
    try:
        query = """
        SELECT 
            u.id,
            u.cidade AS Cidade, 
            u.tipo AS Tipo, 
            u.setor AS Setor, 
            u.sigla AS Sigla, 
            u.titular AS Titular, 
            u.unidade_predio AS 'Unidade (Prédio)', 
            u.telefone AS Telefone, 
            u.url AS URL,
            um.id AS manual_id,
            CASE WHEN um.id IS NOT NULL THEN '📌 Manual' ELSE '🌐 Portal Web' END AS Origem
        FROM unidades u
        LEFT JOIN unidades_manuais um ON u.cidade = um.cidade AND u.tipo = um.tipo AND u.setor = um.setor
        ORDER BY u.cidade ASC, u.setor ASC
        """
        df = pd.read_sql_query(query, conn)
        conn.close()
        return df
    except Exception:
        conn.close()
        return pd.DataFrame()

def add_unidade_manual(cidade: str, tipo: str, setor: str, sigla: str = "", titular: str = "", unidade_predio: str = "", telefone: str = "", url: str = ""):
    """Insere ou atualiza uma unidade manual no SQLite e sincroniza a tabela unificada 'unidades'."""
    setup_unidades_tables()
    conn = get_connection()
    cursor = conn.cursor()
    now_str = datetime.now().strftime("%Y-%m-%d %H:%M:%S")

    cursor.execute("SELECT id FROM unidades_manuais WHERE cidade = ? AND tipo = ? AND setor = ?", (cidade, tipo, setor))
    existing = cursor.fetchone()

    if existing:
        cursor.execute("""
        UPDATE unidades_manuais 
        SET sigla = ?, titular = ?, unidade_predio = ?, telefone = ?, url = ?, data_atualizacao = ?
        WHERE id = ?
        """, (sigla, titular, unidade_predio, telefone, url, now_str, existing[0]))
    else:
        cursor.execute("""
        INSERT INTO unidades_manuais (cidade, tipo, setor, sigla, titular, unidade_predio, telefone, url, data_atualizacao)
        VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?)
        """, (cidade, tipo, setor, sigla, titular, unidade_predio, telefone, url, now_str))

    cursor.execute("SELECT id FROM unidades WHERE cidade = ? AND tipo = ? AND setor = ?", (cidade, tipo, setor))
    existing_uni = cursor.fetchone()
    if existing_uni:
        cursor.execute("""
        UPDATE unidades 
        SET sigla = ?, titular = ?, unidade_predio = ?, telefone = ?, url = ?, data_atualizacao = ?
        WHERE id = ?
        """, (sigla, titular, unidade_predio, telefone, url, now_str, existing_uni[0]))
    else:
        cursor.execute("""
        INSERT INTO unidades (cidade, tipo, setor, sigla, titular, unidade_predio, telefone, url, data_atualizacao)
        VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?)
        """, (cidade, tipo, setor, sigla, titular, unidade_predio, telefone, url, now_str))

    conn.commit()
    conn.close()

def update_unidade_manual_by_id(manual_id: int, cidade: str, tipo: str, setor: str, sigla: str = "", titular: str = "", unidade_predio: str = "", telefone: str = "", url: str = ""):
    """Atualiza uma unidade manual pelo seu manual_id no SQLite nas duas tabelas."""
    setup_unidades_tables()
    conn = get_connection()
    cursor = conn.cursor()
    now_str = datetime.now().strftime("%Y-%m-%d %H:%M:%S")

    cursor.execute("SELECT cidade, tipo, setor FROM unidades_manuais WHERE id = ?", (manual_id,))
    old_row = cursor.fetchone()
    if old_row:
        old_cidade, old_tipo, old_setor = old_row[0], old_row[1], old_row[2]
        cursor.execute("""
        UPDATE unidades
        SET cidade = ?, tipo = ?, setor = ?, sigla = ?, titular = ?, unidade_predio = ?, telefone = ?, url = ?, data_atualizacao = ?
        WHERE cidade = ? AND tipo = ? AND setor = ?
        """, (cidade, tipo, setor, sigla, titular, unidade_predio, telefone, url, now_str, old_cidade, old_tipo, old_setor))

    cursor.execute("""
    UPDATE unidades_manuais 
    SET cidade = ?, tipo = ?, setor = ?, sigla = ?, titular = ?, unidade_predio = ?, telefone = ?, url = ?, data_atualizacao = ?
    WHERE id = ?
    """, (cidade, tipo, setor, sigla, titular, unidade_predio, telefone, url, now_str, manual_id))

    conn.commit()
    conn.close()

def delete_unidade_manual(unit_id: int):
    """Deleta uma unidade manual do SQLite (das duas tabelas) pelo ID."""
    setup_unidades_tables()
    conn = get_connection()
    cursor = conn.cursor()

    cursor.execute("SELECT cidade, tipo, setor FROM unidades_manuais WHERE id = ?", (unit_id,))
    row = cursor.fetchone()
    if row:
        cidade, tipo, setor = row[0], row[1], row[2]
        cursor.execute("DELETE FROM unidades WHERE cidade = ? AND tipo = ? AND setor = ?", (cidade, tipo, setor))

    cursor.execute("DELETE FROM unidades_manuais WHERE id = ?", (unit_id,))
    conn.commit()
    conn.close()
