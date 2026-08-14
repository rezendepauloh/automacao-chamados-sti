import pandas as pd
from .connection import get_connection

def setup_plantoes_tables():
    """Cria as tabelas de plantão matutino e plantão semanal se não existirem."""
    conn = get_connection()
    cursor = conn.cursor()
    
    cursor.execute("""
    CREATE TABLE IF NOT EXISTS plantoes_matutino (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        ano INTEGER,
        data_iso TEXT,
        dia_semana TEXT,
        servidor TEXT,
        telefone TEXT,
        UNIQUE(ano, data_iso, servidor) ON CONFLICT REPLACE
    )
    """)
    
    cursor.execute("""
    CREATE TABLE IF NOT EXISTS plantoes_semanal (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        ano INTEGER,
        mes TEXT,
        periodo_str TEXT,
        data_inicio TEXT,
        data_fim TEXT,
        service_desk TEXT,
        manutencao TEXT,
        infraestrutura TEXT,
        desenvolvimento TEXT,
        UNIQUE(ano, data_inicio, manutencao) ON CONFLICT REPLACE
    )
    """)
    
    conn.commit()
    conn.close()

def save_plantoes_matutino(records: list[dict]):
    """Salva ou atualiza registros de plantão matutino no SQLite."""
    setup_plantoes_tables()
    if not records:
        return
    conn = get_connection()
    cursor = conn.cursor()
    for r in records:
        cursor.execute("""
        INSERT OR REPLACE INTO plantoes_matutino (ano, data_iso, dia_semana, servidor, telefone)
        VALUES (?, ?, ?, ?, ?)
        """, (
            r.get("ano"),
            r.get("data_iso"),
            r.get("dia_semana"),
            r.get("servidor"),
            r.get("telefone")
        ))
    conn.commit()
    conn.close()

def save_plantoes_semanal(records: list[dict]):
    """Salva ou atualiza registros de plantão semanal SIMP no SQLite."""
    setup_plantoes_tables()
    if not records:
        return
    conn = get_connection()
    cursor = conn.cursor()
    for r in records:
        cursor.execute("""
        INSERT OR REPLACE INTO plantoes_semanal 
        (ano, mes, periodo_str, data_inicio, data_fim, service_desk, manutencao, infraestrutura, desenvolvimento)
        VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?)
        """, (
            r.get("ano"),
            r.get("mes"),
            r.get("periodo_str"),
            r.get("data_inicio"),
            r.get("data_fim"),
            r.get("service_desk"),
            r.get("manutencao"),
            r.get("infraestrutura"),
            r.get("desenvolvimento")
        ))
    conn.commit()
    conn.close()

def get_plantoes_matutino(ano: int | None = None) -> pd.DataFrame:
    """Retorna DataFrame pandas dos plantões matutinos cadastrados."""
    setup_plantoes_tables()
    conn = get_connection()
    if ano:
        df = pd.read_sql_query("SELECT * FROM plantoes_matutino WHERE ano = ? ORDER BY data_iso ASC", conn, params=(ano,))
    else:
        df = pd.read_sql_query("SELECT * FROM plantoes_matutino ORDER BY data_iso ASC", conn)
    conn.close()
    return df

def get_plantoes_semanal(ano: int | None = None) -> pd.DataFrame:
    """Retorna DataFrame pandas dos plantões semanais SIMP cadastrados."""
    setup_plantoes_tables()
    conn = get_connection()
    if ano:
        df = pd.read_sql_query("SELECT * FROM plantoes_semanal WHERE ano = ? ORDER BY data_inicio ASC", conn, params=(ano,))
    else:
        df = pd.read_sql_query("SELECT * FROM plantoes_semanal ORDER BY data_inicio ASC", conn)
    conn.close()
    return df
