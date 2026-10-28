import pandas as pd
from .connection import get_connection, DB_TYPE

def get_central_telefonica_df() -> pd.DataFrame:
    """
    Retorna o DataFrame contendo todos os ramais da Central Telefônica (OXE) cadastrados no banco.
    Se a tabela estiver vazia, tenta executar o pré-processamento para popular o banco.
    """
    order_clause = (
        "ORDER BY CASE WHEN ramal ~ '^[0-9]+$' THEN CAST(ramal AS INTEGER) ELSE 999999 END ASC, ramal ASC"
        if DB_TYPE in ["postgres", "postgresql"]
        else "ORDER BY CAST(ramal AS INTEGER) ASC"
    )
    conn = get_connection()
    try:
        df = pd.read_sql_query(f"SELECT * FROM central_telefonica {order_clause}", conn)
        conn.close()
        if not df.empty:
            return df
    except Exception:
        conn.close()

    try:
        from src.preprocess_oxe import preprocess_oxe
        if preprocess_oxe():
            conn = get_connection()
            df = pd.read_sql_query(f"SELECT * FROM central_telefonica {order_clause}", conn)
            conn.close()
            return df
    except Exception:
        pass

    from config import OUTPUT_DIR_TRATADOS
    files = sorted(OUTPUT_DIR_TRATADOS.glob("Central_Telefonica_OXE_Tratados_*.xlsx"), key=lambda f: f.stat().st_mtime)
    if files:
        df = pd.read_excel(files[-1], dtype=str)
        df.fillna("", inplace=True)
        return df

    return pd.DataFrame()
