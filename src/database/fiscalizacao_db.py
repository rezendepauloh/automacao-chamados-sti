import shutil
import tempfile
from pathlib import Path
import pandas as pd
from .connection import get_connection

def sync_fiscalizacao_from_excel(file_path: str) -> bool:
    """Lê os dados da planilha de fiscalização de contratos e salva no SQLite em tabelas dedicadas."""
    from src.config import setup_logging, DEBUG_DIR_FISCALIZACAO

    logger = setup_logging(DEBUG_DIR_FISCALIZACAO / "sync.log", "fiscalizacao_sync")
    logger.info(f"Iniciando sincronização da planilha de fiscalização: {file_path}")

    if not file_path or not Path(file_path).exists():
        logger.error(f"Arquivo da planilha de fiscalização não encontrado: {file_path}")
        return False

    temp_file = tempfile.NamedTemporaryFile(delete=False, suffix=".xlsx")
    temp_path = temp_file.name
    temp_file.close()

    try:
        shutil.copy2(file_path, temp_path)
        excel_data = pd.ExcelFile(temp_path)

        df_indicacoes = pd.read_excel(excel_data, sheet_name="Indicações") if "Indicações" in excel_data.sheet_names else pd.DataFrame()
        df_publicacoes = pd.read_excel(excel_data, sheet_name="Publicações") if "Publicações" in excel_data.sheet_names else pd.DataFrame()
        df_contador = pd.read_excel(excel_data, sheet_name="Contador") if "Contador" in excel_data.sheet_names else pd.DataFrame()

        conn = get_connection()
        try:
            df_indicacoes.to_sql("fiscalizacao_indicacoes", conn, if_exists="replace", index=False)
            df_publicacoes.to_sql("fiscalizacao_publicacoes", conn, if_exists="replace", index=False)
            df_contador.to_sql("fiscalizacao_contador", conn, if_exists="replace", index=False)
            conn.commit()
        finally:
            conn.close()

        logger.info(f"Sincronização de fiscalização concluída com sucesso! Indicações: {len(df_indicacoes)}, Publicações: {len(df_publicacoes)}, Contador: {len(df_contador)}")
        return True
    except Exception as e:
        logger.error(f"Erro ao ler/salvar planilha de fiscalização: {e}")
        raise e
    finally:
        Path(temp_path).unlink(missing_ok=True)

def get_fiscalizacao_indicacoes_df() -> pd.DataFrame:
    """Retorna um DataFrame com os dados da tabela fiscalizacao_indicacoes do SQLite."""
    try:
        conn = get_connection()
        df = pd.read_sql_query("SELECT * FROM fiscalizacao_indicacoes", conn)
        conn.close()
        return df
    except Exception:
        return pd.DataFrame()

def get_fiscalizacao_publicacoes_df() -> pd.DataFrame:
    """Retorna um DataFrame com os dados da tabela fiscalizacao_publicacoes do SQLite."""
    try:
        conn = get_connection()
        df = pd.read_sql_query("SELECT * FROM fiscalizacao_publicacoes", conn)
        conn.close()
        return df
    except Exception:
        return pd.DataFrame()

def get_fiscalizacao_contador_df() -> pd.DataFrame:
    """Retorna um DataFrame com os dados da tabela fiscalizacao_contador do SQLite."""
    try:
        conn = get_connection()
        df = pd.read_sql_query("SELECT * FROM fiscalizacao_contador", conn)
        conn.close()
        return df
    except Exception:
        return pd.DataFrame()
