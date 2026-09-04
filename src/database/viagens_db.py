import shutil
import tempfile
from pathlib import Path
from datetime import datetime
import pandas as pd
from .connection import get_connection

def setup_viagens_table():
    """Cria a tabela de viagens da bancada se não existir."""
    conn = get_connection()
    cursor = conn.cursor()
    cursor.execute("""
    CREATE TABLE IF NOT EXISTS viagens (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        quem_foi TEXT,
        chamado TEXT,
        saida_iso TEXT,
        retorno_iso TEXT,
        saida_br TEXT,
        retorno_br TEXT,
        localidade TEXT,
        created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP
    )
    """)
    conn.commit()
    conn.close()

def _parse_dates(val):
    """
    Interpreta datas da planilha (Timestamp, string YYYY-MM-DD ou DD/MM/YYYY)
    Retorna tupla (iso_str 'YYYY-MM-DD', br_str 'DD/MM/YYYY').
    """
    if val is None or pd.isna(val):
        return "", ""
    
    dt = None
    if isinstance(val, (datetime, pd.Timestamp)):
        dt = val
    else:
        val_str = str(val).strip()
        if not val_str or val_str.lower() in ["nat", "nan", "null", "none"]:
            return "", ""
        for fmt in ("%d/%m/%Y", "%Y-%m-%d %H:%M:%S", "%Y-%m-%d", "%d/%m/%Y %H:%M:%S"):
            try:
                dt = datetime.strptime(val_str.split(".")[0], fmt)
                break
            except Exception:
                continue
        if dt is None:
            try:
                dt = pd.to_datetime(val_str, dayfirst=True)
            except Exception:
                return "", ""

    if dt:
        return dt.strftime("%Y-%m-%d"), dt.strftime("%d/%m/%Y")
    return "", ""

def sync_viagens_from_excel(file_path_or_buffer):
    """
    Lê os dados da planilha de viagens (Planilha de Viagens.xlsx) e salva no banco de dados.
    Suporta caminho local (.xlsx) ou buffer (BytesIO / Streamlit).
    """
    from src.config import setup_logging, DEBUG_DIR_VIAGENS
    logger = setup_logging(DEBUG_DIR_VIAGENS / "viagens_sync.log", "viagens_sync")
    logger.info("Iniciando sincronização da planilha de viagens da bancada...")

    setup_viagens_table()

    try:
        temp_file = tempfile.NamedTemporaryFile(delete=False, suffix=".xlsx")
        temp_path = temp_file.name
        temp_file.close()

        try:
            if isinstance(file_path_or_buffer, (str, Path)):
                p = Path(file_path_or_buffer)
                if not p.exists():
                    logger.error(f"Arquivo de viagens não encontrado: {file_path_or_buffer}")
                    return False
                shutil.copy2(str(p), temp_path)
            elif hasattr(file_path_or_buffer, "read"):
                file_path_or_buffer.seek(0)
                with open(temp_path, "wb") as f_out:
                    f_out.write(file_path_or_buffer.read())
            else:
                logger.error("Tipo de arquivo/buffer inválido fornecido para viagens.")
                return False

            with pd.ExcelFile(temp_path) as xls:
                target_sheet = "Plan1" if "Plan1" in xls.sheet_names else xls.sheet_names[0]
                df = pd.read_excel(xls, sheet_name=target_sheet)
        finally:
            Path(temp_path).unlink(missing_ok=True)

        logger.info(f"Planilha de viagens lida com sucesso. Total de linhas: {len(df)}")
    except Exception as e:
        logger.error(f"Erro ao processar arquivo Excel de viagens: {e}")
        raise e

    df = df.fillna("")

    conn = get_connection()
    cursor = conn.cursor()

    cursor.execute("DELETE FROM viagens")

    added_count = 0
    for _, row in df.iterrows():
        quem_foi = str(row.get("Quem foi", "")).strip()
        chamado = str(row.get("Chamado", "")).strip().replace(".0", "")
        localidade = str(row.get("Localidade", "")).strip()

        saida_val = row.get("Saída", "")
        retorno_val = row.get("Retorno", "")

        saida_iso, saida_br = _parse_dates(saida_val)
        retorno_iso, retorno_br = _parse_dates(retorno_val)

        # Se ambos os campos de data forem vazios e não tiver localidade, ignora linha em branco
        if not saida_iso and not retorno_iso and not localidade:
            continue

        cursor.execute("""
        INSERT INTO viagens (
            quem_foi, chamado, saida_iso, retorno_iso, saida_br, retorno_br, localidade
        ) VALUES (?, ?, ?, ?, ?, ?, ?)
        """, (
            quem_foi, chamado, saida_iso, retorno_iso, saida_br, retorno_br, localidade
        ))
        added_count += 1

    conn.commit()
    conn.close()
    logger.info(f"Sincronização concluída: {added_count} registros de viagens salvos.")
    return True

def get_viagens_df() -> pd.DataFrame:
    """Retorna todas as viagens cadastradas ordenadas pela data de saída decrescente."""
    setup_viagens_table()
    conn = get_connection()
    df = pd.read_sql_query("SELECT * FROM viagens ORDER BY saida_iso DESC, id DESC", conn)
    conn.close()
    return df
