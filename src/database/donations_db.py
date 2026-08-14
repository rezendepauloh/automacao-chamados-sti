import shutil
import tempfile
from pathlib import Path
from datetime import datetime
import pandas as pd
from .connection import get_connection

def setup_donations_table():
    """Cria a tabela de equipamentos doados se não existir."""
    conn = get_connection()
    cursor = conn.cursor()
    cursor.execute("""
    CREATE TABLE IF NOT EXISTS equipamentos_doados (
        patrimonio TEXT,
        modelo TEXT,
        serial_number TEXT,
        equipamento TEXT,
        tipo_movimentacao TEXT,
        data_movimentacao TEXT,
        chamado TEXT,
        ssd TEXT,
        motivo_baixa TEXT
    )
    """)
    conn.commit()
    conn.close()

def sync_donations_from_excel(file_path: str):
    """Lê os dados da planilha de equipamentos doados e salva no SQLite."""
    from src.config import setup_logging, DEBUG_DIR_DONATIONS
    
    logger = setup_logging(DEBUG_DIR_DONATIONS / "donations.log", "donations_sync")
    logger.info(f"Iniciando sincronização da planilha: {file_path}")
    
    setup_donations_table()
    
    try:
        temp_file = tempfile.NamedTemporaryFile(delete=False, suffix=".xlsx")
        temp_path = temp_file.name
        temp_file.close()
        
        try:
            shutil.copy2(file_path, temp_path)
            df = pd.read_excel(temp_path, sheet_name="Equipamentos doados")
        finally:
            Path(temp_path).unlink(missing_ok=True)
            
        logger.info(f"Planilha lida com sucesso. Total de linhas encontradas: {len(df)}")
    except Exception as e:
        logger.error(f"Erro ao ler a planilha Excel: {e}")
        raise e
    
    df = df.fillna("")
    
    conn = get_connection()
    cursor = conn.cursor()
    
    cursor.execute("DELETE FROM equipamentos_doados")
    
    added_count = 0
    for _, row in df.iterrows():
        dt_val = row.get('Data da doação', '')
        if isinstance(dt_val, (datetime, pd.Timestamp)) and not pd.isna(dt_val):
            dt_str = dt_val.strftime("%Y-%m-%d")
        else:
            dt_str = str(dt_val).strip()
            if len(dt_str) > 10:
                dt_str = dt_str[:10]
            if dt_str.lower() in ["nat", "nan", "null", ""]:
                dt_str = ""
        
        patrimonio = str(row.get('Patrimônio', '')).strip().replace(".0", "")
        chamado = str(row.get('Tem chamado?', '')).strip().replace(".0", "")
        
        cursor.execute("""
        INSERT INTO equipamentos_doados (
            patrimonio, modelo, serial_number, equipamento, 
            tipo_movimentacao, data_movimentacao, chamado, ssd, motivo_baixa
        ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?)
        """, (
            patrimonio,
            str(row.get('Modelo', '')).strip(),
            str(row.get('Serial Number PC', '')).strip(),
            str(row.get('Equipamento', '')).strip(),
            str(row.get('Doação ou Baixa', '')).strip(),
            dt_str,
            chamado,
            str(row.get('SSD', '')).strip(),
            str(row.get('Motivo baixa', '')).strip()
        ))
        added_count += 1
        
    conn.commit()
    conn.close()
    logger.info(f"Sincronização concluída. {added_count} registros importados para a tabela equipamentos_doados.")

def get_donations_data() -> pd.DataFrame:
    """Retorna todos os equipamentos doados/movimentados da tabela SQLite."""
    setup_donations_table()
    conn = get_connection()
    df = pd.read_sql_query("SELECT * FROM equipamentos_doados", conn)
    conn.close()
    return df
