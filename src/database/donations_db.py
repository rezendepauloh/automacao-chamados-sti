import shutil
import tempfile
from pathlib import Path
from datetime import datetime
import pandas as pd
from .connection import get_connection

def setup_donations_table():
    """Cria a tabela de equipamentos doados se não existir e garante colunas incrementais."""
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
        motivo_baixa TEXT,
        memoria_ram TEXT
    )
    """)
    conn.commit()

    # Migração incremental segura caso a tabela já exista sem memoria_ram
    try:
        cursor.execute("ALTER TABLE equipamentos_doados ADD COLUMN memoria_ram TEXT")
        conn.commit()
    except Exception:
        try:
            conn.rollback()
        except Exception:
            pass

    conn.close()
    setup_entregas_dmp_table()


def setup_entregas_dmp_table():
    """Cria a tabela de entregas do DMP para a Bancada STI se não existir."""
    conn = get_connection()
    cursor = conn.cursor()
    cursor.execute("""
    CREATE TABLE IF NOT EXISTS entregas_dmp (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        patrimonio TEXT NOT NULL,
        modelo TEXT,
        serial_number TEXT,
        equipamento TEXT,
        tipo_movimentacao TEXT,
        data_entrega TEXT,
        chamado TEXT,
        ssd TEXT,
        memoria_ram TEXT,
        origem_unidade TEXT,
        usuario_anterior TEXT,
        hostname TEXT,
        responsavel_entrega TEXT,
        recebido_por TEXT,
        observacoes TEXT,
        created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP
    )
    """)
    conn.commit()
    conn.close()

def sync_donations_from_excel(file_path_or_buffer):
    """Lê os dados da planilha de equipamentos doados e salva no SQLite.
    Suporta caminho de arquivo local (.xlsx) ou buffer de upload (BytesIO / Streamlit).
    """
    from src.config import setup_logging, DEBUG_DIR_DONATIONS
    
    logger = setup_logging(DEBUG_DIR_DONATIONS / "donations.log", "donations_sync")
    logger.info(f"Iniciando sincronização da planilha de doações/redistribuição...")
    
    setup_donations_table()
    
    try:
        temp_file = tempfile.NamedTemporaryFile(delete=False, suffix=".xlsx")
        temp_path = temp_file.name
        temp_file.close()
        
        try:
            if isinstance(file_path_or_buffer, (str, Path)):
                p = Path(file_path_or_buffer)
                if not p.exists():
                    logger.error(f"Arquivo não encontrado: {file_path_or_buffer}")
                    return False
                shutil.copy2(str(p), temp_path)
            elif hasattr(file_path_or_buffer, "read"):
                file_path_or_buffer.seek(0)
                with open(temp_path, "wb") as f_out:
                    f_out.write(file_path_or_buffer.read())
            else:
                logger.error("Tipo de arquivo/buffer inválido fornecido para doações.")
                return False

            with pd.ExcelFile(temp_path) as xls:
                target_sheet = "Equipamentos doados" if "Equipamentos doados" in xls.sheet_names else xls.sheet_names[0]
                df = pd.read_excel(xls, sheet_name=target_sheet)
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
        
        # Memória RAM: busca flexível por variações de nome de coluna
        mem_ram_val = ""
        for ram_col in ['Memória RAM', 'Memória Ram', 'Memoria RAM', 'Memoria Ram', 'Memória', 'Memoria', 'RAM', 'Ram']:
            if ram_col in row and str(row.get(ram_col, '')).strip():
                mem_ram_val = str(row.get(ram_col, '')).strip()
                break
        if mem_ram_val.lower() in ["nan", "none", "null"]:
            mem_ram_val = ""
        mem_ram_val = mem_ram_val.replace(".0", "")

        # SSD / Armazenamento: busca flexível
        ssd_val = ""
        for ssd_col in ['SSD', 'Ssd', 'HD', 'Hd', 'HD/SSD', 'Armazenamento', 'Disco']:
            if ssd_col in row and str(row.get(ssd_col, '')).strip():
                ssd_val = str(row.get(ssd_col, '')).strip()
                break
        if ssd_val.lower() in ["nan", "none", "null"]:
            ssd_val = ""
        ssd_val = ssd_val.replace(".0", "")

        cursor.execute("""
        INSERT INTO equipamentos_doados (
            patrimonio, modelo, serial_number, equipamento, 
            tipo_movimentacao, data_movimentacao, chamado, ssd, motivo_baixa, memoria_ram
        ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
        """, (
            patrimonio,
            str(row.get('Modelo', '')).strip(),
            str(row.get('Serial Number PC', '')).strip(),
            str(row.get('Equipamento', '')).strip(),
            str(row.get('Doação ou Baixa', '')).strip(),
            dt_str,
            chamado,
            ssd_val,
            str(row.get('Motivo baixa', '')).strip(),
            mem_ram_val
        ))
        added_count += 1
        
    conn.commit()
    conn.close()
    logger.info(f"Sincronização concluída. {added_count} registros importados para a tabela equipamentos_doados.")

def get_donations_data() -> pd.DataFrame:
    """Retorna todos os equipamentos doados/movimentados da tabela relacional."""
    setup_donations_table()
    conn = get_connection()
    df = pd.read_sql_query("SELECT * FROM equipamentos_doados", conn)
    conn.close()
    if 'memoria_ram' not in df.columns:
        df['memoria_ram'] = ""
    else:
        df['memoria_ram'] = df['memoria_ram'].fillna("").astype(str).replace({"nan": "", "None": "", "null": ""})
    return df


def save_entregas_dmp(records: list) -> int:
    """
    Persiste uma lista de registros de equipamentos entregues pelo DMP à Bancada STI.
    Ignora registros sem patrimônio e retorna a quantidade gravada.
    """
    setup_entregas_dmp_table()
    if not records:
        return 0

    def _clean(val) -> str:
        if val is None:
            return ""
        s = str(val).strip()
        if s.lower() in ("nan", "none", "null", "<na>"):
            return ""
        return s

    conn = get_connection()
    cursor = conn.cursor()
    inserted = 0
    try:
        for rec in records:
            pat = _clean(rec.get("patrimonio", "")).replace(".0", "")
            if not pat:
                continue

            data_ent = _clean(rec.get("data_entrega", "")) or datetime.now().strftime("%Y-%m-%d")
            if len(data_ent) > 10:
                data_ent = data_ent[:10]
            if "/" in data_ent:
                try:
                    data_ent = datetime.strptime(data_ent, "%d/%m/%Y").strftime("%Y-%m-%d")
                except ValueError:
                    pass

            cursor.execute(
                """
                INSERT INTO entregas_dmp (
                    patrimonio, modelo, serial_number, equipamento, tipo_movimentacao,
                    data_entrega, chamado, ssd, memoria_ram, origem_unidade,
                    usuario_anterior, hostname, responsavel_entrega, recebido_por, observacoes
                ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
                """,
                (
                    pat,
                    _clean(rec.get("modelo", "")),
                    _clean(rec.get("serial_number", "")),
                    _clean(rec.get("equipamento", "")) or "Computador",
                    _clean(rec.get("tipo_movimentacao", "")) or "Recebimento DMP",
                    data_ent,
                    _clean(rec.get("chamado", "")).replace(".0", ""),
                    _clean(rec.get("ssd", "")),
                    _clean(rec.get("memoria_ram", "")),
                    _clean(rec.get("origem_unidade", "")),
                    _clean(rec.get("usuario_anterior", "")),
                    _clean(rec.get("hostname", "")),
                    _clean(rec.get("responsavel_entrega", "")),
                    _clean(rec.get("recebido_por", "")),
                    _clean(rec.get("observacoes", "")),
                ),
            )
            inserted += 1
        conn.commit()
    finally:
        cursor.close()
        conn.close()

    return inserted


def get_entregas_dmp_data() -> pd.DataFrame:
    """Retorna todos os equipamentos entregues pelo DMP ordenados por data mais recente."""
    setup_entregas_dmp_table()
    conn = get_connection()
    try:
        df = pd.read_sql_query(
            "SELECT * FROM entregas_dmp ORDER BY COALESCE(data_entrega, '') DESC, id DESC",
            conn,
        )
    finally:
        conn.close()

    expected_cols = [
        "id", "patrimonio", "modelo", "serial_number", "equipamento",
        "tipo_movimentacao", "data_entrega", "chamado", "ssd", "memoria_ram",
        "origem_unidade", "usuario_anterior", "hostname",
        "responsavel_entrega", "recebido_por", "observacoes", "created_at",
    ]
    for col in expected_cols:
        if col not in df.columns:
            df[col] = ""
        elif col != "id":
            df[col] = df[col].fillna("").astype(str).replace({"nan": "", "None": "", "null": ""})
    return df


def delete_entrega_dmp(record_id: int) -> bool:
    """Remove um registro específico de entrega do DMP pelo ID."""
    setup_entregas_dmp_table()
    conn = get_connection()
    cursor = conn.cursor()
    try:
        cursor.execute("DELETE FROM entregas_dmp WHERE id = ?", (int(record_id),))
        deleted = cursor.rowcount > 0
        conn.commit()
        return deleted
    finally:
        cursor.close()
        conn.close()


def update_entrega_dmp(record_id: int, updates: dict) -> bool:
    """Atualiza os campos editáveis de um registro de entrega do DMP pelo ID."""
    setup_entregas_dmp_table()
    if not updates:
        return False

    allowed_cols = (
        "patrimonio", "modelo", "serial_number", "equipamento", "tipo_movimentacao",
        "data_entrega", "chamado", "ssd", "memoria_ram", "origem_unidade",
        "responsavel_entrega", "observacoes",
    )
    set_clauses = []
    params = []
    for col in allowed_cols:
        if col in updates:
            val = "" if updates[col] is None else str(updates[col]).strip()
            if val.lower() in ("nan", "none", "null", "<na>"):
                val = ""
            if col == "data_entrega" and "/" in val:
                try:
                    val = datetime.strptime(val[:10], "%d/%m/%Y").strftime("%Y-%m-%d")
                except ValueError:
                    pass
            set_clauses.append(f"{col} = ?")
            params.append(val)

    if not set_clauses:
        return False

    params.append(int(record_id))
    conn = get_connection()
    cursor = conn.cursor()
    try:
        cursor.execute(
            f"UPDATE entregas_dmp SET {', '.join(set_clauses)} WHERE id = ?",
            tuple(params),
        )
        updated = cursor.rowcount > 0
        conn.commit()
        return updated
    finally:
        cursor.close()
        conn.close()
