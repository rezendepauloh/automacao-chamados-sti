import shutil
import tempfile
from pathlib import Path
from datetime import datetime
import pandas as pd
from .connection import get_connection

def setup_garantia_tables():
    """Cria as tabelas de contratos e chamados de garantia no SQLite se não existirem."""
    conn = get_connection()
    cursor = conn.cursor()
    cursor.execute("""
    CREATE TABLE IF NOT EXISTS garantia_contratos (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        contrato TEXT,
        pu_saj TEXT,
        item TEXT,
        contratacao_por TEXT,
        fornecedor TEXT,
        termo_referencia TEXT,
        termo_recebimento TEXT,
        nota_fiscal TEXT,
        garantia_inicio TEXT,
        garantia_fim TEXT,
        status_garantia TEXT,
        dias_restantes INTEGER,
        link_suporte TEXT,
        data_atualizacao TEXT
    )
    """)
    cursor.execute("""
    CREATE TABLE IF NOT EXISTS garantia_chamados (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        item TEXT,
        status TEXT,
        numero_serie TEXT,
        patrimonio TEXT,
        chamado_mpm TEXT,
        chamado_externo TEXT,
        defeito TEXT,
        solucao TEXT,
        nota_no_chamado TEXT,
        chamado_dmp TEXT,
        data_atualizacao TEXT
    )
    """)
    conn.commit()
    conn.close()

def sync_garantia_from_excel(excel_path: str = None) -> bool:
    """
    Sincroniza os dados da planilha de Garantia para as tabelas SQLite.
    """
    setup_garantia_tables()
    if not excel_path:
        from src.config import WARRANTY_FILE_PATH
        excel_path = str(WARRANTY_FILE_PATH)

    p = Path(excel_path)
    if not p.exists():
        return False

    temp_file = tempfile.NamedTemporaryFile(delete=False, suffix=".xlsx")
    temp_path = temp_file.name
    temp_file.close()

    try:
        shutil.copy2(excel_path, temp_path)
        with pd.ExcelFile(temp_path) as xls:
            conn = get_connection()
            cursor = conn.cursor()
            now_str = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
            today_date = datetime.now().date()

            if 'Contratos' in xls.sheet_names:
                df_raw_c = pd.read_excel(xls, sheet_name='Contratos', header=None, dtype=str)
                header_idx_c = 0
                for r_idx, r in df_raw_c.iterrows():
                    r_str = " ".join([str(val).lower() for val in r if pd.notnull(val)])
                    if 'item' in r_str or 'contrato' in r_str or 'fornecedor' in r_str or 'saj' in r_str:
                        header_idx_c = r_idx
                        break

                df_contratos = pd.read_excel(xls, sheet_name='Contratos', header=header_idx_c, dtype=str)
                df_contratos.fillna("", inplace=True)
                cursor.execute("DELETE FROM garantia_contratos")

                for _, row in df_contratos.iterrows():
                    contrato = ""
                    for k in row.index:
                        if 'contrato' in str(k).lower():
                            contrato = str(row[k]).strip()
                            break

                    pu_saj = ""
                    for k in row.index:
                        if 'saj' in str(k).lower() or 'pu' in str(k).lower():
                            pu_saj = str(row[k]).strip()
                            break

                    item = ""
                    for k in row.index:
                        if 'item' in str(k).lower() or 'equipamento' in str(k).lower():
                            item = str(row[k]).strip()
                            break

                    contratacao_por = ""
                    for k in row.index:
                        if 'contrata' in str(k).lower() or 'preg' in str(k).lower() or 'ata' in str(k).lower():
                            contratacao_por = str(row[k]).strip()
                            break

                    fornecedor = ""
                    for k in row.index:
                        if 'fornecedor' in str(k).lower() or 'empresa' in str(k).lower():
                            fornecedor = str(row[k]).strip()
                            break

                    termo_ref = ""
                    for k in row.index:
                        if 'refer' in str(k).lower():
                            termo_ref = str(row[k]).strip()
                            break

                    termo_rec = ""
                    for k in row.index:
                        if 'receb' in str(k).lower() or 'definitiv' in str(k).lower():
                            termo_rec = str(row[k]).strip()
                            break

                    nota_fiscal = ""
                    for k in row.index:
                        if 'nota' in str(k).lower() or 'fiscal' in str(k).lower():
                            nota_fiscal = str(row[k]).strip()
                            break

                    g_inicio = ""
                    for k in row.index:
                        if 'começ' in str(k).lower() or 'iníc' in str(k).lower() or 'inicio' in str(k).lower():
                            g_inicio = str(row[k]).strip()
                            break

                    g_fim = ""
                    for k in row.index:
                        if 'fim' in str(k).lower() or 'térm' in str(k).lower() or 'termino' in str(k).lower():
                            g_fim = str(row[k]).strip()
                            break

                    link_sup = ""
                    for k in row.index:
                        if 'link' in str(k).lower() or 'chamado' in str(k).lower() or 'site' in str(k).lower() or 'abertura' in str(k).lower():
                            link_sup = str(row[k]).strip()
                            break

                    if not item and not contrato and not fornecedor and not pu_saj:
                        continue

                    dias_restantes = None
                    status_garantia = "Não Informada"
                    if g_fim:
                        try:
                            dt_fim_obj = pd.to_datetime(g_fim, dayfirst=False, errors='coerce')
                            if pd.notnull(dt_fim_obj):
                                dias_restantes = (dt_fim_obj.date() - today_date).days
                                if dias_restantes < 0:
                                    status_garantia = "Garantia Vencida"
                                elif dias_restantes <= 30:
                                    status_garantia = "A Vencer (≤ 30 dias)"
                                else:
                                    status_garantia = "Garantia Ativa"
                                g_fim = dt_fim_obj.strftime("%Y-%m-%d")
                        except Exception:
                            pass

                    cursor.execute("""
                    INSERT INTO garantia_contratos (
                        contrato, pu_saj, item, contratacao_por, fornecedor,
                        termo_referencia, termo_recebimento, nota_fiscal,
                        garantia_inicio, garantia_fim, status_garantia,
                        dias_restantes, link_suporte, data_atualizacao
                    ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
                    """, (
                        contrato, pu_saj, item, contratacao_por, fornecedor,
                        termo_ref, termo_rec, nota_fiscal,
                        g_inicio, g_fim, status_garantia,
                        dias_restantes, link_sup, now_str
                    ))

            if 'Chamados' in xls.sheet_names:
                df_raw_ch = pd.read_excel(xls, sheet_name='Chamados', header=None, dtype=str)
                header_idx_ch = 1
                for r_idx, r in df_raw_ch.iterrows():
                    r_str = " ".join([str(val).lower() for val in r if pd.notnull(val)])
                    if 'item' in r_str and ('status' in r_str or 'série' in r_str or 'serie' in r_str or 'patrimô' in r_str or 'defeito' in r_str):
                        header_idx_ch = r_idx
                        break

                df_chamados = pd.read_excel(xls, sheet_name='Chamados', header=header_idx_ch, dtype=str)
                df_chamados.fillna("", inplace=True)
                cursor.execute("DELETE FROM garantia_chamados")

                for _, row in df_chamados.iterrows():
                    item = ""
                    for k in row.index:
                        if 'item' in str(k).lower() or 'equipamento' in str(k).lower():
                            item = str(row[k]).strip()
                            break

                    status = ""
                    for k in row.index:
                        if 'status' in str(k).lower() or 'situac' in str(k).lower():
                            status = str(row[k]).strip()
                            break

                    n_serie = ""
                    for k in row.index:
                        if 'série' in str(k).lower() or 'serie' in str(k).lower() or 'serial' in str(k).lower():
                            n_serie = str(row[k]).strip()
                            break

                    patrimonio = ""
                    for k in row.index:
                        if 'patrimô' in str(k).lower() or 'patrimo' in str(k).lower() or 'tombo' in str(k).lower():
                            patrimonio = str(row[k]).strip()
                            break

                    c_mpm = ""
                    for k in row.index:
                        if 'mpm' in str(k).lower() or 'otrs' in str(k).lower() or 'citsmart' in str(k).lower():
                            c_mpm = str(row[k]).strip()
                            break

                    c_ext = ""
                    for k in row.index:
                        if 'externo' in str(k).lower() or 'fornecedor' in str(k).lower():
                            c_ext = str(row[k]).strip()
                            break

                    defeito = ""
                    for k in row.index:
                        if 'defeito' in str(k).lower() or 'problema' in str(k).lower():
                            defeito = str(row[k]).strip()
                            break

                    solucao = ""
                    for k in row.index:
                        if 'soluç' in str(k).lower() or 'soluc' in str(k).lower() or 'acao' in str(k).lower():
                            solucao = str(row[k]).strip()
                            break

                    nota_chamado = ""
                    for k in row.index:
                        if 'nota' in str(k).lower():
                            nota_chamado = str(row[k]).strip()
                            break

                    c_dmp = ""
                    for k in row.index:
                        if 'dmp' in str(k).lower():
                            c_dmp = str(row[k]).strip()
                            break

                    if not item and not patrimonio and not n_serie and not c_mpm and not c_ext:
                        continue

                    cursor.execute("""
                    INSERT INTO garantia_chamados (
                        item, status, numero_serie, patrimonio, chamado_mpm,
                        chamado_externo, defeito, solucao, nota_no_chamado,
                        chamado_dmp, data_atualizacao
                    ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
                    """, (
                        item, status, n_serie, patrimonio, c_mpm,
                        c_ext, defeito, solucao, nota_chamado,
                        c_dmp, now_str
                    ))

            conn.commit()
            conn.close()
            return True
    except Exception as e:
        print(f"Erro ao sincronizar garantia da planilha: {e}")
        return False
    finally:
        try:
            Path(temp_path).unlink(missing_ok=True)
        except Exception as unlink_err:
            print(f"Aviso ao excluir arquivo temporário de garantia: {unlink_err}")

def get_garantia_contratos_df() -> pd.DataFrame:
    """Retorna o DataFrame de Contratos de Garantia do SQLite."""
    setup_garantia_tables()
    conn = get_connection()
    try:
        df = pd.read_sql_query("SELECT * FROM garantia_contratos ORDER BY id ASC", conn)
        conn.close()
        return df
    except Exception:
        conn.close()
        return pd.DataFrame()

def get_garantia_chamados_df() -> pd.DataFrame:
    """Retorna o DataFrame de Chamados de Garantia do SQLite."""
    setup_garantia_tables()
    conn = get_connection()
    try:
        df = pd.read_sql_query("SELECT * FROM garantia_chamados ORDER BY id ASC", conn)
        conn.close()
        return df
    except Exception:
        conn.close()
        return pd.DataFrame()
