import os
import re
import shutil
import tempfile
from pathlib import Path
from datetime import datetime, date
import pandas as pd
import openpyxl
from .connection import get_connection

def setup_ferias_table():
    """Cria a tabela de férias da bancada se não existir."""
    conn = get_connection()
    cursor = conn.cursor()
    cursor.execute("""
    CREATE TABLE IF NOT EXISTS ferias_bancada (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        ano INTEGER,
        membro TEXT,
        tipo_escala TEXT, -- 'ferias', 'licencas_compensacoes', 'recesso_forense'
        mes INTEGER, -- 1 a 12 ou NULL para recesso
        mes_nome TEXT,
        periodo_bruto TEXT,
        data_inicio_iso TEXT,
        data_fim_iso TEXT,
        data_inicio_br TEXT,
        data_fim_br TEXT,
        dias INTEGER,
        status TEXT, -- 'Confirmada', 'Prevista', 'Usufruto', 'Venda', 'Cancelada', 'Recesso'
        cor_hex TEXT,
        created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP
    )
    """)
    conn.commit()
    conn.close()

def _parse_month_from_header(val, default_year: int) -> tuple[int, str]:
    """Extrai número do mês (1-12) e nome do cabeçalho da coluna."""
    if isinstance(val, (datetime, pd.Timestamp, date)):
        mes_num = val.month
    else:
        s = str(val).strip().lower()
        meses = {
            "jan": 1, "fev": 2, "mar": 3, "abr": 4, "mai": 5, "jun": 6,
            "jul": 7, "ago": 8, "set": 9, "out": 10, "nov": 11, "dez": 12
        }
        mes_num = 1
        for k, v in meses.items():
            if k in s:
                mes_num = v
                break
    
    nomes_meses = [
        "", "Janeiro", "Fevereiro", "Março", "Abril", "Maio", "Junho",
        "Julho", "Agosto", "Setembro", "Outubro", "Novembro", "Dezembro"
    ]
    return mes_num, nomes_meses[mes_num] if 1 <= mes_num <= 12 else "Geral"

def parse_periodo_cell(val_str: str, ano: int, mes_padrao: int) -> list[dict]:
    """
    Interpreta o texto da célula de período de férias/licença.
    Suporta:
      - '20 a 29', '6 a 15', '07 a 16', '08 a 17'
      - '16/09 a 04/10', '30/09 a 09/10'
      - '16/11 à 29/11'
      - '15-26/03'
      - '3-7/ago'
      - '30 e 31', '16 e 17', '18 e 19', '18 a 19'
      - '6, 7 e 9'
      - '17,18,21 e 22'
    Retorna uma lista de tuplas/dicionários com datas início e fim para cada intervalo contínuo.
    """
    if not val_str or not isinstance(val_str, str):
        return []
        
    s = val_str.strip()
    if not s or s.lower() in ["nan", "none", "-", "x"]:
        return []

    results = []

    meses_abreviados = {
        "jan": 1, "fev": 2, "mar": 3, "abr": 4, "mai": 5, "jun": 6,
        "jul": 7, "ago": 8, "set": 9, "out": 10, "nov": 11, "dez": 12
    }

    # Padrão 1: "DD/MM a DD/MM" ou "DD/MM à DD/MM" (ex: "16/09 a 04/10", "16/11 à 29/11")
    m1 = re.match(r'^(\d{1,2})[\/\.](\d{1,2})\s*(?:a|à|-)\s*(\d{1,2})[\/\.](\d{1,2})$', s, re.IGNORECASE)
    if m1:
        d1, m1_m, d2, m2_m = map(int, m1.groups())
        try:
            dt1 = datetime(ano, m1_m, d1)
            dt2 = datetime(ano, m2_m, d2)
            dias = (dt2 - dt1).days + 1
            results.append({
                "dt_ini_iso": dt1.strftime("%Y-%m-%d"),
                "dt_fim_iso": dt2.strftime("%Y-%m-%d"),
                "dt_ini_br": dt1.strftime("%d/%m/%Y"),
                "dt_fim_br": dt2.strftime("%d/%m/%Y"),
                "dias": max(1, dias)
            })
            return results
        except Exception:
            pass

    # Padrão 2: "DD-DD/MM" (ex: "15-26/03")
    m2 = re.match(r'^(\d{1,2})\s*-\s*(\d{1,2})[\/\.](\d{1,2})$', s)
    if m2:
        d1, d2, m_val = map(int, m2.groups())
        try:
            dt1 = datetime(ano, m_val, d1)
            dt2 = datetime(ano, m_val, d2)
            dias = (dt2 - dt1).days + 1
            results.append({
                "dt_ini_iso": dt1.strftime("%Y-%m-%d"),
                "dt_fim_iso": dt2.strftime("%Y-%m-%d"),
                "dt_ini_br": dt1.strftime("%d/%m/%Y"),
                "dt_fim_br": dt2.strftime("%d/%m/%Y"),
                "dias": max(1, dias)
            })
            return results
        except Exception:
            pass

    # Padrão 3: "DD-DD/mes" (ex: "3-7/ago")
    m3 = re.match(r'^(\d{1,2})\s*-\s*(\d{1,2})[\/\.]([a-zçáéíóúãõ]{3,})$', s, re.IGNORECASE)
    if m3:
        d1, d2, m_str = m3.group(1), m3.group(2), m3.group(3).lower()
        m_val = meses_abreviados.get(m_str[:3], mes_padrao)
        try:
            dt1 = datetime(ano, m_val, int(d1))
            dt2 = datetime(ano, m_val, int(d2))
            dias = (dt2 - dt1).days + 1
            results.append({
                "dt_ini_iso": dt1.strftime("%Y-%m-%d"),
                "dt_fim_iso": dt2.strftime("%Y-%m-%d"),
                "dt_ini_br": dt1.strftime("%d/%m/%Y"),
                "dt_fim_br": dt2.strftime("%d/%m/%Y"),
                "dias": max(1, dias)
            })
            return results
        except Exception:
            pass

    # Padrão 4: "DD a DD" ou "DD à DD" (ex: "20 a 29", "6 a 15", "18 a 19")
    m4 = re.match(r'^(\d{1,2})\s*(?:a|à|-)\s*(\d{1,2})$', s, re.IGNORECASE)
    if m4:
        d1, d2 = map(int, m4.groups())
        try:
            dt1 = datetime(ano, mes_padrao, d1)
            dt2 = datetime(ano, mes_padrao, d2)
            dias = (dt2 - dt1).days + 1
            results.append({
                "dt_ini_iso": dt1.strftime("%Y-%m-%d"),
                "dt_fim_iso": dt2.strftime("%Y-%m-%d"),
                "dt_ini_br": dt1.strftime("%d/%m/%Y"),
                "dt_fim_br": dt2.strftime("%d/%m/%Y"),
                "dias": max(1, dias)
            })
            return results
        except Exception:
            pass

    # Padrão 5: Lista de dias (ex: "30 e 31", "16 e 17", "6, 7 e 9", "17,18,21 e 22")
    # Extrai todos os números
    day_numbers = [int(x) for x in re.findall(r'\b\d{1,2}\b', s)]
    if day_numbers:
        day_numbers.sort()
        # Se os dias forem consecutivos (ex: 30 e 31), agrupa em um só intervalo
        min_d = day_numbers[0]
        max_d = day_numbers[-1]
        try:
            dt1 = datetime(ano, mes_padrao, min_d)
            dt2 = datetime(ano, mes_padrao, max_d)
            dias = len(day_numbers)
            results.append({
                "dt_ini_iso": dt1.strftime("%Y-%m-%d"),
                "dt_fim_iso": dt2.strftime("%Y-%m-%d"),
                "dt_ini_br": dt1.strftime("%d/%m/%Y"),
                "dt_fim_br": dt2.strftime("%d/%m/%Y"),
                "dias": dias
            })
            return results
        except Exception:
            pass

    return results

def sync_ferias_from_excel(file_path_or_buffer) -> bool:
    """
    Lê a planilha oficial de Previsão de Férias da Bancada/Manutenção (.xlsx),
    extrai as abas anuais (2024, 2025, 2026, 2027...) e persiste no SQLite em 'ferias_bancada'.
    """
    from src.config import setup_logging, DEBUG_DIR_FERIAS
    logger = setup_logging(DEBUG_DIR_FERIAS / "sync_ferias.log", "sync_ferias")
    logger.info("Iniciando sincronização da planilha de férias da bancada...")

    setup_ferias_table()

    temp_file = tempfile.NamedTemporaryFile(delete=False, suffix=".xlsx")
    temp_path = temp_file.name
    temp_file.close()

    try:
        if isinstance(file_path_or_buffer, (str, Path)):
            p = Path(file_path_or_buffer)
            if not p.exists():
                logger.error(f"Arquivo de férias não encontrado: {file_path_or_buffer}")
                return False
            shutil.copy2(str(p), temp_path)
        elif hasattr(file_path_or_buffer, "read"):
            file_path_or_buffer.seek(0)
            with open(temp_path, "wb") as f_out:
                f_out.write(file_path_or_buffer.read())
        else:
            logger.error("Tipo de arquivo/buffer inválido fornecido para férias.")
            return False

        wb = openpyxl.load_workbook(temp_path, data_only=True)
    except Exception as e:
        logger.error(f"Erro ao abrir arquivo Excel de férias: {e}")
        return False
    finally:
        try:
            Path(temp_path).unlink(missing_ok=True)
        except Exception:
            pass

    records = []
    
    for sheet_name in wb.sheetnames:
        sheet_clean = str(sheet_name).strip()
        # Abas anuais: '2024', '2025', '2026', '2027', etc.
        if not re.match(r'^\d{4}$', sheet_clean):
            continue

        ano = int(sheet_clean)
        ws = wb[sheet_name]
        logger.info(f"Processando aba de férias do ano: {ano} ({ws.max_row} linhas x {ws.max_column} colunas)")

        # Localiza seções dentro da planilha
        current_section = "ferias" # 'ferias', 'licencas_compensacoes', 'recesso_forense'
        
        # Mapeamento de colunas de meses (geralmente colunas 2 a 13)
        month_headers = {}
        for c in range(2, 14):
            val = ws.cell(1, c).value
            m_num, m_name = _parse_month_from_header(val, ano)
            month_headers[c] = (m_num, m_name)

        for r in range(1, ws.max_row + 1):
            col1_val = ws.cell(r, 1).value
            if not col1_val:
                continue

            col1_str = str(col1_val).strip()

            if "Licenças" in col1_str or "Licencas" in col1_str or "Compensações" in col1_str:
                current_section = "licencas_compensacoes"
                # Atualiza cabeçalhos de mês para esta seção se houver
                for c in range(2, 14):
                    v_cell = ws.cell(r, c).value
                    if v_cell:
                        m_num, m_name = _parse_month_from_header(v_cell, ano)
                        month_headers[c] = (m_num, m_name)
                continue
            elif "Recesso Forense" in col1_str:
                current_section = "recesso_forense"
                continue
            elif col1_str.lower() in ["férias", "ferias", "legenda", "ano"]:
                continue

            membro = col1_str

            if current_section in ["ferias", "licencas_compensacoes"]:
                for c in range(2, 14):
                    cell = ws.cell(r, c)
                    cell_val = cell.value
                    if not cell_val:
                        continue
                    
                    val_str = str(cell_val).strip()
                    if not val_str or val_str in ["-", "nan", "none"]:
                        continue

                    mes_num, mes_nome = month_headers.get(c, (c - 1, f"Mês {c-1}"))

                    # Interpreta cores ou regras de status
                    # Padrão: Confirmada (verde), Prevista (amarelo/azul), Cancelada (vermelho)
                    status = "Confirmada"
                    cor_hex = "#10b981" # Verde padrão

                    if current_section == "licencas_compensacoes":
                        status = "Compensação/Licença"
                        cor_hex = "#3b82f6" # Azul

                    # Extrai períodos de datas
                    intervals = parse_periodo_cell(val_str, ano, mes_num)
                    if intervals:
                        for inter in intervals:
                            records.append({
                                "ano": ano,
                                "membro": membro,
                                "tipo_escala": current_section,
                                "mes": mes_num,
                                "mes_nome": mes_nome,
                                "periodo_bruto": val_str,
                                "data_inicio_iso": inter["dt_ini_iso"],
                                "data_fim_iso": inter["dt_fim_iso"],
                                "data_inicio_br": inter["dt_ini_br"],
                                "data_fim_br": inter["dt_fim_br"],
                                "dias": inter["dias"],
                                "status": status,
                                "cor_hex": cor_hex
                            })
                    else:
                        # Fallback se não conseguir extrair datas precisas
                        records.append({
                            "ano": ano,
                            "membro": membro,
                            "tipo_escala": current_section,
                            "mes": mes_num,
                            "mes_nome": mes_nome,
                            "periodo_bruto": val_str,
                            "data_inicio_iso": f"{ano}-{mes_num:02d}-01",
                            "data_fim_iso": f"{ano}-{mes_num:02d}-05",
                            "data_inicio_br": f"01/{mes_num:02d}/{ano}",
                            "data_fim_br": f"05/{mes_num:02d}/{ano}",
                            "dias": 5,
                            "status": status,
                            "cor_hex": cor_hex
                        })

            elif current_section == "recesso_forense":
                # Colunas 2, 3, 4: 1ª sem, 2ª sem, 3ª sem
                recesso_semanas = {
                    2: ("1ª Semana", "20/12", "26/12", 7),
                    3: ("2ª Semana", "27/12", "02/01", 7),
                    4: ("3ª Semana", "03/01", "06/01", 4)
                }
                for c in [2, 3, 4]:
                    cell = ws.cell(r, c)
                    val = cell.value
                    if val and str(val).strip().upper() == "X":
                        sem_nome, d_ini, d_fim, q_dias = recesso_semanas[c]
                        # Calcula datas ISO
                        if c in [2, 3]:
                            dt_ini_iso = f"{ano}-12-{d_ini.split('/')[0]}"
                            dt_fim_iso = f"{ano}-12-{d_fim.split('/')[0]}" if "12" in d_fim else f"{ano+1}-01-{d_fim.split('/')[0]}"
                        else:
                            dt_ini_iso = f"{ano+1}-01-{d_ini.split('/')[0]}"
                            dt_fim_iso = f"{ano+1}-01-{d_fim.split('/')[0]}"

                        records.append({
                            "ano": ano,
                            "membro": membro,
                            "tipo_escala": "recesso_forense",
                            "mes": 12 if c in [2, 3] else 1,
                            "mes_nome": f"Recesso ({sem_nome})",
                            "periodo_bruto": f"Recesso Forense - {sem_nome}",
                            "data_inicio_iso": dt_ini_iso,
                            "data_fim_iso": dt_fim_iso,
                            "data_inicio_br": f"{d_ini}/{ano}",
                            "data_fim_br": f"{d_fim}/{ano if '12' in d_fim else ano+1}",
                            "dias": q_dias,
                            "status": "Recesso Forense",
                            "cor_hex": "#8b5cf6" # Roxo
                        })

    logger.info(f"Total de registros de férias/licenças/recesso extraídos: {len(records)}")

    conn = get_connection()
    cursor = conn.cursor()
    cursor.execute("DELETE FROM ferias_bancada")

    for rec in records:
        cursor.execute("""
        INSERT INTO ferias_bancada (
            ano, membro, tipo_escala, mes, mes_nome, periodo_bruto,
            data_inicio_iso, data_fim_iso, data_inicio_br, data_fim_br,
            dias, status, cor_hex
        ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
        """, (
            rec["ano"], rec["membro"], rec["tipo_escala"], rec["mes"], rec["mes_nome"],
            rec["periodo_bruto"], rec["data_inicio_iso"], rec["data_fim_iso"],
            rec["data_inicio_br"], rec["data_fim_br"], rec["dias"], rec["status"], rec["cor_hex"]
        ))

    conn.commit()
    conn.close()
    logger.info("💾 Sincronização de férias finalizada com sucesso no banco SQLite!")
    return True

def get_ferias_df(ano: int | None = None, membro: str | None = None) -> pd.DataFrame:
    """Retorna DataFrame com as férias cadastradas no SQLite com filtros opcionais."""
    setup_ferias_table()
    conn = get_connection()
    query = "SELECT * FROM ferias_bancada WHERE 1=1"
    params = []
    if ano:
        query += " AND ano = ?"
        params.append(ano)
    if membro and membro != "Todos":
        query += " AND membro = ?"
        params.append(membro)
    query += " ORDER BY ano DESC, data_inicio_iso ASC"
    df = pd.read_sql_query(query, conn, params=params)
    conn.close()
    return df

def get_ferias_membros() -> list[str]:
    """Retorna a lista de membros cadastrados nas escalas de férias."""
    setup_ferias_table()
    conn = get_connection()
    cursor = conn.cursor()
    cursor.execute("SELECT DISTINCT membro FROM ferias_bancada WHERE membro IS NOT NULL AND membro != '' ORDER BY membro")
    membros = [r[0] for r in cursor.fetchall()]
    conn.close()
    return membros

def get_ferias_anos() -> list[int]:
    """Retorna a lista de anos disponíveis no banco de férias."""
    setup_ferias_table()
    conn = get_connection()
    cursor = conn.cursor()
    cursor.execute("SELECT DISTINCT ano FROM ferias_bancada ORDER BY ano DESC")
    anos = [r[0] for r in cursor.fetchall()]
    conn.close()
    return anos if anos else [2026, 2025, 2024]
