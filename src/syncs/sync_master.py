import sys
from pathlib import Path
import pandas as pd
import shutil
from typing import Tuple, Any

root_dir = Path(__file__).resolve().parent.parent.parent
src_dir = Path(__file__).resolve().parent.parent
if str(root_dir) not in sys.path:
    sys.path.insert(0, str(root_dir))
if str(src_dir) not in sys.path:
    sys.path.insert(0, str(src_dir))

from config import (
    DEBUG_DIR_SYNC, OUTPUT_DIR_PRONTO,
    MASTER_FILE_PATH, TREINO_PATH,
    setup_logging
)

try:
    import win32com.client as win32
    from win32com.client import constants
except ImportError:
    win32 = None
    constants = None

logger = setup_logging(DEBUG_DIR_SYNC / "sync_master.log", __name__)

def format_excel(excel_data: Tuple[Any, Any], fechar_apos=True):
    excel_app, wb = excel_data
    ws = wb.Sheets(1)
    used = ws.UsedRange

    TAG_COLORS = {
        "BACKUP": "#dd5358",
        "EVENTO": "#ce66ce",
        "FORMATAÇÃO": "#d38a62",
        "GARANTIA": "#518bbb",
        "IMPRESSORA": "#C6EFCE",
        "INSTALAÇÃO HARDWARE": "#FCE4D6",
        "INSTALAÇÃO SOFTWARE": "#86BEEE",
        "MANUTENÇÃO": "#E9CF69",
        "MONITOR": "#cbdd6f",
        "MUDANÇA": "#21ffe0",
        "PREPARAÇÃO COMPUTADORES": "#f09c72",
        "REDE": "#B7F391",
        "SOLICITAÇÃO SSD": "#f5a89b",
        "SUPORTE": "#FFE699",
        "TELEFONIA FIXA": "#e273a1",
        "VIAGEM": "#61e7c6",
        "VISTORIA CPDS": "#b2740e",
    }

    COLUMN_WIDTHS = {
        1: 15,
        2: 25,
        3: 20,
        4: 30,
        5: 25,
        6: 25,
        7: 12,
        8: 20,
        9: 100,
        10: 15,
    }

    def hex_to_excel_color(hex_str):
        h = hex_str.lstrip('#')
        r, g, b = int(h[0:2], 16), int(h[2:4], 16), int(h[4:6], 16)
        return r + (g * 256) + (b * 65536)

    for col_idx, width in COLUMN_WIDTHS.items():
        try:
            ws.Columns(col_idx).ColumnWidth = width
        except Exception:
            pass

    try:
        ws.Columns(3).NumberFormat = "dd/mm/aaaa hh:mm:ss"
    except Exception:
        pass

    tag_col_idx = -1
    for c in range(1, used.Columns.Count + 1):
        if ws.Cells(1, c).Value == "TAG":
            tag_col_idx = c
            break

    header = ws.Range(ws.Cells(1, 1), ws.Cells(1, used.Columns.Count))
    header.Interior.Color = 0x808080
    header.Font.Bold = True
    header.Font.Color = 0xFFFFFF

    if tag_col_idx != -1:
        for r in range(2, used.Rows.Count + 1):
            cell_value = ws.Cells(r, tag_col_idx).Value
            if cell_value:
                tag_name = str(cell_value).strip().upper()
                if tag_name in TAG_COLORS:
                    row_range = ws.Range(ws.Cells(r, 1), ws.Cells(r, used.Columns.Count))
                    row_range.Interior.Color = hex_to_excel_color(TAG_COLORS[tag_name])

    try:
        used.WrapText = True
    except Exception:
        pass

    if used.Rows.Count > 1:
        try:
            data_rows = ws.Range(ws.Cells(2, 1), ws.Cells(used.Rows.Count, used.Columns.Count))
            data_rows.RowHeight = 20
        except Exception:
            pass

    try:
        used.Rows.AutoFit()
    except Exception:
        pass

    wb.Save()
    
    if fechar_apos:
        wb.Close(SaveChanges=True)
        try:
            if excel_app.Workbooks.Count == 0:
                excel_app.Quit()
        except Exception:
            pass

def read_excel_com_to_df(ws) -> pd.DataFrame:
    used_range = ws.UsedRange.Value
    if not used_range:
        return pd.DataFrame()
    
    cleaned_data = []
    for row in used_range:
        cleaned_row = []
        for cell in row:
            if hasattr(cell, 'tzinfo') and cell.tzinfo is not None:
                cleaned_row.append(cell.replace(tzinfo=None))
            else:
                cleaned_row.append(cell)
        cleaned_data.append(cleaned_row)
    
    df = pd.DataFrame(cleaned_data[1:], columns=cleaned_data[0])
    df.dropna(how='all', inplace=True)
    return df

def sync_to_master(novo_excel_path: Path, master_excel_path: Path) -> Tuple[Any, Any, bool, bool]:
    def clean_ticket_id(series):
        return series.astype(str).str.strip().str.replace(r'\.0$', '', regex=True)

    df_tagged_novo = pd.read_excel(novo_excel_path)
    df_tagged_novo['Chamado#'] = clean_ticket_id(df_tagged_novo['Chamado#'])

    try:
        excel = win32.GetActiveObject("Excel.Application")
    except Exception:
        excel = win32.Dispatch("Excel.Application")
        excel.Visible = False
        
    excel.DisplayAlerts = False
    wb_master = None
    was_already_open = False

    for wb in excel.Workbooks:
        if wb.Name.lower() == master_excel_path.name.lower():
            wb_master = wb
            was_already_open = True
            break

    if wb_master is None:
        wb_master = excel.Workbooks.Open(str(master_excel_path.resolve()))
        
    ws_tagged = wb_master.Sheets(1)
    df_master_tagged = read_excel_com_to_df(ws_tagged)

    if not df_master_tagged.empty and 'Chamado#' in df_master_tagged.columns:
        df_master_tagged['Chamado#'] = clean_ticket_id(df_master_tagged['Chamado#'])
        chamados_master = {str(cid).strip() for cid in df_master_tagged['Chamado#'] if cid and str(cid).strip().lower() not in ('nan', 'none', '')}
    else:
        chamados_master = set()

    chamados_novos = {str(cid).strip() for cid in df_tagged_novo['Chamado#'] if cid and str(cid).strip().lower() not in ('nan', 'none', '')}

    if not chamados_novos:
        logger.warning("⚠️ ALERTA DE SEGURANÇA: A lista de novos chamados está VAZIA! Sincronização abortada.")
        return excel, wb_master, False, was_already_open

    houve_exclusao = False
    fechados_ids = chamados_master - chamados_novos
    
    if fechados_ids:
        logger.info(f"Chamados fechados identificados: {len(fechados_ids)}. Movendo para Treino e limpando da Master...")
        df_fechados = df_master_tagged[df_master_tagged['Chamado#'].isin(fechados_ids)].copy()
        
        try:
            df_fechados_clean = df_fechados.dropna(how='all') 
            
            if not df_fechados_clean.empty:
                if TREINO_PATH.exists():
                    df_treino_atual = pd.read_excel(TREINO_PATH)
                    df_treino_novo = pd.concat([df_treino_atual, df_fechados_clean], ignore_index=True)
                else:
                    df_treino_novo = df_fechados_clean
                    
                df_treino_novo = df_treino_novo.drop_duplicates(subset=['Chamado#'], keep='last')
                df_treino_novo = df_treino_novo.fillna("")
                df_treino_novo.to_excel(TREINO_PATH, index=False)
                logger.info("Chamados fechados adicionados ao dataset de treino com sucesso.")
        except Exception as e:
            logger.error(f"Erro ao salvar chamados fechados no treino: {e}", exc_info=True)

        last_row_master = ws_tagged.Cells(ws_tagged.Rows.Count, 1).End(-4162).Row
        linhas_deletadas = 0
        
        for r in range(last_row_master, 1, -1):
            celula_id = ws_tagged.Cells(r, 1).Value
            raw_id = str(celula_id).strip()
            if raw_id.endswith('.0'):
                raw_id = raw_id[:-2]
                
            if raw_id and raw_id.lower() not in ('none', 'nan', '') and raw_id in fechados_ids:
                ws_tagged.Rows(r).Delete()
                linhas_deletadas += 1
                
        logger.info(f"{linhas_deletadas} linhas apagadas da Planilha Master.")
        houve_exclusao = True

    novos_ids = chamados_novos - chamados_master
    if not novos_ids:
        logger.info("Nenhum chamado novo para adicionar à Master.")
        return excel, wb_master, houve_exclusao, was_already_open

    df_apenas_novos = df_tagged_novo[df_tagged_novo['Chamado#'].isin(novos_ids)].copy()
    
    if df_master_tagged is not None and len(df_master_tagged.columns) > 0:
        colunas_master = df_master_tagged.columns.tolist()
        
        for col in colunas_master:
            if col not in df_apenas_novos.columns:
                df_apenas_novos[col] = ""
                
        df_apenas_novos = df_apenas_novos[colunas_master]
    
    if 'Data Criação' in df_apenas_novos.columns:
        try:
            dt_col = pd.to_datetime(df_apenas_novos['Data Criação'], errors='coerce')
            df_apenas_novos['Data Criação'] = dt_col.dt.strftime('%Y-%m-%d %H:%M:%S').fillna(df_apenas_novos['Data Criação'])
        except Exception as e:
            logger.error(f"Erro ao sanitizar Data Criação para ISO: {e}", exc_info=True)
    
    df_apenas_novos = df_apenas_novos.fillna("")

    data_list = df_apenas_novos.values.tolist()
    
    last_row = ws_tagged.UsedRange.Rows.Count

    if last_row == 1 and ws_tagged.Cells(1,1).Value is None:
        start_row = 1
    else:
        start_row = last_row + 1

    data_list = df_apenas_novos.values.tolist()
    if start_row == 1:
        headers = df_apenas_novos.columns.tolist()
        ws_tagged.Range(ws_tagged.Cells(1, 1), ws_tagged.Cells(1, len(headers))).Value = headers
        start_row = 2

    end_row = start_row + len(data_list) - 1
    end_col = len(df_apenas_novos.columns)
    
    ws_tagged.Range(ws_tagged.Cells(start_row, 1), ws_tagged.Cells(end_row, end_col)).Value = data_list
    
    if ws_tagged.ListObjects.Count > 0:
        tabela = ws_tagged.ListObjects(1)
        coluna_inicial = tabela.Range.Column
        linha_inicial = tabela.Range.Row
        
        novo_range = ws_tagged.Range(
            ws_tagged.Cells(linha_inicial, coluna_inicial), 
            ws_tagged.Cells(end_row, end_col)
        )
        tabela.Resize(novo_range)

    logger.info(f"Adicionados {len(novos_ids)} novos chamados à Master.")

    return excel, wb_master, True, was_already_open

def main():
    logger.info("=== INICIANDO SINCRONIZAÇÃO MASTER ===")
    
    arquivos = list(OUTPUT_DIR_PRONTO.glob("Chamados_Tagged_*.xlsx"))
    if not arquivos:
        logger.error("Nenhum arquivo Tagged encontrado.")
        sys.exit(1)
        
    recente = max(arquivos, key=lambda p: p.stat().st_mtime)
    logger.info(f"Lendo base classificada: {recente.name}")

    if not MASTER_FILE_PATH.exists():
        logger.warning(f"Master não encontrado. Usando arquivo atual como base.")
        shutil.copy(recente, MASTER_FILE_PATH)

    excel_app, wb_master, changed, was_already_open = sync_to_master(recente, MASTER_FILE_PATH)

    if changed:
        logger.info("Houve alterações, aplicando formatação visual...")
        format_excel((excel_app, wb_master), fechar_apos=not was_already_open)
        logger.info("Sincronização e formatação concluídas com sucesso!")
    else:
        logger.info("Sem alterações. Fechando recursos...")
        if not was_already_open:
            wb_master.Close(SaveChanges=True)
            try:
                if excel_app.Workbooks.Count == 0:
                    excel_app.Quit()
            except Exception:
                pass
                
    logger.info("=== FIM DA SINCRONIZAÇÃO ===")

if __name__ == "__main__":
    main()
