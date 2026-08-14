import re
import sys
import platform
from datetime import datetime
from pathlib import Path
import unicodedata
import pandas as pd
from fuzzywuzzy import process
import shutil, time, tempfile
from typing import cast
from xlsxwriter.workbook import Workbook as XlsxWorkbook # Alias para não confundir
import logging
from logging.handlers import RotatingFileHandler
import sys

from config import (
    INPUT_DIR_BRUTOS,
    OUTPUT_DIR_TRATADOS,
    DEBUG_DIR_PREPROCESS,
    setup_logging,
    save_df_to_excel_formatted,
    cleanup_old_files,
    clean_otrs_description
)

try:
    import win32com.client as win32
except ImportError:
    win32 = None

# --- Configuração de logging ---
logger = setup_logging(DEBUG_DIR_PREPROCESS / "preprocess_chamados.log", __name__)

# --- Excel auto-fit via COM on Windows ---
def autofit_excel_rows(filepath: Path):
    if platform.system() != "Windows" or win32 is None:
        return

    abs_path = filepath.resolve()
    
    logger.info(f"Auto-fit nas linhas de {abs_path} …")
    
    excel = win32.DispatchEx('Excel.Application')
    excel.Visible = False
    wb = excel.Workbooks.Open(str(abs_path))
    try:
        for sheet in wb.Sheets:
            sheet.UsedRange.Rows.AutoFit()
        wb.Save()
    
    finally:
        wb.Close(SaveChanges=True)
        excel.Quit()


# --- Safe Excel read with COM fallback ---
def safe_read_excel(path: Path) -> pd.DataFrame:
    try:
        return pd.read_excel(path, engine='openpyxl')
    
    except PermissionError:
        if win32 is None:
            raise RuntimeError(f"Não foi possível ler {path!r} e pywin32 não disponível.")
    
    abs_path = path.resolve()
    
    logger.error(f"Falha de permissão, copiando via COM: {abs_path}")
    
    tmp = tempfile.NamedTemporaryFile(suffix=abs_path.suffix, delete=False)
    tmp_path = Path(tmp.name)
    tmp.close()
    
    excel = win32.DispatchEx('Excel.Application')
    excel.Visible = False
    excel.DisplayAlerts = False
    try:
        wb = excel.Workbooks.Open(str(abs_path), ReadOnly=True, UpdateLinks=False, IgnoreReadOnlyRecommended=True)
        wb.SaveCopyAs(str(tmp_path))
        wb.Close(False)
    
    finally:
        excel.Quit()
    df = pd.read_excel(tmp_path, engine='openpyxl')
    tmp_path.unlink(missing_ok=True)
    
    return df


# --- Timestamp ---
ts = datetime.now().strftime('%Y-%m-%d_%H-%M-%S')




def parse_date_safely(v):
    if pd.isna(v) or not str(v).strip():
        return v
    s = str(v).strip()
    if re.match(r'^\d{1,2}/\d{1,2}/\d{4}', s):
        dt = pd.to_datetime(s, dayfirst=True, errors='coerce')
    else:
        dt = pd.to_datetime(s, errors='coerce')
    if pd.isna(dt):
        return s
    return dt.strftime('%Y-%m-%d %H:%M:%S')


def normalize_text(text: str) -> str:
    text = unicodedata.normalize('NFKD', str(text)).lower()
    text = ''.join(c for c in text if not unicodedata.combining(c))
    text = re.sub(r'[^\w\s-]', '', text)
    
    return re.sub(r'\s+', ' ', text).strip()


def prepare_unidades_lookup():
    """Carrega o DataFrame de unidades preferencialmente do SQLite, utilizando o Excel como fallback."""
    units_df = pd.DataFrame()
    try:
        from database import get_unidades_df
        units_df = get_unidades_df()
    except Exception as e:
        logger.warning(f"Não foi possível carregar unidades do SQLite: {e}")

    if units_df.empty:
        units_file = INPUT_DIR_BRUTOS / "Unidades_MPMS.xlsx"
        if units_file.exists():
            units_df = pd.read_excel(units_file)
        else:
            logger.error(f"Erro: não foi possível carregar unidades do SQLite nem do arquivo {units_file}")
            sys.exit(1)

    # Limpa linhas com valores nulos para evitar casamentos falsos de nulos
    units_df = units_df.dropna(subset=['Setor', 'Unidade (Prédio)'])
    units_df['setor_normalizado'] = units_df['Setor'].apply(normalize_text)
    units_df['prédio_normalizado'] = units_df['Unidade (Prédio)'].apply(normalize_text)
    
    return units_df


def match_unidade(row: pd.Series, units_df: pd.DataFrame, base: str) -> pd.Series:
    val = row.get('Unidade', '')
    if pd.isna(val):
        return pd.Series()
    val_str = str(val).strip()
    # Ignora valores vazios, genéricos ou de falha do AD para evitar falso positivo no fuzzy matching
    if not val_str or val_str.lower() in (
        'nan', 'n/a', 'nao encontrada no ad', 'sem departamento', 
        'não encontrado no ad', 'cadastro incompleto (ad)', 'erro na consulta',
        'cadastro incompleto', 'nao encontrado', 'não encontrada'
    ):
        return pd.Series()

    query = normalize_text(val_str)
    if not query or query == 'nan':
        return pd.Series()

    matches = process.extractBests(query, units_df['setor_normalizado'], score_cutoff=75, limit=1)
    
    if matches:
        best = matches[0][0]
        return units_df[units_df['setor_normalizado']==best].iloc[0]
    
    return pd.Series()


def enrich_with_unidades(df: pd.DataFrame, base: str) -> pd.DataFrame:
    units_df = prepare_unidades_lookup()
    df = df.copy()
    
    # Listas para armazenar os dados processados
    lista_siglas = []
    lista_locais = []
    lista_unidades = []

    # Iteramos sobre as linhas
    for _, row in df.iterrows():
        match = match_unidade(row, units_df, base)
        
        if not match.empty:
            # Se achou correspondência, pegamos os valores
            sigla_encontrada = match['Sigla']
            local_encontrado = match['Unidade (Prédio)']
            
            lista_siglas.append(sigla_encontrada)
            lista_locais.append(local_encontrado)
            lista_unidades.append(sigla_encontrada) # Sobrescreve 'Unidade' com a Sigla
        else:
            # Se não achou, preenche com vazio ou mantém o original
            lista_siglas.append("")
            lista_locais.append("")
            # Mantém o valor original da coluna 'Unidade' se existir, senão vazio
            lista_unidades.append(row.get('Unidade', ''))

    # Atribuição direta (O Pandas faz isso instantaneamente e o Pylance entende perfeitamente)
    df['Sigla'] = lista_siglas
    df['Cidade - Prédio'] = lista_locais
    df['Unidade'] = lista_unidades
    
    return df


# --- Process OTRS ---
def process_otrs(ts: str) -> pd.DataFrame:
    files = sorted(INPUT_DIR_BRUTOS.glob("Chamados_OTRS_*.xlsx"))
    
    if not files:
        logger.info("Nenhum arquivo OTRS encontrado")
        sys.exit(1)
    path = files[-1]
    
    logger.info(f"Processando OTRS: {path.name}")
    
    df = safe_read_excel(path)
    df['Descrição'] = df['Descrição'].apply(clean_otrs_description)
    df['Base'] = 'OTRS'
    df = enrich_with_unidades(df, base='OTRS')
    
    # Adiciona fallback para Título caso não exista por algum motivo
    df['Título'] = df.get('Título', '').fillna('')
    
    if 'Data Criação' in df.columns:
        df['Data Criação'] = df['Data Criação'].apply(parse_date_safely)
    
    if 'Comentários' not in df.columns:
        df['Comentários'] = '[]'
    if 'IP_Origem' not in df.columns:
        df['IP_Origem'] = ""
    if 'Hostname' not in df.columns:
        df['Hostname'] = ""
    if 'Link' not in df.columns:
        df['Link'] = ""
    
    # colunas finais
    cols = ["Chamado#","Nome do Usuário","ID do Cliente","Data Criação",
            "Cidade - Prédio","Unidade","Descrição","Base","Título","IP_Origem","Hostname","Link","Comentários"]
    
    return df[cols]


# --- Process CitSmart ---
def process_citsmart(ts: str) -> pd.DataFrame:
    files = sorted(INPUT_DIR_BRUTOS.glob("Chamados_CitSmart_*.xlsx"))
    
    if not files:
        logger.info("Nenhum arquivo CitSmart encontrado")
        sys.exit(1)
    path = files[-1]
    
    logger.info(f"Processando CitSmart: {path.name}")
    
    df = safe_read_excel(path)
    df['Data Criação'] = df['Data Criação'].astype(str)
    df['Data Criação'] = df['Data Criação'].apply(parse_date_safely)
    df['Base'] = 'CitSmart'
    df = enrich_with_unidades(df, base='CitSmart')
    
    # CitSmart não tem Título nativo, então criamos vazio
    df['Título'] = ""
    
    if 'Comentários' not in df.columns:
        df['Comentários'] = '[]'
    if 'IP_Origem' not in df.columns:
        df['IP_Origem'] = ""
    if 'Hostname' not in df.columns:
        df['Hostname'] = ""
    if 'ID do Cliente' not in df.columns:
        df['ID do Cliente'] = ""
    if 'Link' not in df.columns:
        df['Link'] = ""
    
    cols = ["Chamado#","Nome do Usuário","ID do Cliente","Data Criação",
            "Cidade - Prédio","Unidade","Descrição","Base","Título","IP_Origem","Hostname","Link","Comentários"]
    
    return df[cols]


# --- Main ---
def main():
    
    otrs_df = process_otrs(ts)
    citsmart_df = process_citsmart(ts)
    
    # salvar ambos
    for name, df in [('OTRS', otrs_df), ('CitSmart', citsmart_df)]:
        out = OUTPUT_DIR_TRATADOS / f"{name}_tratado_{ts}.xlsx"
        widths = {col: 25 for col in df.columns}
        widths['Descrição'] = 100
        save_df_to_excel_formatted(
            df, out, sheet_name=name,
            widths=widths, wrap_cols=['Descrição', 'Comentários'], height_col='Descrição'
        )
        autofit_excel_rows(out)
    
    # unificar
    combined = pd.concat([otrs_df, citsmart_df], ignore_index=True)
    
    # =========================================================
    # NOVO: Padronização de datas e estrutura no pré-processamento
    # =========================================================
    # 1. Remove chamados repetidos, mantendo apenas a ocorrência mais recente
    tamanho_antes = len(combined)
    combined = combined.drop_duplicates(subset=['Chamado#'], keep='last')
    tamanho_depois = len(combined)
    
    if tamanho_antes != tamanho_depois:
        logger.info(f"⚠️ {tamanho_antes - tamanho_depois} chamados duplicados foram removidos!")
    
    # 2. Padroniza a Data de Criação em formato ISO (YYYY-MM-DD HH:MM:SS) sem sobrescrever com parsing ambíguo
    if 'Data Criação' in combined.columns:
        combined['Data Criação'] = combined['Data Criação'].apply(parse_date_safely)

    # 3. Limpa espaços extras no começo e no fim dos IDs e Nomes
    #if 'Chamado#' in combined.columns:
    #    combined['Chamado#'] = combined['Chamado#'].astype(str).str.strip()
        
    if 'Nome do Usuário' in combined.columns:
        # Tira espaços em branco sobrando e deixa a primeira letra de cada nome maiúscula
        combined['Nome do Usuário'] = combined['Nome do Usuário'].astype(str).str.strip().str.title()

    # 4. Adiciona as colunas vazias para manter o padrão estrutural da Master
    for col in ['Ramal', 'Andamento']:
        if col not in combined.columns:
            combined[col] = ""
            
    # 5. Organiza a ordem base das colunas (O classificador só vai espremer a TAG no meio depois)
    colunas_ordem = ['Chamado#', 'Nome do Usuário', 'ID do Cliente', 'Data Criação', 'Cidade - Prédio', 'Unidade', 'Ramal', 'Andamento', 'Descrição', 'Base', 'Título', 'IP_Origem', 'Hostname', 'Link', 'Comentários']
    colunas_existentes = [c for c in colunas_ordem if c in combined.columns]
    combined = combined[colunas_existentes]
    # =========================================================
    
    out = OUTPUT_DIR_TRATADOS / f"Chamados_Unificados_{ts}.xlsx"

    logger.info(f"Processando Unificado: Chamados_Unificados_{ts}.xlsx")
    
    widths = {col: 25 for col in combined.columns}
    widths['Descrição'] = 100
    save_df_to_excel_formatted(
        combined, out, sheet_name='Unificados',
        widths=widths, wrap_cols=['Descrição', 'Comentários'], height_col='Descrição'
    )
    autofit_excel_rows(out)
    
    # Limpeza de planilhas unificadas antigas (mantém no máximo as 10 mais recentes)
    cleanup_old_files(OUTPUT_DIR_TRATADOS, "Chamados_Unificados_*.xlsx", keep_count=10)
    cleanup_old_files(OUTPUT_DIR_TRATADOS, "OTRS_tratado_*.xlsx", keep_count=10)
    cleanup_old_files(OUTPUT_DIR_TRATADOS, "CitSmart_tratado_*.xlsx", keep_count=10)
    
    logger.info("Script finalizado!")
    # logger.info(f"SUCESSO! Total de {len(todos_os_dados)} chamados salvos em: {file}")


if __name__ == '__main__':
    try:
        main()
    except Exception as e:
        logger.exception(f"Erro crítico no pré-processamento de chamados: {e}")
        sys.exit(1)