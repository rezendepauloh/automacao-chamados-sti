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

from config import (
    INPUT_DIR_BRUTOS,
    OUTPUT_DIR_TRATADOS
)

try:
    import win32com.client as win32
except ImportError:
    win32 = None


def debug_print(msg):
    ts = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    print(f"[DEBUG {ts}] {msg}")


# --- Excel auto-fit via COM on Windows ---
def autofit_excel_rows(filepath: Path):
    if platform.system() != "Windows" or win32 is None:
        return

    abs_path = filepath.resolve()
    
    debug_print(f"Auto-fit nas linhas de {abs_path} …")
    
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
    
    debug_print(f"Falha de permissão, copiando via COM: {abs_path}")
    
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


# --- OTRS description cleaning ---
def clean_otrs_description(desc: str) -> str:
    if pd.isna(desc):
        return ""
    text = str(desc).strip()
    text = re.sub(r'(?si)^.*?^Descrição:\s*[\r\n]+', '', text, flags=re.MULTILINE)
    text = re.sub(r'(?si)[\r\n]+(?:Para acompanhamento.*|É possível acompanhar.*)$', '', text)
    text = re.sub(r'(?m)^\.\.\.\s*$', '', text)
    text = re.sub(r'(?m)^Prazo:.*$', '', text)
    text = re.sub(
        r'(?ms)^Att\.\s*\r?\n--\s*\r?\n(?:.*\r?\n)*?Telefones:.*?3318-8959\s*(?:\r?\n)?',
        '---\n', text
    )
    text = re.sub(
        r'(?ms)^Att\.\s*\r?\n--\s*\r?\n(?:.*\r?\n)*?Telefone:.*?3318-5503\s*(?:\r?\n)?',
        '---\n', text
    )
    text = re.sub(
        r'(?mi)(Descrição do pedido:\s*)([\s\S]+)$',
        lambda m: m.group(1) + m.group(2).replace('\n', ' '),
        text
    )
    parts = re.split(r'(?m)^#2\b', text, maxsplit=1)
    block1 = parts[0]
    cleaned = []
    
    for line in block1.splitlines():
        l = line.strip()
        if not l:
            continue
        if re.fullmatch(r'[A-Z]{1,2}', l):
            continue
        if l.lower().startswith('responder a nota') or l.lower() in ('imprimir','dividir'):
            continue
        cleaned.append(l)
    
    return '\n'.join(cleaned)


# --- Normalize units lookup ---
def normalize_text(text: str) -> str:
    text = unicodedata.normalize('NFKD', str(text)).lower()
    text = ''.join(c for c in text if not unicodedata.combining(c))
    text = re.sub(r'[^\w\s-]', '', text)
    
    return re.sub(r'\s+', ' ', text).strip()


def prepare_unidades_lookup():
    units_file = INPUT_DIR_BRUTOS / "Unidades_MPMS.xlsx"
    if not units_file.exists():
        debug_print(f"Erro: não encontrei {units_file}")
        sys.exit(1)
    units_df = pd.read_excel(units_file)
    units_df['setor_normalizado'] = units_df['Setor'].apply(normalize_text)
    units_df['prédio_normalizado'] = units_df['Unidade (Prédio)'].apply(normalize_text)
    
    return units_df


def match_unidade(row: pd.Series, units_df: pd.DataFrame, base: str) -> pd.Series:
    query = normalize_text(row['Unidade'])
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
        debug_print("Nenhum arquivo OTRS encontrado")
        sys.exit(1)
    path = files[-1]
    
    debug_print(f"Processando OTRS: {path.name}")
    
    df = safe_read_excel(path)
    df['Descrição'] = df['Descrição'].apply(clean_otrs_description)
    df['Base'] = 'OTRS'
    df = enrich_with_unidades(df, base='OTRS')
    
    # colunas finais
    cols = ["Chamado#","Nome do Usuário","ID do Cliente","Data Criação",
            "Cidade - Prédio","Unidade","Descrição","Base"]
    
    return df[cols]


# --- Process CitSmart ---
def process_citsmart(ts: str) -> pd.DataFrame:
    files = sorted(INPUT_DIR_BRUTOS.glob("Chamados_CitSmart_*.xlsx"))
    
    if not files:
        debug_print("Nenhum arquivo CitSmart encontrado")
        sys.exit(1)
    path = files[-1]
    
    debug_print(f"Processando CitSmart: {path.name}")
    
    df = safe_read_excel(path)
    df['Data Criação'] = df['Data Criação'].astype(str)
    df['Base'] = 'CitSmart'
    df = enrich_with_unidades(df, base='CitSmart')
    cols = ["Chamado#","Nome do Usuário","Data Criação",
            "Cidade - Prédio","Unidade","Descrição","Base"]
    
    return df[cols]


# --- Main ---
def main():
    
    otrs_df = process_otrs(ts)
    citsmart_df = process_citsmart(ts)
    
    # salvar ambos
    for name, df in [('OTRS', otrs_df), ('CitSmart', citsmart_df)]:
        out = OUTPUT_DIR_TRATADOS / f"{name}_tratado_{ts}.xlsx"
        with pd.ExcelWriter(out, engine='xlsxwriter') as writer:
            df.to_excel(writer, sheet_name=name, index=False)
            wb = cast(XlsxWorkbook, writer.book)
            ws = writer.sheets[name]
            wrap = wb.add_format({'text_wrap': True})
            widths = {col:25 for col in df.columns}
            widths['Descrição'] = 100
            
            for i,col in enumerate(df.columns):
                ws.set_column(i,i,widths.get(col,15), wrap if col=='Descrição' else None)
            
            for r,cell in enumerate(df['Descrição'], start=1):
                ws.set_row(r,15*(str(cell).count('\n')+1))
        autofit_excel_rows(out)
    
    # unificar
    combined = pd.concat([otrs_df, citsmart_df], ignore_index=True)
    out = OUTPUT_DIR_TRATADOS / f"Chamados_Unificados_{ts}.xlsx"
    
    with pd.ExcelWriter(out, engine='xlsxwriter') as writer:
        combined.to_excel(writer, sheet_name='Unificados', index=False)
        wb = cast(XlsxWorkbook, writer.book)
        ws = writer.sheets['Unificados']
        wrap = wb.add_format({'text_wrap': True})
        widths = {col:25 for col in combined.columns}
        widths['Descrição'] = 100
        
        for i,col in enumerate(combined.columns):
            ws.set_column(i,i,widths.get(col,15), wrap if col=='Descrição' else None)
        
        for r,cell in enumerate(combined['Descrição'], start=1):
            ws.set_row(r,15*(str(cell).count('\n')+1))
    autofit_excel_rows(out)
    
    debug_print("Script finalizado!")


if __name__ == '__main__':
    main()