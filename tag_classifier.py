import sys
import re
import logging
import numpy as np
from pathlib import Path
from datetime import datetime
import pandas as pd
import joblib
import shutil

from sklearn.pipeline import Pipeline
from sklearn.naive_bayes import MultinomialNB, ComplementNB
from sklearn.svm import LinearSVC
from sklearn.ensemble import RandomForestClassifier
from sklearn.feature_extraction.text import TfidfVectorizer
from sklearn.model_selection import (
    StratifiedKFold, GridSearchCV
)
from sklearn.metrics import classification_report
from typing import cast, Dict, Any, Tuple

import spacy
from spacy.lang.pt.stop_words import STOP_WORDS
# from xlsxwriter.utility import xl_col_to_name
from openpyxl import load_workbook
from openpyxl.utils.dataframe import dataframe_to_rows

# from openpyxl import load_workbook
from config import (
    TREINO_PATH, MODEL_PATH, DEBUG_DIR_TAG,
    OUTPUT_DIR_TRATADOS, OUTPUT_DIR_PRONTO,
    MASTER_FILE_PATH
)

# Excel COM constants import
import platform
try:
    import win32com.client as win32
    from win32com.client import constants
except ImportError:
    win32 = None
    constants = None  # type: ignore


# --------------------------------------------------------------------------
# Configurações iniciais
# --------------------------------------------------------------------------
nlp = spacy.load("pt_core_news_sm", disable=["parser", "ner"])
pt_stop = STOP_WORDS

logging.basicConfig(
    level=logging.DEBUG,
    # Adicionamos os colchetes [] e removemos os milissegundos do datefmt para ficar limpo
    format='[%(asctime)s] [%(levelname)s] %(message)s',
    datefmt='%Y-%m-%d %H:%M:%S',
    handlers=[
        logging.FileHandler(DEBUG_DIR_TAG / "tag_classifier.log", encoding='utf-8'),
        logging.StreamHandler(sys.stdout)
    ]
)
logger = logging.getLogger(__name__)

# --------------------------------------------------------------------------
# 1) Extração específica para OTRS
# --------------------------------------------------------------------------
def process_otrs_header(text: str) -> str | None:
    """
    Tenta extrair Título do Pedido e Descrição do pedido.
    Se não encontrar nenhum, retorna None para fallback.
    """
    titulo = None
    descricao = None

    m1 = re.search(r"T[ií]tulo do Pedido\s*[:\-]\s*(.*)", text, flags=re.IGNORECASE)
    if m1:
        titulo = m1.group(1).strip()

    m2 = re.search(
        r"Descri[cç][aã]o do pedido\s*[:\-]\s*(.*)",
        text,
        flags=re.IGNORECASE | re.DOTALL
    )
    if m2:
        descricao = m2.group(1).strip()

    if not titulo and not descricao:
        return None

    parts = []
    if titulo:
        parts.append(f"titulo de o pedido {titulo}")
    if descricao:
        parts.append(f"descricao de o pedido {descricao}")
    return "\n".join(parts)


# --------------------------------------------------------------------------
# 2) Funções de pré-processamento
# --------------------------------------------------------------------------
def clean_accents(text: str) -> str:
    # Substituições de acentos e ordinais
    accent_map = {
        'á':'a','à':'a','â':'a','ã':'a','ä':'a',
        'é':'e','è':'e','ê':'e','ë':'e',
        'í':'i','ì':'i','î':'i','ï':'i',
        'ó':'o','ò':'o','ô':'o','õ':'o','ö':'o',
        'ú':'u','ù':'u','û':'u','ü':'u',
        'ç':'c','ñ':'n',
        'Á':'A','À':'A','Â':'A','Ã':'A','Ä':'A',
        'É':'E','È':'E','Ê':'E','Ë':'E',
        'Í':'I','Ì':'I','Î':'I','Ï':'I',
        'Ó':'O','Ò':'O','Ô':'O','Õ':'O','Ö':'O',
        'Ú':'U','Ù':'U','Û':'U','Ü':'U',
        'Ç':'C','Ñ':'N',
        # Indicadores ordinais
        'ª':'a', 'º':'o'
    }
    return ''.join(accent_map.get(c, c) for c in text)


def conserve_tech_terms(text: str) -> str:
    tech_map = {
        r"\bssd\b": "ssd",
        r"\bcpu\b": "cpu",
        r"\bhd\b": "hd",
        r"\bram\b": "memoriaram",
        r"\bip\b": "enderecoip",
        r"\bti\b": "tecnologiainformacao",
        r"\bmb\b": "megabyte",
        r"\bgb\b": "gigabyte"
    }
    for patt, rep in tech_map.items():
        text = re.sub(patt, rep, text, flags=re.IGNORECASE)
    return text


def lemmatize_text(text: str) -> str:
    doc = nlp(text)
    lemmas = []
    
    for token in doc:
        # aceita palavras (is_alpha) ou números (like_num)
        if (token.is_alpha or token.like_num) and token.lemma_.lower() not in pt_stop:
            
            # mantém o número literal, e lematiza as palavras
            lem = token.text if token.like_num else token.lemma_.lower()
            lemmas.append(lem)
    
    return " ".join(lemmas)

def remove_generic_patterns(text: str) -> str:
    # 1) Normalize: já deve estar sem acentos neste ponto
    # 2) Honoríficos
    honorifics = [
        r'\bexcelent[íi]ssimo\b', r'\bEx\.º\b',
        r'\bexcelent[íi]ssima\b', r'\bEx\.ª\b',
        r'\bdr\b', r'\bdra\b', r'\bdrª\b',
        r'\bsr\b', r'\bsra\b', r'\bsrª\b'
    ]
    for pat in honorifics:
        text = re.sub(pat, '', text, flags=re.IGNORECASE)

    # 2) Prefixos de saudação (apenas o prefixo)
    prefix_patterns = [
        r'^(?:ol[áa]|bom dia|boa tarde|boa noite|prezados?|caros?)[\s,;:!\-]+'
    ]

    # 3) Sufixos de cortesia
    suffix_patterns = [
        r"[\s,;:]+(?:[Aa]tenciosament(?:e)?|[Aa]tt|cordiais? saudações|grato(?:)?|obrigad[oa]|desde já agradeço|[Aa]gradecidamente)[\s\.!]*$"
    ]

    for pat in prefix_patterns + suffix_patterns:
        text = re.sub(pat, '', text, flags=re.IGNORECASE)

    return text

# --------------------------------------------------------------------------
# 3) Pipeline completo de limpeza
# --------------------------------------------------------------------------
def full_clean(text: str, base: str = "") -> str:
    original = text

    # 3.1) Tratamento OTRS vs Citsmart
    if base.strip().upper() == "OTRS":
        extracted = process_otrs_header(text)
        if extracted is not None:
            text = extracted
        else:
            # só remove linhas de cabeçalho padrão, preserva resto
            text = re.sub(
                r"^(?:Usuário|IP|Data\s*/?\s*Hora|Comarca|Setor|Função|"
                r"Telefone/Ramal|Área\s*\(SIMP\)|Unidade|Assunto|"
                r"T[ií]tulo do Pedido|Descri[cç][aã]o do pedido)\s*[:\-].*?$",
                "",
                text,
                flags=re.IGNORECASE | re.MULTILINE
            )
    else:
        # bases genéricas: apenas limpa cabeçalhos simples
        text = re.sub(
            r"^(?:Usuário|IP|Data\s*/?\s*Hora|Comarca|Setor|Função|"
            r"Telefone/Ramal|Área\s*\(SIMP\)|Unidade|Assunto)\s*[:\-].*?$",
            "",
            text,
            flags=re.IGNORECASE | re.MULTILINE
        )

    # 3.2) Remover pref/suf de cortesia
    text = remove_generic_patterns(text)

    # 3.3) Minúsculas & trim
    text = text.lower().strip()

    # 3.4) Acentos & caracteres extras
    text = clean_accents(text)
    text = re.sub(r"[^\w\s\.,%/\-]", " ", text)

    # 3.5) Termos técnicos
    text = conserve_tech_terms(text)

    # 3.6) Espaço entre dígitos/letras
    text = re.sub(r"(?<=\d)(?=[^\d\s])|(?<=[^\d\s])(?=\d)", " ", text)

    # 3.7) Lematização + stopwords
    cleaned = lemmatize_text(text)

    # 3.8) Normalize espaços
    cleaned = re.sub(r"\s+", " ", cleaned).strip()

    # 3.9) Fallback final se ficar vazio
    if not cleaned:
        cleaned = lemmatize_text(clean_accents(original))

    logger.debug(f"Texto processado: {cleaned[:150]}...")
    return cleaned

# --------------------------------------------------------------------------
# 4) Logs e métricas
# --------------------------------------------------------------------------
def log_classification_details(y_true, y_pred, labels):
    # Forçamos o Pylance a entender que o resultado é um dicionário
    report = cast(Dict[str, Any], classification_report(y_true, y_pred, target_names=labels, output_dict=True))
    
    logger.info("\nMétricas por Classe:")
    for cls in labels:
        # Verifica se a classe existe no report para evitar KeyErrors
        if cls in report:
            metrics = report[cls]
            logger.info(
                f"{cls}: Precision={metrics['precision']:.2f}  "
                f"Recall={metrics['recall']:.2f}  F1={metrics['f1-score']:.2f}"
            )
            
    logger.info(
        f"Acurácia Geral: {report['accuracy']:.2f}    "
        f"Macro F1: {report['macro avg']['f1-score']:.2f}    "
        f"Weighted F1: {report['weighted avg']['f1-score']:.2f}"
    )

# --------------------------------------------------------------------------
# 5) Treino & tuning com GridSearchCV
# --------------------------------------------------------------------------
def train_and_tune_model(train_df: pd.DataFrame):
    if "Base" not in train_df.columns:
        train_df["Base"] = ""
    X = train_df.apply(lambda r: full_clean(r["Descrição"], r["Base"]), axis=1)
    y = train_df["TAG"].astype(str)

    # Pipeline com placeholder de clf
    pipe = Pipeline([
        ("tfidf", TfidfVectorizer()),
        ("clf", MultinomialNB())
    ])

    # Grid de parâmetros
    param_grid = [
        {
            "tfidf__ngram_range": [(1,1), (1,2)],
            "tfidf__max_df": [0.8, 1.0],
            "tfidf__min_df": [1, 3],
            "tfidf__max_features": [3000, None],
            "clf": [MultinomialNB()],
            "clf__alpha": [0.01, 0.1]
        },
        {
            "tfidf__ngram_range": [(1,1), (1,2)],
            "tfidf__max_df": [0.8, 1.0],
            "tfidf__min_df": [1, 3],
            "tfidf__max_features": [3000, None],
            "clf": [ComplementNB()],
            "clf__alpha": [0.01, 0.1]
        },
        {
            "tfidf__ngram_range": [(1,1), (1,2), (1,4), (2,2)],
            "tfidf__max_df": [0.7, 0.8, 0.9, 1.0],
            "tfidf__min_df": [1, 2, 3, 4, 5],
            "tfidf__max_features": [3000, None],
            "clf": [LinearSVC()],
            "clf__C": [0.01, 0.1, 1, 10]
        },
        {
            "tfidf__ngram_range": [(1,1), (1,2)],
            "tfidf__max_df": [0.8, 1.0],
            "tfidf__min_df": [1, 3],
            "tfidf__max_features": [3000, None],
            "clf": [RandomForestClassifier(random_state=42)],
            "clf__n_estimators": [100, 200],
            "clf__max_depth": [None, 10]
        }
    ]

    grid = GridSearchCV(
        pipe, param_grid, scoring="f1_weighted",
        cv=StratifiedKFold(3, shuffle=True, random_state=42),
        n_jobs=-1, verbose=2
    )
    # .tolist() transforma a Série do Pandas em uma lista padrão do Python
    # O Pylance aceita listas perfeitamente.
    grid.fit(X.tolist(), y.tolist())

    logger.info(f"Melhores parâmetros: {grid.best_params_}")
    logger.info(f"Melhor F1_weighted (CV): {grid.best_score_:.4f}")

    best_pipe = grid.best_estimator_
    joblib.dump(best_pipe, MODEL_PATH)

    # relatório no conjunto de treino
    y_pred = best_pipe.predict(X)
    log_classification_details(y, y_pred, best_pipe.named_steps["clf"].classes_)
    return best_pipe

# --------------------------------------------------------------------------
# 6) Checar necessidade de retrain
# --------------------------------------------------------------------------
def needs_retrain(treino_path: Path, model_path: Path) -> bool:
    return (not model_path.exists()) or (
        treino_path.stat().st_mtime > model_path.stat().st_mtime
    )


# --------------------------------------------------------------------------
# 7) Predição
# --------------------------------------------------------------------------
def predict_tag(text: str, pipeline: Pipeline, base: str) -> str:
    clean = full_clean(text, base)
    # Convertemos explicitamente para string para satisfazer o tipo de retorno
    return str(pipeline.predict([clean])[0])


# --------------------------------------------------------------------------
# 8) Excel formatting: table + conditional formatting + autofit + col widths + wrap
# --------------------------------------------------------------------------
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

# Mapeamento de colunas (1-based) para largura inicial
COLUMN_WIDTHS = {
    1: 15,   # Chamado#
    2: 25,   # Nome do Usuário
    3: 20,   # Data Criação
    4: 30,   # TAG
    5: 25,   # Cidade - Prédio
    6: 25,   # Unidade
    7: 12,   # Ramal
    8: 20,  # Andamento
    9: 100,   # Descrição
    10: 15,  # Base
}

def format_excel(workbook_or_path, fechar_apos: bool = True):
    """
    Se `workbook_or_path` for um tuple (excel_app, wb), reaproveita-os para formatar.
    Se for um Path, abre ou reaproveita do disco.
    Parâmetros:
      - workbook_or_path: ou (excel_app, wb_master) ou Path para o arquivo .xlsx.
      - fechar_apos: se True, fecha o Workbook e dá .Quit() no Excel; se False, mantém aberto.
    """

    # Só funciona no Windows + pywin32
    if platform.system() != 'Windows' or win32 is None:
        return

    # Caso o usuário já me passe (excel_app, wb) diretamente:
    if isinstance(workbook_or_path, tuple):
        excel_app, wb = workbook_or_path
        obtained_from_open = True

    else:
        # É um Path → abre ou reaproveita via GetActiveObject
        master_path = Path(workbook_or_path).resolve()
        try:
            excel_app = win32.GetActiveObject("Excel.Application")
            wb = excel_app.Workbooks(master_path.name)
            obtained_from_open = True
        except Exception:
            excel_app = win32.DispatchEx("Excel.Application")
            excel_app.Visible = False
            excel_app.DisplayAlerts = False
            wb = excel_app.Workbooks.Open(str(master_path))
            obtained_from_open = False

    # Abas que queremos formatar; podem não existir
    aba_nomes = ("Tagged", "Fechados")
    for sheet_name in aba_nomes:
        try:
            ws = wb.Sheets(sheet_name)
        except Exception:
            continue  # se a aba não existir, pula

        used = ws.UsedRange

        # ---- Remover ListObject pré-existente, se houver ----
        for lo in ws.ListObjects:
            if lo.Name == "TabelaChamados":
                lo.Delete()
        # ---- Criar nova tabela em todo o UsedRange ----
        tbl = ws.ListObjects.Add(
            SourceType=1,             # xlSrcRange
            Source=used,
            XlListObjectHasHeaders=1  # xlYes
        )
        tbl.Name = "TabelaChamados"

        # ---- Ajustar larguras + wrap nas colunas ----
        for col_idx, width in COLUMN_WIDTHS.items():
            col = ws.Columns(col_idx)
            col.ColumnWidth = width
            col.WrapText = True

        # ---- Formatação condicional por TAG, se existir a coluna "TAG" ----
        tag_col = None
        for header_cell in used.Rows(1).Cells:
            if header_cell.Value == "TAG":
                tag_col = header_cell.Column
                break

        if tag_col:
            last_row = used.Rows.Count + used.Row - 1
            rng = ws.Range(f"A2:Z{last_row}")  # ajuste se tiver mais de coluna Z

            XL_EXPRESSION = 2  # xlExpression
            XL_AND        = 1  # xlAnd
            col_letter = chr(64 + tag_col)
            for tag, hex_color in TAG_COLORS.items():
                c = hex_color.lstrip("#")
                color_int = int(c[4:6] + c[2:4] + c[0:2], 16)
                formula = f"=${col_letter}2=\"{tag}\""
                cond = rng.FormatConditions.Add(XL_EXPRESSION, XL_AND, formula)
                cond.Interior.Color = color_int

        # ---- AutoFit das linhas (mantém col widths) ----
        used.Rows.AutoFit()

    # ---- Salvar e, se necessário, fechar/quitar ----
    wb.Save()
    if fechar_apos and not obtained_from_open:
        wb.Close(SaveChanges=True)
        excel_app.Quit()

    # Se obtained_from_open=True e fechar_apos=False, deixamos o Excel aberto.
    return None

# --------------------------------------------------------------------------
# 9) Sincronização com o “master” (planilha aberta no OneDrive/SharePoint)
# --------------------------------------------------------------------------

# Função para ler dados mantendo formatação original
def get_sheet_as_dataframe(wb, sheet_name):
    try:
        ws = wb.Sheets(sheet_name)
        used_range = ws.UsedRange
        
        # Extrai cabeçalhos
        headers = [used_range.Cells(1, col).Value 
                    for col in range(1, used_range.Columns.Count + 1)]
        
        # Extrai dados
        data = []
        for row in range(2, used_range.Rows.Count + 1):
            row_data = []
            for col in range(1, used_range.Columns.Count + 1):
                cell = used_range.Cells(row, col)
                row_data.append(cell.Value if cell.Value is not None else "")
            data.append(row_data)
        
        return pd.DataFrame(data, columns=headers)
        
    except Exception:
        return pd.DataFrame()

# Normalização de IDs
def normalize_id(id_val):
    try:
        return str(int(float(id_val))).strip()
    except:
        return str(id_val).strip()

# Backup completo de dados sensíveis
def backup_sensitive_data(df, sheet_name, sensitive_data, sensitive_cols):
    if not df.empty and "Chamado#" in df.columns:
        for _, row in df.iterrows():
            chamado_id = normalize_id(row["Chamado#"])  # <-- normaliza aqui!
            if chamado_id:
                sensitive_data.setdefault(sheet_name, {})[chamado_id] = {
                    col: row.get(col, "") for col in sensitive_cols
                }

# Escreve dados atualizados SEM apagar formatação
def safe_write_to_sheet(wb, sheet_name, df):
    try:
        ws = wb.Sheets(sheet_name)
        ws.UsedRange.ClearContents()
        df = df.fillna("").astype(str)

        n_rows, n_cols = df.shape
        # Cabeçalhos
        headers = [str(col) for col in df.columns]
        ws.Range(ws.Cells(1, 1), ws.Cells(1, n_cols)).Value = headers

        # Dados (como tuplas)
        if n_rows > 0:
            data = [tuple(row) for row in df.values]
            ws.Range(ws.Cells(2, 1), ws.Cells(n_rows + 1, n_cols)).Value = data
    except Exception as e:
        logger.error(f"Erro ao escrever em {sheet_name}: {str(e)}")

def read_excel_com_to_df(ws) -> pd.DataFrame:
    """Lê os dados de uma aba do Excel via COM e converte para DataFrame Pandas."""
    used_range = ws.UsedRange.Value
    
    # Se a planilha estiver vazia, retorna um DataFrame vazio
    if not used_range:
        return pd.DataFrame()
    
    # O Excel via COM retorna uma tupla de tuplas. 
    # A primeira linha (índice 0) é o cabeçalho, o resto (1 em diante) são os dados.
    df = pd.DataFrame(used_range[1:], columns=used_range[0])
    
    # Remove linhas que vieram totalmente em branco do UsedRange do Excel
    df.dropna(how='all', inplace=True)
    
    return df

def sync_to_master(novo_excel_path: Path, master_excel_path: Path) -> Tuple[Any, Any, bool]:
    
    # 1. Função para limpar o ID do chamado (Preserva as suas alterações manuais!)
    def clean_ticket_id(series):
        return series.astype(str).str.strip().str.replace(r'\.0$', '', regex=True)

    # Lê a planilha nova (Que agora já vem com a data perfeitamente formatada do pré-processamento)
    df_tagged_novo = pd.read_excel(novo_excel_path, sheet_name=0)
    df_tagged_novo['Chamado#'] = clean_ticket_id(df_tagged_novo['Chamado#'])

    # Pega carona no Excel aberto na tela
    excel = win32.Dispatch("Excel.Application") # type: ignore
    excel.DisplayAlerts = False

    wb_master = None
    for wb in excel.Workbooks:
        if wb.Name.lower() == master_excel_path.name.lower():
            wb_master = wb
            break

    if wb_master is None:
        wb_master = excel.Workbooks.Open(str(master_excel_path.resolve()))
        
    ws_tagged = wb_master.Sheets("Tagged")

    # Lê os dados do master e limpa os IDs
    df_master_tagged = read_excel_com_to_df(ws_tagged)
    if 'Chamado#' in df_master_tagged.columns:
        df_master_tagged['Chamado#'] = clean_ticket_id(df_master_tagged['Chamado#'])
    else:
        df_master_tagged['Chamado#'] = ""

    chamados_novos = set(df_tagged_novo['Chamado#'].unique())
    chamados_master = set(df_master_tagged['Chamado#'].unique())

    # =====================================================================
    # FEEDBACK LOOP: Salva os chamados fechados no dataset de treino
    # =====================================================================
    fechados_ids = chamados_master - chamados_novos
    if fechados_ids:
        logger.info(f"Chamados fechados identificados: {len(fechados_ids)}. Salvando no dataset de treino...")
        df_fechados = df_master_tagged[df_master_tagged['Chamado#'].isin(fechados_ids)].copy()
        
        from config import TREINO_PATH
        try:
            if TREINO_PATH.exists():
                df_treino_atual = pd.read_excel(TREINO_PATH)
                df_treino_novo = pd.concat([df_treino_atual, df_fechados], ignore_index=True)
            else:
                df_treino_novo = df_fechados
                
            df_treino_novo = df_treino_novo.drop_duplicates(subset=['Chamado#'], keep='last')
            df_treino_novo.to_excel(TREINO_PATH, index=False)
            logger.info("Chamados fechados adicionados ao Chamados_Treino.xlsx com sucesso.")
        except Exception as e:
            logger.error(f"Erro ao salvar chamados fechados no treino: {e}")

    # =====================================================================
    # MANUTENÇÃO DOS DADOS ABERTOS
    # =====================================================================
    # Mantém os antigos que ainda estão abertos (Salva seu trabalho manual)
    df_master_tagged_novo = df_master_tagged[df_master_tagged['Chamado#'].isin(chamados_novos)].copy()

    # Adiciona apenas os inéditos
    chamados_adicionar = chamados_novos - chamados_master
    if chamados_adicionar:
        logger.info(f"Novos chamados identificados: {len(chamados_adicionar)}")
        df_add = df_tagged_novo[df_tagged_novo['Chamado#'].isin(chamados_adicionar)].copy()
        
        # Garante a ordem exata das colunas antes de concatenar
        col_order = ['Chamado#', 'Nome do Usuário', 'Data Criação', 'TAG', 'Cidade - Prédio', 'Unidade', 'Ramal', 'Andamento', 'Descrição', 'Base']
        for c in col_order:
            if c not in df_add.columns:
                df_add[c] = ""
        df_add = df_add[col_order]

        df_master_tagged_novo = pd.concat([df_master_tagged_novo, df_add], ignore_index=True)

    changed = False
    if len(chamados_adicionar) > 0 or (len(chamados_master - chamados_novos) > 0):
        changed = True

    # =====================================================================
    # ESCRITA NA PLANILHA (Mágica Anti-Fantasma)
    # =====================================================================
    if changed:
        # Destrói as Tabelas (ListObjects) transformando em células normais
        for tbl in ws_tagged.ListObjects:
            tbl.Unlist()
            
        # Limpa tudo
        ws_tagged.Cells.Clear()

        if not df_master_tagged_novo.empty:
            # Transforma todos os dados em string limpa (Impede o Excel de estragar as datas que o preprocessamento arrumou)
            df_master_tagged_novo = df_master_tagged_novo.fillna("")
            df_master_tagged_novo = df_master_tagged_novo.astype(str)
            df_master_tagged_novo = df_master_tagged_novo.replace(["nan", "NaT", "<NA>", "None"], "")

            data_to_write_t = [df_master_tagged_novo.columns.tolist()] + df_master_tagged_novo.values.tolist()
            num_rows_t = len(data_to_write_t)
            num_cols_t = len(data_to_write_t[0])
            range_to_write_t = ws_tagged.Range(ws_tagged.Cells(1, 1), ws_tagged.Cells(num_rows_t, num_cols_t))
            range_to_write_t.Value = data_to_write_t

        wb_master.Save()
        logger.info("Planilha master atualizada com sucesso.")

    return excel, wb_master, changed

# --------------------------------------------------------------------------
# 9) Fluxo principal
# --------------------------------------------------------------------------
def main():
    try:

        # --------------------------------------------------
        # (1) treino ou tuning / carregamento do modelo
        # --------------------------------------------------
        if needs_retrain := (not Path(MODEL_PATH).exists() or
                             Path(TREINO_PATH).stat().st_mtime > Path(MODEL_PATH).stat().st_mtime):
            logger.info("Iniciando tuning e treinamento do modelo...")
            df_train = pd.read_excel(TREINO_PATH)
            pipeline = train_and_tune_model(df_train)
        else:
            logger.info("Usando modelo existente")
            pipeline = joblib.load(MODEL_PATH)        

        # --------------------------------------------------
        # (2) lê último “Chamados_Unificados_*.xlsx” como slave
        # --------------------------------------------------        
        files = sorted(OUTPUT_DIR_TRATADOS.glob("Chamados_Unificados_*.xlsx"))
        if not files:
            logger.error("Nenhum arquivo de entrada encontrado")
            return

        df = pd.read_excel(files[-1])

        # preenche manualmente Ramal e Andamento se ausentes
        if "Ramal" not in df.columns: df["Ramal"] = ""
        if "Andamento" not in df.columns: df["Andamento"] = ""

        # gera TAG
        df["TAG"] = df.apply(
            lambda r: predict_tag(r["Descrição"], pipeline, r["Base"]), axis=1
        )

        # reordena e descarta colunas
        final_cols = [
            "Chamado#", "Nome do Usuário", "Data Criação", "TAG",
            "Cidade - Prédio", "Unidade", "Ramal", "Andamento",
            "Descrição", "Base"
        ]
        df = df[final_cols]

        # --------------------------------------------------
        # (3) salva o “slave” temporário (apenas aba Tagged)
        # --------------------------------------------------
        out = OUTPUT_DIR_PRONTO / f"Chamados_Tagged_{datetime.now():%Y-%m-%d_%H-%M-%S}.xlsx"
        with pd.ExcelWriter(out, engine='xlsxwriter') as writer:
            df.to_excel(writer, sheet_name='Tagged', index=False)

        # Format
        format_excel(out, fechar_apos=True)
        logger.info(f"Processo concluído. Arquivo salvo em: {out}")


        # --------------------------------------------------
        # (4) sincroniza com o “master” (que pode estar aberto)
        # --------------------------------------------------
        logger.info("Começando a fazer a sincronização com a planilha master.")

        master = MASTER_FILE_PATH

        # Se o arquivo master não existir (primeira execução), usa o próprio arquivo atual como base
        if not master.exists():
            logger.warning(f"Master não encontrado em {master}. Usando o arquivo atual como base.")
            shutil.copy(out, master)

        # Devolve (excel_app, wb_master) para reutilizarmos no format_excel
        excel_app, wb_master, changed = sync_to_master(out, master)

        logger.info("Sincronização com o master concluída.") 

        # --------------------------------------------------
        # (5) agora reaplica a formatação no “master” atualizado
        # --------------------------------------------------
        # Reaproveitamos a mesma instância COM e o mesmo wb_master
        if changed:
            logger.info("Houve alterações → reaplicando formatação no master.")
            format_excel((excel_app, wb_master), fechar_apos=False)
            logger.info("Formatação do arquivo master concluída.")
        else:
            logger.info("Nenhuma alteração no master → pulei a formatação.")             

    except Exception:
        logger.error("Erro durante a execução", exc_info=True)
        sys.exit(1)

if __name__ == "__main__":
    main()