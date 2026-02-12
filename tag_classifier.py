import sys
import re
import logging
import numpy as np
from pathlib import Path
from datetime import datetime
import pandas as pd
import joblib

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
        logging.FileHandler("tag_classifier.log", encoding='utf-8'),
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

def sync_to_master(slave_path: Path, master_path: Path) -> Tuple[Any, Any, bool]:
    """
    Sincroniza o "master" com o "slave" preservando "Ramal" e "Andamento"
    """
    try:
        # Adicione esta verificação antes de chamar o Dispatch
        if win32 is None:
            logger.error("Win32com não está disponível. Não é possível manipular Excel via COM.")
            return None, None, False        
        
        # 1) Conecta ao Excel
        excel_app = win32.Dispatch("Excel.Application")
        excel_app.Visible = True
        workbook_name = master_path.name
        
        try:
            wb = excel_app.Workbooks(workbook_name)
            obtained_from_open = True
        except Exception:
            wb = excel_app.Workbooks.Open(str(master_path))
            obtained_from_open = False

        # 2) Carrega dados preservando colunas sensíveis
        df_master_tagged = get_sheet_as_dataframe(wb, "Tagged")
        df_master_fechados = get_sheet_as_dataframe(wb, "Fechados")
        df_slave_tagged = pd.read_excel(slave_path, sheet_name="Tagged")

        if df_master_tagged.empty:
            logger.error("df_master_tagged está vazio!")
        else:
            logger.debug(f"Colunas disponíveis: {df_master_tagged.columns}")

        #df_master_tagged.to_csv("df_master_tagged_ANTES.csv", index=False, encoding="utf-8-sig")
        #df_master_fechados.to_csv("df_master_fechados_ANTES.csv", index=False, encoding="utf-8-sig")
        #df_slave_tagged.to_csv("df_slave_tagged_ANTES.csv", index=False, encoding="utf-8-sig")

        # 3) Backup completo de dados sensíveis
        sensitive_data = {}
        sensitive_cols = ["Ramal", "Andamento", "Cidade - Prédio", "Unidade", "TAG"]

        backup_sensitive_data(df_master_tagged, "Tagged", sensitive_data, sensitive_cols)
        backup_sensitive_data(df_master_fechados, "Fechados", sensitive_data, sensitive_cols)

        logger.debug(f"sensitive_data (após backup): {sensitive_data}")

        # 4) Normalização de IDs
        df_master_tagged["Chamado#"] = df_master_tagged["Chamado#"].apply(normalize_id)
        df_master_fechados["Chamado#"] = df_master_fechados["Chamado#"].apply(normalize_id)
        df_slave_tagged["Chamado#"] = df_slave_tagged["Chamado#"].apply(normalize_id)

        # 5) Identifica mudanças
        master_ids = set(df_master_tagged["Chamado#"].unique())
        slave_ids = set(df_slave_tagged["Chamado#"].unique())
        
        ids_mantidos = master_ids & slave_ids
        ids_novos = slave_ids - master_ids
        ids_fechados = master_ids - slave_ids

        logger.debug(f"set_master_ids  (master.Tagged): {sorted(master_ids)}")
        logger.debug(f"set_slave_ids   (slave.Tagged) : {sorted(slave_ids)}")
        logger.debug(f"ids_mantidos    (intersection)   : {sorted(ids_mantidos)}")
        logger.debug(f"ids_novos       = {sorted(ids_novos)}")
        logger.debug(f"ids_fechados    = {sorted(ids_fechados)}") 

        # 6) Se não houver alterações, retorna sem mudanças
        if not ids_novos and not ids_fechados:
            return (excel_app, wb, False)
        
        # 7) Atualiza a aba Tagged preservando dados
        tagged_rows = []

        # 7.1) Mantém chamados existentes com dados atualizados        
        for _, slave_row in df_slave_tagged.iterrows():
            chamado_id = slave_row["Chamado#"]
            
            if chamado_id in ids_mantidos:
                if "Tagged" in sensitive_data and chamado_id in sensitive_data["Tagged"]:
                    for col in sensitive_cols:
                        logger.debug(f"Antes: chamado {chamado_id} {col}={slave_row.get(col)}")
                        
                        slave_row[col] = sensitive_data["Tagged"][chamado_id].get(col, slave_row.get(col, ""))
                        
                        logger.debug(f"Depois: chamado {chamado_id} {col}={slave_row.get(col)}")
                
                tagged_rows.append(slave_row)

        logger.debug(f"Restaurando Ramal/Andamento para chamado {chamado_id}: {sensitive_data['Tagged'].get(chamado_id)}")

        # 7.2) Adiciona novos chamados
        for _, slave_row in df_slave_tagged.iterrows():
            chamado_id = slave_row["Chamado#"]
            if chamado_id in ids_novos:
                tagged_rows.append(slave_row)
        
        df_master_tagged_novo = pd.DataFrame(tagged_rows)
        logger.debug(f"df_master_tagged_novo preview:\n{df_master_tagged_novo.head()}")

        # 8) Atualiza a aba Fechados preservando dados
        fechados_rows = df_master_fechados.to_dict('records')
        
        for _, master_row in df_master_tagged.iterrows():
            chamado_id = master_row["Chamado#"]
            
            if chamado_id in ids_fechados:
                # Adiciona dados sensíveis existentes
                if "Tagged" in sensitive_data and chamado_id in sensitive_data["Tagged"]:
                    for col in sensitive_cols:
                        if col in master_row:
                            master_row[col] = sensitive_data["Tagged"][chamado_id][col]
                fechados_rows.append(master_row.to_dict())
        
        df_master_fechados_novo = pd.DataFrame(fechados_rows)
        logger.debug(f"df_master_fechados_novo preview:\n{df_master_fechados_novo.head()}")

        # 9) Mantém ordem original das colunas
        def maintain_column_order(df, reference_df):
            if not reference_df.empty:
                return df[reference_df.columns]
            return df
        
        df_master_tagged_novo = maintain_column_order(df_master_tagged_novo, df_master_tagged)
        df_master_fechados_novo = maintain_column_order(df_master_fechados_novo, df_master_fechados)

        # 10) Escreve dados atualizados SEM apagar formatação
        #df_master_tagged_novo.to_csv("df_master_tagged_DEPOIS.csv", index=False, encoding="utf-8-sig")
        #df_master_fechados_novo.to_csv("df_master_fechados_DEPOIS.csv", index=False, encoding="utf-8-sig")

        logger.debug(f"df_master_tagged_novo shape: {df_master_tagged_novo.shape}, columns: {df_master_tagged_novo.columns}")

        # Desabilite atualização visual e cálculo durante a escrita
        excel_app.ScreenUpdating = False
        excel_app.Calculation = -4135  # xlCalculationManual
        excel_app.EnableEvents = False

        safe_write_to_sheet(wb, "Tagged", df_master_tagged_novo)
        safe_write_to_sheet(wb, "Fechados", df_master_fechados_novo)

        # Alterar atualização visual e cálculo durante a escrita
        excel_app.ScreenUpdating = True
        excel_app.Calculation = -4105  # xlCalculationAutomatic
        excel_app.EnableEvents = True        

        # 11) Salva e retorna
        wb.Save()
        return (excel_app, wb, True)

    except Exception as e:
        logger.error(f"Erro crítico na sincronização: {str(e)}", exc_info=True)
        return (None, None, False)

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

        # df["cleaned_text"] = df.apply(
        #     lambda r: full_clean(r["Descrição"], r["Base"]), axis=1
        # )

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

        # master = Path(
        #     r"C:\Users\paulogoncalves\OneDrive - Ministerio Público do Estado de Mato Grosso do Sul\Documentos SharePoint DIT-Manutenção\Chamados\Chamados_Unificados_Final.xlsx"
        # )
        master = MASTER_FILE_PATH
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