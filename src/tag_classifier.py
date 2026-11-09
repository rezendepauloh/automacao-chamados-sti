import os
os.environ["OMP_NUM_THREADS"] = "1"
os.environ["OPENBLAS_NUM_THREADS"] = "1"
os.environ["MKL_NUM_THREADS"] = "1"
os.environ["VECLIB_MAXIMUM_THREADS"] = "1"
os.environ["NUMEXPR_NUM_THREADS"] = "1"

import sys
import re
import logging
from logging.handlers import RotatingFileHandler
from pathlib import Path
from datetime import datetime
import pandas as pd
import joblib

from sklearn.pipeline import Pipeline
from sklearn.svm import LinearSVC
from sklearn.feature_extraction.text import TfidfVectorizer
from sklearn.model_selection import StratifiedKFold, GridSearchCV
from sklearn.naive_bayes import MultinomialNB, ComplementNB
from sklearn.ensemble import RandomForestClassifier
from sklearn.metrics import classification_report

from typing import cast, Dict, Any

import spacy
from spacy.lang.pt.stop_words import STOP_WORDS

root_dir = Path(__file__).resolve().parent.parent
src_dir = Path(__file__).resolve().parent
if str(root_dir) not in sys.path:
    sys.path.insert(0, str(root_dir))
if str(src_dir) not in sys.path:
    sys.path.insert(0, str(src_dir))

from config import (
    TREINO_PATH, MODEL_PATH, DEBUG_DIR_TAG,
    OUTPUT_DIR_TRATADOS, OUTPUT_DIR_PRONTO,
    setup_logging, cleanup_old_files,
    clean_otrs_description, get_ml_n_jobs
)
from terminal import log, print_header, CYAN, GREEN, RED, YELLOW, WHITE

# --------------------------------------------------------------------------
# Configuração de Logging
# --------------------------------------------------------------------------
logger = setup_logging(DEBUG_DIR_TAG / "tag_classifier.log", __name__)

# --------------------------------------------------------------------------
# Configuração de NLP (spaCy Otimizado: sem parser e sem ner para velocidade 10x)
# --------------------------------------------------------------------------
try:
    nlp = spacy.load("pt_core_news_sm", disable=["parser", "ner"])
except OSError:
    logger.error("Modelo do spacy 'pt_core_news_sm' não encontrado. Instale com: python -m spacy download pt_core_news_sm")
    sys.exit(1)

def preclean_raw_text(text: str) -> str:
    """Pré-limpeza rápida em nível de caracteres e termos técnicos."""
    if not isinstance(text, str) or not text:
        return ""
    
    text = text.lower()
    
    # Remoção de acentos
    accent_map = {
        'á':'a','à':'a','â':'a','ã':'a','ä':'a',
        'é':'e','è':'e','ê':'e','ë':'e',
        'í':'i','ì':'i','î':'i','ï':'i',
        'ó':'o','ò':'o','ô':'o','õ':'o','ö':'o',
        'ú':'u','ù':'u','û':'u','ü':'u',
        'ç':'c','ñ':'n', 'ª':'a', 'º':'o'
    }
    text = ''.join(accent_map.get(c, c) for c in text)
    
    # Preservação de Termos Técnicos de TI
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
        text = re.sub(patt, rep, text)

    # Limpeza de HTML, links e pontuação extra
    text = re.sub(r'<[^>]+>', ' ', text)
    text = re.sub(r'http\S+', ' ', text)
    text = re.sub(r'[^\w\s]', ' ', text)
    text = re.sub(r'\s+', ' ', text).strip()
    return text

def clean_texts_batch(texts: list[str]) -> list[str]:
    """
    Processa uma lista de textos em lote com spaCy (pipe) desativando parser e NER.
    É até 10x mais rápido do que processar linha a linha.
    """
    precleaned = [preclean_raw_text(t) for t in texts]
    cleaned = []
    for doc in nlp.pipe(precleaned, batch_size=256):
        tokens = []
        for token in doc:
            if (token.is_alpha or token.like_num) and token.text not in STOP_WORDS:
                lem = token.text if token.like_num else token.lemma_
                tokens.append(lem)
        cleaned.append(" ".join(tokens))
    return cleaned

def clean_text(text: str) -> str:
    """Função utilitária para limpar um único texto."""
    if not text:
        return ""
    return clean_texts_batch([text])[0]

def compose_ticket_text(row) -> str:
    """
    Combina Título, Unidade/Cidade, Descrição e Comentários para
    fornecer contexto rico à classificação de Machine Learning.
    """
    titulo = str(row.get('Título', '') or '').strip()
    unidade = str(row.get('Unidade', '') or row.get('Cidade - Prédio', '') or '').strip()
    desc_raw = str(row.get('Descrição', '') or '').strip()
    desc_clean = clean_otrs_description(desc_raw)
    comments = extract_comments_text(row.get('Comentários', ''))
    
    parts = []
    if titulo and titulo.lower() not in ['nan', 'none', 'null', 'sem título', '<na>', '']:
        parts.append(titulo)
    if unidade and unidade.lower() not in ['nan', 'none', 'null', 'n/a', '<na>', '']:
        parts.append(unidade)
    if desc_clean:
        parts.append(desc_clean)
    if comments:
        parts.append(comments)
        
    return " ".join(parts)

def extract_comments_text(comments_val) -> str:
    """Extrai texto puro dos comentários em formato JSON string ou lista (filtrando robôs)."""
    from config import clean_otrs_comments
    cleaned_comments = clean_otrs_comments(comments_val)
    texts = []
    for c in cleaned_comments:
        if isinstance(c, dict):
            texts.append(str(c.get('texto', '')))
    return " ".join(texts)

# --------------------------------------------------------------------------
# Funções de Machine Learning
# --------------------------------------------------------------------------
def log_classification_details(y_true, y_pred, labels):
    report = cast(Dict[str, Any], classification_report(y_true, y_pred, target_names=labels, output_dict=True, zero_division=0))
    
    logger.info("\nMétricas por Classe:")
    for cls in labels:
        if cls in report:
            metrics = report[cls]
            logger.info(
                f"{cls}: Precision={metrics['precision']:.2f}  "
                f"Recall={metrics['recall']:.2f}  F1={metrics['f1-score']:.2f}"
            )
            
    logger.info(
        f"Acurácia Geral: {report['accuracy']:.2f}    "
        f"Macro F1: {report.get('macro avg', {}).get('f1-score', 0):.2f}    "
        f"Weighted F1: {report.get('weighted avg', {}).get('f1-score', 0):.2f}"
    )

def train_champion_model(train_df: pd.DataFrame) -> Pipeline:
    """
    Treinamento direto ultrarrápido (<1 segundo) do modelo campeão validado (LinearSVC).
    Utilizado na rotina frequente do orquestrador para evitar rodar GridSearchCV pesado a cada 5 minutos.
    """
    import time
    logger.info("🚀 Treinando modelo campeão (LinearSVC Otimizado)...")
    t0 = time.time()
    
    X = train_df['Descrição_Limpa'].tolist()
    y = train_df["TAG"].astype(str).tolist()

    pipe = Pipeline([
        ("tfidf", TfidfVectorizer(
            ngram_range=(1, 2),
            max_df=0.7,
            min_df=1,
            max_features=None
        )),
        ("clf", LinearSVC(
            class_weight='balanced',
            random_state=42,
            dual=False,
            C=1.0
        ))
    ])

    pipe.fit(X, y)
    elapsed = time.time() - t0
    num_classes = len(set(y))
    logger.info(f"✅ Treinamento campeão concluído em {elapsed:.2f}s com {len(X)} exemplos e {num_classes} classes!")
    
    # Salva métricas de status no banco para a UI
    try:
        from src.database.settings_db import set_setting
        now_str = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        set_setting("IA_LAST_TRAIN_DATE", now_str, False, "ia", "Data do último treino rápido de IA")
        set_setting("IA_TRAIN_EXAMPLES_COUNT", str(len(X)), False, "ia", "Total de exemplos na base de treino")
        set_setting("IA_TAGS_COUNT", str(num_classes), False, "ia", "Total de TAGs no vocabulário de IA")
        set_setting("IA_ACTIVE_ALGORITHM", "LinearSVC (TF-IDF N-Grams 1-2)", False, "ia", "Algoritmo ativo de IA")
    except Exception as e_db:
        logger.debug(f"Aviso ao registrar métricas no banco: {e_db}")

    return pipe

def train_and_tune_model(train_df: pd.DataFrame) -> Pipeline:
    """
    Re-calibração profunda e busca exaustiva de hiperparâmetros (GridSearchCV).
    Projetado para rodar sob demanda ou mensalmente via Cron Job agendado.
    """
    logger.info("🔬 Iniciando re-calibração PROFUNDA do modelo (GridSearchCV multi-algoritmo)...")
    
    X = train_df['Descrição_Limpa']
    y = train_df["TAG"].astype(str)

    # Pipeline base
    pipe = Pipeline([
        ("tfidf", TfidfVectorizer()),
        ("clf", MultinomialNB())
    ])

    # Grade de busca de hiperparâmetros
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
            "clf": [LinearSVC(class_weight='balanced', random_state=42, dual=False)],
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

    n_jobs = get_ml_n_jobs()
    logger.info(f"⚡ Executando GridSearchCV com paralelismo: n_jobs={n_jobs}...")

    grid = GridSearchCV(
        pipe, param_grid, scoring="f1_weighted",
        cv=StratifiedKFold(3, shuffle=True, random_state=42),
        n_jobs=n_jobs, verbose=0
    )
    
    grid.fit(X.tolist(), y.tolist())

    logger.info(f"Melhores parâmetros: {grid.best_params_}")
    logger.info(f"Melhor F1_weighted (CV): {grid.best_score_:.4f}")

    best_pipe = grid.best_estimator_
    y_pred = best_pipe.predict(X.tolist())
    
    log_classification_details(y.tolist(), y_pred, best_pipe.named_steps["clf"].classes_)
    
    # Salva métricas de fine-tuning no banco
    try:
        from src.database.settings_db import set_setting
        now_str = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        set_setting("IA_LAST_TUNE_DATE", now_str, False, "ia", "Data da última re-calibração profunda (GridSearchCV)")
        set_setting("IA_BEST_SCORE", f"{grid.best_score_ * 100:.1f}%", False, "ia", "Melhor F1-Score do GridSearchCV")
        set_setting("IA_BEST_MODEL_PARAMS", str(grid.best_params_), False, "ia", "Melhores hiperparâmetros encontrados")
        set_setting("IA_LAST_TRAIN_DATE", now_str, False, "ia", "Data do último treino de IA")
    except Exception as e_db:
        logger.debug(f"Aviso ao registrar métricas de fine-tuning: {e_db}")

    return best_pipe

def predict_tags(df_novos: pd.DataFrame, pipeline: Pipeline) -> pd.DataFrame:
    logger.info("Prevendo TAGs para chamados não classificados...")
    df_result = df_novos.copy()
    X_novos = df_result['Descrição_Limpa'].tolist()
    y_pred = pipeline.predict(X_novos)
    df_result['TAG'] = list(y_pred)
    return df_result

def normalize_for_extraction(text: str) -> str:
    if not isinstance(text, str):
        return ""
    text = text.lower()
    # Remove acentos
    accent_map = {
        'á':'a','à':'a','â':'a','ã':'a','ä':'a',
        'é':'e','è':'e','ê':'e','ë':'e',
        'í':'i','ì':'i','î':'i','ï':'i',
        'ó':'o','ò':'o','ô':'o','õ':'o','ö':'o',
        'ú':'u','ù':'u','û':'u','ü':'u',
        'ç':'c','ñ':'n', 'ª':'a', 'º':'o'
    }
    text = ''.join(accent_map.get(c, c) for c in text)
    return text

def detect_and_update_remote_locations(df: pd.DataFrame) -> pd.DataFrame:
    """
    Analisa a localidade física do chamado com base no IP (SCCM/OTRS) 
    ou por NLP na descrição (trabalho remoto).
    Cria a coluna 'Localidade física' combinada.
    """
    logger.info("Analisando localidade física (IP + NLP)...")
    df_result = df.copy()
    
    # Prédios conhecidos de Campo Grande e seus padrões de busca em regex (normalizados)
    predios_cg_patterns = {
        "Campo Grande - Chácara Cachoeira": [r"chacara\s+cachoeira", r"\bchacara\b"],
        "Campo Grande - Ricardo Brandão": [r"ricardo\s+brandao"],
        "Campo Grande - Rua da Paz": [r"rua\s+da\s+paz", r"\bda\s+paz\b"],
        "Campo Grande - PGJ": [r"\bpgj\b", r"parque\s+dos\s+poderes", r"procuradoria\s+geral"],
        "Campo Grande - Casa da Mulher Brasileira": [r"casa\s+da\s+mulher", r"mulher\s+brasileira"],
        "Campo Grande - GAECO": [r"\bgaeco\b"],
        "Campo Grande - DMP": [r"\bdmp\b"]
    }
    
    context_patterns = [
        r"\bteletrabalho\b", r"trabalho\s+remoto", r"apoio\s+remoto",
        r"suporte\s+remoto", r"trabalhando\b", r"presencialmente\b",
        r"fisicamente\b", r"sala\s+(?:do|de)\s+apoio", r"sala\s+(?:do|de)\s+suporte",
        r"desempenhar\s+(?:minhas\s+)?funcoes", r"\bestou\s+(?:na|no|em|trabalhando)\b",
        r"\bunidade\s+(?:de\s+|da\s+)?", r"\bpredio\s+(?:de\s+|da\s+)?",
        r"\blotado\s+(?:na|no|em)\b", r"\blotada\s+(?:na|no|em)\b"
    ]
    
    updated_count = 0
    try:
        from src.manual_entries import get_location_by_ip
    except ImportError:
        try:
            from manual_entries import get_location_by_ip
        except ImportError:
            try:
                from src.manual_entries_TEMPLATE import get_location_by_ip
            except ImportError:
                def get_location_by_ip(ip: str) -> str:
                    return ""
    
    for idx, row in df_result.iterrows():
        desc = row.get("Descrição", "")
        ip_origem = row.get("IP_Origem", "")
        
        matched_predio = ""
        
        # 1. Tenta por IP
        if ip_origem:
            matched_predio = get_location_by_ip(ip_origem)
            if matched_predio:
                logger.info(f"📍 [IP MATCH] Chamado {row.get('Chamado#')} -> {matched_predio} (IP: {ip_origem})")
                
        # 2. Se não achou por IP, tenta por NLP na descrição
        if not matched_predio and desc and not pd.isna(desc):
            desc_norm = normalize_for_extraction(desc)
            
            for predio_oficial, patterns in predios_cg_patterns.items():
                for pattern in patterns:
                    match = re.search(pattern, desc_norm)
                    if match:
                        match_pos = match.start()
                        start_ctx = max(0, match_pos - 100)
                        end_ctx = min(len(desc_norm), match.end() + 100)
                        context_chunk = desc_norm[start_ctx:end_ctx]
                        
                        has_context = any(re.search(pat, context_chunk) for pat in context_patterns)
                        if has_context:
                            matched_predio = predio_oficial
                            
                            # Detecta sufixo II / 2
                            after_match = desc_norm[match.end():match.end() + 25]
                            suffix_ii_pattern = r"\b(ii|2|unidade\s+ii|unidade\s+2)\b"
                            if re.search(suffix_ii_pattern, after_match):
                                matched_predio += " II"
                            break
                if matched_predio:
                    break
                    
            if matched_predio:
                logger.info(f"✨ [NLP MATCH] Chamado {row.get('Chamado#')} -> {matched_predio} via descrição")
                
        # 3. Define a Localidade Física final
        if matched_predio:
            df_result.at[idx, "Localidade física"] = matched_predio
            updated_count += 1
        else:
            # Fallback: Define a localidade física com base na Cidade - Prédio (sem concatenar unidades internas)
            def _clean_loc(val):
                if val is None or pd.isna(val):
                    return ""
                s = str(val).strip()
                s = re.sub(r'\s*-\s*Sede\b', '', s, flags=re.IGNORECASE).strip()
                low = s.lower()
                if low in ["nan", "none", "null", "<na>", "n/d", ""]:
                    return ""
                if "não encontrad" in low or "nao encontrad" in low:
                    return ""
                return s

            cidade_predio = _clean_loc(row.get("Cidade - Prédio", ""))
            unidade = _clean_loc(row.get("Unidade", ""))
            
            # Para cidades do interior e prédios, a localidade física deve ser a Cidade / Prédio (ex: 'Cassilândia', 'Terenos', 'Campo Grande - PGJ')
            # sem embutir a unidade interna (ex: '1ª PJ de Terenos' fica apenas na coluna Unidade)
            if cidade_predio:
                df_result.at[idx, "Localidade física"] = cidade_predio
            elif unidade:
                # Se não tem prédio cadastrado mas tem unidade, extrai o nome da cidade se for padrão "Xª PJ de Cidade"
                m_pj = re.search(r'\bde\s+([A-Za-zÀ-ÿ\s]+)$', unidade, flags=re.IGNORECASE)
                if m_pj:
                    df_result.at[idx, "Localidade física"] = m_pj.group(1).strip()
                else:
                    df_result.at[idx, "Localidade física"] = unidade
            else:
                df_result.at[idx, "Localidade física"] = "Não identificada"

                
    logger.info(f"Análise de localidades concluída. {updated_count} chamados direcionados por IP/NLP.")
    return df_result

def generate_synthetic_title(tag: str, description: str) -> str:
    """Gera um título descritivo inteligente para chamados que não possuem título nativo (ex: CitSmart)."""
    tag_clean = str(tag).strip() if tag and str(tag).lower() not in ['nan', 'none', 'null'] else ''
    if not description or str(description).lower() in ['nan', 'none', 'null', '']:
        return f'[{tag_clean}] Chamado CitSmart' if tag_clean else 'Chamado CitSmart'

    t = str(description)
    t = re.sub(r'<[^>]+>', ' ', t)
    t = re.sub(r'&\w+;', ' ', t)

    # Remove saudações iniciais repetidas (ex: "Bom dia Prezados!", "Olá, bom dia")
    greeting_pattern = re.compile(
        r'^\s*(?:prezados?|prezadas?|caros?|caras?|olá|ola|bom\s+dia|boa\s+tarde|boa\s+noite|prezada\s+equipe|prezada\s+sti|tudo\s+bem\??|espero\s+que\s+esteja\s+tudo\s+bem\??|espero\s+que\s+sim\??)\b[,\.\-!\s]*',
        flags=re.IGNORECASE
    )
    for _ in range(3):
        t = greeting_pattern.sub('', t).strip()
    # Remove fórmulas burocráticas e de solicitação
    t = re.sub(r'\bpor\s+determina[çc][ãa]o\s+d[eao]\s+[^,]+,\s*', '', t, flags=re.IGNORECASE)
    t = re.sub(r'\bpor\s+meio\s+deste\b[,\s]*', '', t, flags=re.IGNORECASE)
    t = re.sub(
        r'\b(?:solicito\s+provid[êe]ncias\s+para\s+|gostaria\s+de\s+solicitar\s+(?:a|o|que|os|as)?\s*|solicito\s+(?:a|o|que|os|as)?\s*|solicitamos\s+(?:a|o|que|os|as)?\s*|gostaria\s+que\s+verificasse\s+|venho\s+(?:a\s+)?solicitar\s+|preciso\s+(?:de\s+|que\s+)?)\b',
        '', t, flags=re.IGNORECASE
    )
    t = re.sub(r'\b(?:informo\s+que\s+|venho\s+informar\s+que\s+|comunico\s+que\s+)\b', '', t, flags=re.IGNORECASE)
    t = re.sub(r'\s+', ' ', t).strip()

    # Divide na primeira quebra ou oração explicativa
    first_clause = re.split(r'[\.\n\r;]|,\s*(?:pois|a\s+fim|para|conforme|atenciosamente|grata|grato|att|desde\s+já)\b', t, flags=re.IGNORECASE)[0].strip()

    if len(first_clause) < 10:
        first_clause = t[:80].strip()

    first_clause = re.sub(r'^[,\.\-\:\/]+\s*', '', first_clause).strip()
    first_clause = first_clause[0].upper() + first_clause[1:] if first_clause else 'Atendimento técnico'

    if len(first_clause) > 85:
        first_clause = first_clause[:82].rstrip() + '...'

    return f'[{tag_clean}] {first_clause}' if tag_clean else first_clause

def generate_missing_titles(df: pd.DataFrame) -> pd.DataFrame:
    """Preenche a coluna 'Título' para chamados sem título (ex: CitSmart) a partir da TAG e descrição."""
    df_result = df.copy()
    if 'Título' not in df_result.columns:
        df_result['Título'] = ""

    count = 0
    for idx, row in df_result.iterrows():
        curr_title = str(row.get('Título', '')).strip()
        if not curr_title or curr_title.lower() in ['nan', 'none', 'null', 'sem título', '<na>']:
            tag = row.get('TAG', '')
            desc = row.get('Descrição', '')
            syn_title = generate_synthetic_title(tag, desc)
            df_result.at[idx, 'Título'] = syn_title
            count += 1

    if count > 0:
        logger.info(f"✨ [TÍTULOS INTELIGENTES] {count} chamados sem título receberam título gerado via NLP.")
    return df_result

def needs_retrain(treino_path: Path, model_path: Path) -> bool:
    """Verifica se a base de treino é mais recente que o modelo salvo."""
    return (not model_path.exists()) or (
        treino_path.stat().st_mtime > model_path.stat().st_mtime
    )

def prepare_train_dataset(treino_path: Path) -> pd.DataFrame:
    """
    Carrega o dataset de treino com cache persistente da 'Descrição_Limpa'
    para não reprocessar textos antigos via spaCy a cada ciclo.
    """
    if not treino_path.exists():
        logger.error(f"❌ Arquivo de treino não encontrado em {treino_path}.")
        sys.exit(1)

    df_train = pd.read_excel(treino_path)
    df_train = df_train.dropna(subset=['TAG', 'Descrição'])
    if 'Comentários' not in df_train.columns:
        df_train['Comentários'] = '[]'

    # Se a coluna Descrição_Limpa já existe e está populada, processa apenas as linhas faltantes
    needs_save_cache = False
    if 'Descrição_Limpa' not in df_train.columns:
        df_train['Descrição_Limpa'] = ""

    missing_mask = df_train['Descrição_Limpa'].isna() | (df_train['Descrição_Limpa'].astype(str).str.strip() == "")
    missing_count = missing_mask.sum()

    if missing_count > 0:
        logger.info(f"⚡ Limpando {missing_count} textos novos de treino em lote com spaCy...")
        texts_to_clean = [compose_ticket_text(row) for _, row in df_train[missing_mask].iterrows()]
        cleaned_batch = clean_texts_batch(texts_to_clean)
        df_train.loc[missing_mask, 'Descrição_Limpa'] = cleaned_batch
        needs_save_cache = True

    if needs_save_cache:
        try:
            df_train.to_excel(treino_path, index=False)
            logger.info("💾 Cache de 'Descrição_Limpa' atualizado no arquivo de treino.")
        except Exception as e_save:
            logger.warning(f"⚠️ Não foi possível persistir cache de treino no Excel: {e_save}")

    return df_train

def retrain_and_tune_model_job() -> str:
    """
    Job oficial executado pelo Cron Scheduler (mensal ou sob demanda)
    para re-otimizar o classificador com GridSearchCV multi-algoritmo.
    """
    logger.info("🔬 [CRON JOB] Iniciando Re-calibração profunda de IA (GridSearchCV)...")
    df_train = prepare_train_dataset(TREINO_PATH)
    pipeline = train_and_tune_model(df_train)
    joblib.dump(pipeline, MODEL_PATH)
    logger.info(f"✅ Novo modelo campeão salvo em: {MODEL_PATH}")
    return "Re-calibração com GridSearchCV concluída com sucesso e novo modelo gravado."

# --------------------------------------------------------------------------
# Pipeline Principal
# --------------------------------------------------------------------------
def main(force_tune: bool = False):
    print_header("CLASSIFICADOR DE TAGS POR IA (NLP/ML)", color=CYAN)
    logger.info("🧠 Iniciando classificação inteligente de chamados com IA...")
    
    # 1. Busca arquivo unificado mais recente
    arquivos = list(OUTPUT_DIR_TRATADOS.glob("Chamados_Unificados_*.xlsx"))
    if not arquivos:
        logger.error("⚠️ Nenhum arquivo Unificado encontrado.")
        sys.exit(1)
    
    recente = max(arquivos, key=lambda p: p.stat().st_mtime)
    logger.info(f"📂 Lendo base para classificação: {recente.name}")
    df_unificado = pd.read_excel(recente)
    
    if 'Comentários' not in df_unificado.columns:
        df_unificado['Comentários'] = '[]'
        
    # Combina Título, Unidade, Descrição e Comentários e limpa em batch ultrarrápido
    logger.info(f"⚡ Pré-processando {len(df_unificado)} chamados unificados em lote (spaCy batch)...")
    unif_raw_texts = [compose_ticket_text(row) for _, row in df_unificado.iterrows()]
    df_unificado['Descrição_Limpa'] = clean_texts_batch(unif_raw_texts)

    # 2. Carrega ou Treina o Modelo
    df_train = prepare_train_dataset(TREINO_PATH)

    is_tune_requested = force_tune or ("--tune" in sys.argv)
    
    if is_tune_requested:
        logger.info("🔬 Modo de re-calibração profunda solicitado (--tune). Executando GridSearchCV...")
        pipeline = train_and_tune_model(df_train)
        joblib.dump(pipeline, MODEL_PATH)
        logger.info(f"Novo modelo salvo em: {MODEL_PATH}")
    elif needs_retrain(TREINO_PATH, MODEL_PATH):
        logger.info("⚡ Base de treino atualizada ou modelo inexistente. Treinando modelo campeão (LinearSVC)...")
        pipeline = train_champion_model(df_train)
        joblib.dump(pipeline, MODEL_PATH)
        logger.info(f"Novo modelo campeão salvo em: {MODEL_PATH}")
    else:
        logger.info("🤖 Usando modelo IA existente (Nenhuma alteração na base de treino).")
        pipeline = joblib.load(MODEL_PATH)       

    # 3. Classifica
    logger.info(f"🏷️ Classificando {len(df_unificado)} chamados...")
    df_tagged = predict_tags(df_unificado, pipeline)
    df_tagged.drop(columns=['Descrição_Limpa'], inplace=True, errors='ignore')

    df_tagged = detect_and_update_remote_locations(df_tagged)
    df_tagged = generate_missing_titles(df_tagged)

    try:
        from database import save_tickets_to_db, close_missing_tickets_by_base, sync_closed_tickets_to_train_dataset
        logger.info("💾 Salvando dados no banco de dados...")
        save_tickets_to_db(df_tagged)
        
        active_ids_otrs = df_tagged[df_tagged['Base'] == 'OTRS']['Chamado#'].dropna().astype(str).tolist()
        active_ids_citsmart = df_tagged[df_tagged['Base'] == 'CitSmart']['Chamado#'].dropna().astype(str).tolist()
        
        MIN_CONFIDENCE_THRESHOLD = 5
        
        if len(active_ids_otrs) >= MIN_CONFIDENCE_THRESHOLD:
            close_missing_tickets_by_base(active_ids_otrs, 'OTRS')
            logger.info(f"Chamados OTRS antigos fechados. Ativos nesta rodada: {len(active_ids_otrs)}")
        else:
            logger.warning(f"⚠️ Coleta OTRS abaixo do limite de segurança ({len(active_ids_otrs)} < {MIN_CONFIDENCE_THRESHOLD}). Nenhum chamado OTRS foi fechado para evitar falsos positivos.")
            
        if len(active_ids_citsmart) >= MIN_CONFIDENCE_THRESHOLD:
            close_missing_tickets_by_base(active_ids_citsmart, 'CitSmart')
            logger.info(f"Chamados CitSmart antigos fechados. Ativos nesta rodada: {len(active_ids_citsmart)}")
        else:
            logger.warning(f"⚠️ Coleta CitSmart abaixo do limite de segurança ({len(active_ids_citsmart)} < {MIN_CONFIDENCE_THRESHOLD}). Nenhum chamado CitSmart foi fechado para evitar falsos positivos.")
            
        logger.info("Sincronizando chamados fechados com o dataset de treino (Chamados_Treino.xlsx)...")
        sync_closed_tickets_to_train_dataset()
        
        logger.info("Banco de dados e Dataset de Treino atualizados com sucesso.")
    except Exception as db_err:
        logger.error(f"⚠️ Erro ao atualizar banco / dataset de treino: {db_err}")

    ts = datetime.now().strftime("%Y-%m-%d_%H-%M-%S")
    out = OUTPUT_DIR_PRONTO / f"Chamados_Tagged_{ts}.xlsx"
    df_tagged.to_excel(out, index=False)

    cleanup_old_files(OUTPUT_DIR_PRONTO, "Chamados_Tagged_*.xlsx", keep_count=10)
    
    logger.info(f"✅ Classificação por IA finalizada com SUCESSO! Salvo em: {out.name}")
    logger.info("=== FIM DA CLASSIFICAÇÃO ===")

if __name__ == "__main__":
    main()