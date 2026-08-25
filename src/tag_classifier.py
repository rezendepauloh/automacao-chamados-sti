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
    clean_otrs_description
)
from terminal import log, print_header, CYAN, GREEN, RED, YELLOW, WHITE

# --------------------------------------------------------------------------
# Configuração de Logging
# --------------------------------------------------------------------------
logger = setup_logging(DEBUG_DIR_TAG / "tag_classifier.log", __name__)

# --------------------------------------------------------------------------
# Configuração de NLP (spaCy)
# --------------------------------------------------------------------------
try:
    nlp = spacy.load("pt_core_news_sm")
except OSError:
    logger.error("Modelo do spacy 'pt_core_news_sm' não encontrado. Instale com: python -m spacy download pt_core_news_sm")
    sys.exit(1)

def clean_text(text: str) -> str:
    """Limpeza de texto (NLP) focada em extração de features para TI."""
    if not isinstance(text, str):
        return ""
    
    # 1. Tudo em minúsculas
    text = text.lower()
    
    # 2. Remoção de acentos (sua lógica)
    accent_map = {
        'á':'a','à':'a','â':'a','ã':'a','ä':'a',
        'é':'e','è':'e','ê':'e','ë':'e',
        'í':'i','ì':'i','î':'i','ï':'i',
        'ó':'o','ò':'o','ô':'o','õ':'o','ö':'o',
        'ú':'u','ù':'u','û':'u','ü':'u',
        'ç':'c','ñ':'n', 'ª':'a', 'º':'o'
    }
    text = ''.join(accent_map.get(c, c) for c in text)
    
    # 3. Preservação de Termos Técnicos de TI (sua lógica)
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

    # 4. Limpeza de HTML, links e pontuação extra
    text = re.sub(r'<[^>]+>', ' ', text)
    text = re.sub(r'http\S+', ' ', text)
    text = re.sub(r'[^\w\s]', ' ', text)
    text = re.sub(r'\s+', ' ', text).strip()
    
    # 5. NLP com spaCy (Lematização + Preservação de números)
    doc = nlp(text)
    tokens = []
    for token in doc:
        # Aceita palavras (is_alpha) ou números (like_num) que não sejam Stop Words
        if (token.is_alpha or token.like_num) and token.text not in STOP_WORDS:
            # Mantém o número literal, e lematiza as palavras
            lem = token.text if token.like_num else token.lemma_
            tokens.append(lem)
            
    return " ".join(tokens)

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

def train_and_tune_model(train_df: pd.DataFrame) -> Pipeline:
    logger.info("Iniciando treinamento PESADO e otimização do modelo (GridSearchCV)...")
    
    X = train_df['Descrição_Limpa']
    y = train_df["TAG"].astype(str)

    # Pipeline com placeholder genérico de classificador
    pipe = Pipeline([
        ("tfidf", TfidfVectorizer()),
        ("clf", MultinomialNB())
    ])

    # A sua super grade de parâmetros (A Arena)
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

    grid = GridSearchCV(
        pipe, param_grid, scoring="f1_weighted",
        cv=StratifiedKFold(3, shuffle=True, random_state=42),
        n_jobs=1, verbose=0 # mantido 0 para não poluir os logs de produção
    )
    
    grid.fit(X.tolist(), y.tolist())

    logger.info(f"Melhores parâmetros: {grid.best_params_}")
    logger.info(f"Melhor F1_weighted (CV): {grid.best_score_:.4f}")

    best_pipe = grid.best_estimator_
    y_pred = best_pipe.predict(X.tolist())
    
    # Chama o seu log detalhado
    log_classification_details(y.tolist(), y_pred, best_pipe.named_steps["clf"].classes_)
    
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
            # Fallback: Concatena Cidade - Prédio e Unidade (removendo " - Sede" desnecessário)
            cidade_predio = str(row.get("Cidade - Prédio", "")).strip()
            cidade_predio = re.sub(r'\s*-\s*Sede\b', '', cidade_predio, flags=re.IGNORECASE).strip()
            
            unidade = str(row.get("Unidade", "")).strip()
            unidade = re.sub(r'\s*-\s*Sede\b', '', unidade, flags=re.IGNORECASE).strip()
            
            if unidade and unidade not in ["N/D", "Não encontrada no AD", "Não encontrada"]:
                if cidade_predio and cidade_predio != unidade:
                    df_result.at[idx, "Localidade física"] = f"{cidade_predio} - {unidade}"
                else:
                    df_result.at[idx, "Localidade física"] = unidade
            else:
                df_result.at[idx, "Localidade física"] = cidade_predio or "Não identificada"

                
    logger.info(f"Análise de localidades concluída. {updated_count} chamados direcionados por IP/NLP.")
    return df_result

def needs_retrain(treino_path: Path, model_path: Path) -> bool:
    """Verifica se a base de treino é mais recente que o modelo salvo."""
    return (not model_path.exists()) or (
        treino_path.stat().st_mtime > model_path.stat().st_mtime
    )

# --------------------------------------------------------------------------
# Pipeline Principal
# --------------------------------------------------------------------------
def main():
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
    df_unificado['Descrição_Limpa'] = (
        df_unificado['Descrição'].apply(clean_otrs_description) + " " + 
        df_unificado['Comentários'].apply(extract_comments_text)
    ).apply(clean_text)

    # 2. Carrega ou Treina o Modelo
    if not TREINO_PATH.exists():
        logger.error(f"❌ Arquivo de treino não encontrado em {TREINO_PATH}.")
        sys.exit(1)

    df_train = pd.read_excel(TREINO_PATH)
    df_train = df_train.dropna(subset=['TAG', 'Descrição'])
    if 'Comentários' not in df_train.columns:
        df_train['Comentários'] = '[]'
    df_train['Descrição_Limpa'] = (
        df_train['Descrição'].apply(clean_otrs_description) + " " + 
        df_train['Comentários'].apply(extract_comments_text)
    ).apply(clean_text)

    if needs_retrain(TREINO_PATH, MODEL_PATH):
        logger.info("⚡ Base de treino atualizada ou modelo inexistente. Iniciando retreinamento...")
        pipeline = train_and_tune_model(df_train)
        joblib.dump(pipeline, MODEL_PATH)
        logger.info(f"Novo modelo salvo em: {MODEL_PATH}")
    else:
        logger.info("🤖 Usando modelo IA existente (Nenhuma alteração na base de treino).")
        pipeline = joblib.load(MODEL_PATH)       

    # 3. Classifica
    logger.info(f"🏷️ Classificando {len(df_unificado)} chamados...")
    df_tagged = predict_tags(df_unificado, pipeline)
    df_tagged.drop(columns=['Descrição_Limpa'], inplace=True, errors='ignore')

    df_tagged = detect_and_update_remote_locations(df_tagged)

    try:
        from database import save_tickets_to_db, close_missing_tickets_by_base, sync_closed_tickets_to_train_dataset
        logger.info("💾 Salvando dados no banco SQLite...")
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
        
        logger.info("Banco SQLite e Dataset de Treino atualizados com sucesso.")
    except Exception as db_err:
        logger.error(f"⚠️ Erro ao atualizar banco SQLite/Treino: {db_err}")

    ts = datetime.now().strftime("%Y-%m-%d_%H-%M-%S")
    out = OUTPUT_DIR_PRONTO / f"Chamados_Tagged_{ts}.xlsx"
    df_export = df_tagged.drop(columns=['Título'], errors='ignore')
    df_export.to_excel(out, index=False)

    cleanup_old_files(OUTPUT_DIR_PRONTO, "Chamados_Tagged_*.xlsx", keep_count=10)
    
    logger.info(f"✅ Classificação por IA finalizada com SUCESSO! Salvo em: {out.name}")
    logger.info("=== FIM DA CLASSIFICAÇÃO ===")

if __name__ == "__main__":
    main()