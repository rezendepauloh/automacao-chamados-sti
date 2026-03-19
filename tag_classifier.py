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

from config import (
    TREINO_PATH, MODEL_PATH, DEBUG_DIR_TAG,
    OUTPUT_DIR_TRATADOS, OUTPUT_DIR_PRONTO
)

# --------------------------------------------------------------------------
# Configuração de Logging
# --------------------------------------------------------------------------
DEBUG_DIR_TAG.mkdir(parents=True, exist_ok=True)

file_handler = RotatingFileHandler(
    filename=DEBUG_DIR_TAG / "tag_classifier.log",
    maxBytes=5 * 1024 * 1024,
    backupCount=3,
    encoding='utf-8'
)
stream_handler = logging.StreamHandler(sys.stdout)

logging.basicConfig(
    level=logging.DEBUG,
    format='[%(asctime)s] [%(levelname)s] %(message)s',
    datefmt='%Y-%m-%d %H:%M:%S',
    handlers=[file_handler, stream_handler]
)
logger = logging.getLogger(__name__)

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

def needs_retrain(treino_path: Path, model_path: Path) -> bool:
    """Verifica se a base de treino é mais recente que o modelo salvo."""
    return (not model_path.exists()) or (
        treino_path.stat().st_mtime > model_path.stat().st_mtime
    )

# --------------------------------------------------------------------------
# Pipeline Principal
# --------------------------------------------------------------------------
def main():
    logger.info("=== INICIANDO CLASSIFICADOR DE TAGS (NLP/ML) ===")
    
    # 1. Busca arquivo unificado mais recente
    arquivos = list(OUTPUT_DIR_TRATADOS.glob("Chamados_Unificados_*.xlsx"))
    if not arquivos:
        logger.error("Nenhum arquivo Unificado encontrado.")
        sys.exit(1)
    
    recente = max(arquivos, key=lambda p: p.stat().st_mtime)
    logger.info(f"Lendo base para classificação: {recente.name}")
    df_unificado = pd.read_excel(recente)
    
    # Prepara descrições
    df_unificado['Descrição_Limpa'] = df_unificado['Descrição'].apply(clean_text)

    # 2. Carrega ou Treina o Modelo
    if not TREINO_PATH.exists():
        logger.error(f"Arquivo de treino não encontrado em {TREINO_PATH}.")
        sys.exit(1)

    df_train = pd.read_excel(TREINO_PATH)
    df_train = df_train.dropna(subset=['TAG', 'Descrição'])
    df_train['Descrição_Limpa'] = df_train['Descrição'].apply(clean_text)

    # Se já existir modelo salvo, reutiliza. Senão, treina um novo.
    if needs_retrain(TREINO_PATH, MODEL_PATH):
        logger.info("Base de treino atualizada ou modelo inexistente. Iniciando retreinamento...")
        pipeline = train_and_tune_model(df_train)
        joblib.dump(pipeline, MODEL_PATH)
        logger.info(f"Novo modelo salvo em: {MODEL_PATH}")
    else:
        logger.info("Usando modelo IA existente (Nenhuma alteração na base de treino).")
        pipeline = joblib.load(MODEL_PATH)       

    # 3. Classifica
    df_tagged = predict_tags(df_unificado, pipeline)
    df_tagged.drop(columns=['Descrição_Limpa'], inplace=True, errors='ignore')

    # 4. Salva a saída final do classificador
    ts = datetime.now().strftime("%Y-%m-%d_%H-%M-%S")
    out = OUTPUT_DIR_PRONTO / f"Chamados_Tagged_{ts}.xlsx"
    df_tagged.to_excel(out, index=False)
    
    logger.info(f"Classificação concluída. Salvo em: {out.name}")
    logger.info("=== FIM DA CLASSIFICAÇÃO ===")

if __name__ == "__main__":
    main()