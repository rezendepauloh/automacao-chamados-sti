# -*- coding: utf-8 -*-
"""
Testes unitários para o modelo campeão de Machine Learning e o pipeline otimizado (src.tag_classifier).
Valida:
1. Estrutura do pipeline campeão (LinearSVC + TF-IDF) e seus hiperparâmetros fixados.
2. Aceleração de pré-processamento spaCy em lote (clean_texts_batch).
3. Enriquecimento de features combinando Título, Unidade, Descrição e Comentários (compose_ticket_text).
4. Treinamento rápido e inferência do modelo campeão.
5. Integração com a rotina de re-calibração agendada (retrain_and_tune_model_job).
"""

import sys
import unittest
from pathlib import Path
import pandas as pd
from sklearn.pipeline import Pipeline
from sklearn.svm import LinearSVC
from sklearn.feature_extraction.text import TfidfVectorizer

ROOT_DIR = Path(__file__).resolve().parent.parent.parent
if str(ROOT_DIR) not in sys.path:
    sys.path.insert(0, str(ROOT_DIR))

import tests.test_helpers
from src.tag_classifier import (
    clean_texts_batch,
    clean_text,
    compose_ticket_text,
    train_champion_model,
    retrain_and_tune_model_job,
    needs_retrain
)
from src.database.cron_db import DEFAULT_TASKS

class TestMLChampionPipeline(unittest.TestCase):

    def test_champion_pipeline_hyperparameters(self):
        """Valida se o modelo campeão utiliza os hiperparâmetros exatos validados no GridSearchCV."""
        # Cria pequeno dataset sintético para instanciar o pipeline
        df_dummy = pd.DataFrame({
            "Descrição_Limpa": [
                "computador nao liga fonte queimada",
                "troca de toner impressora laser",
                "sem acesso ao sistema saaj e-mail institucional"
            ],
            "TAG": ["HARDWARE", "IMPRESSORA", "SISTEMAS"]
        })

        pipe = train_champion_model(df_dummy)
        
        # 1. Verifica se é uma Pipeline do Scikit-Learn
        self.assertIsInstance(pipe, Pipeline)
        self.assertIn("tfidf", pipe.named_steps)
        self.assertIn("clf", pipe.named_steps)

        # 2. Valida hiperparâmetros do TF-IDF
        tfidf = pipe.named_steps["tfidf"]
        self.assertIsInstance(tfidf, TfidfVectorizer)
        self.assertEqual(tfidf.ngram_range, (1, 2))
        self.assertEqual(tfidf.max_df, 0.7)
        self.assertEqual(tfidf.min_df, 1)
        self.assertIsNone(tfidf.max_features)

        # 3. Valida hiperparâmetros do classificador campeão (LinearSVC)
        clf = pipe.named_steps["clf"]
        self.assertIsInstance(clf, LinearSVC)
        self.assertEqual(clf.C, 1.0)
        self.assertEqual(clf.class_weight, "balanced")
        self.assertFalse(clf.dual)
        self.assertEqual(clf.random_state, 42)

    def test_clean_texts_batch_fast_spacy(self):
        """Valida que o processamento em lote do spaCy lematiza e preserva termos de TI."""
        sample_texts = [
            "Computador com problema no SSD e na memória RAM.",
            "Troca de TONER na impressora Samsung 10.111.64.59.",
            "Sem conexão de IP na rede da promotoria."
        ]
        cleaned = clean_texts_batch(sample_texts)

        self.assertEqual(len(cleaned), 3)
        # Termos técnicos devem ser preservados/normalizados
        self.assertTrue("ssd" in cleaned[0] or "memoriaram" in cleaned[0])
        # Acentos devem ser removidos e texto em minúsculas
        self.assertNotIn("memória", cleaned[0])
        self.assertNotIn("conexão", cleaned[2])
        self.assertTrue("enderecoip" in cleaned[2] or "ip" in cleaned[2])

    def test_compose_ticket_text_feature_enrichment(self):
        """Valida que Título e Unidade são incorporados à Descrição para aumentar assertividade."""
        row_full = {
            "Título": "Sem Internet no Cartório",
            "Unidade": "1ª PJ de Três Lagoas",
            "Descrição": "Não conseguimos acessar nada desde cedo.",
            "Comentários": "[]"
        }
        composed = compose_ticket_text(row_full)
        
        self.assertIn("Sem Internet no Cartório", composed)
        self.assertIn("1ª PJ de Três Lagoas", composed)
        self.assertIn("Não conseguimos acessar nada desde cedo.", composed)

    def test_compose_ticket_text_handles_missing_fields(self):
        """Valida que campos nulos ou NaN não quebram a composição de texto."""
        row_partial = {
            "Título": None,
            "Unidade": float('nan'),
            "Descrição": "Microcomputador travando na inicialização.",
            "Comentários": ""
        }
        composed = compose_ticket_text(row_partial)
        self.assertNotIn("nan", composed.lower())
        self.assertNotIn("none", composed.lower())
        self.assertIn("Microcomputador travando na inicialização.", composed)

    def test_champion_model_fit_and_prediction_accuracy(self):
        """Valida o treinamento rápido e a capacidade de predição do modelo campeão."""
        df_train_synthetic = pd.DataFrame({
            "Descrição_Limpa": [
                "troca de toner impressora papel atolado",
                "toner preto impressora laser falhando impressao",
                "computador nao liga fonte memoria tela azul",
                "mouse teclado monitor nao funciona hardware",
                "sem acesso internet rede cabo desconectado",
                "rede wifi lenta conexao caindo switch"
            ],
            "TAG": [
                "IMPRESSORA", "IMPRESSORA",
                "HARDWARE", "HARDWARE",
                "REDE", "REDE"
            ]
        })

        pipe = train_champion_model(df_train_synthetic)
        
        # Teste de predição com novas frases nunca vistas
        test_samples = [
            "preciso de toner novo para impressora",
            "o computador esta com tela azul de memoria",
            "a rede de internet caiu na unidade"
        ]
        preds = pipe.predict(test_samples)
        
        self.assertEqual(preds[0], "IMPRESSORA")
        self.assertEqual(preds[1], "HARDWARE")
        self.assertEqual(preds[2], "REDE")

    def test_cron_task_retrain_ia_gridsearch_is_registered(self):
        """Valida que a rotina mensal de re-calibração profunda está cadastrada nas tarefas padrão."""
        task_ids = [t["task_id"] for t in DEFAULT_TASKS]
        self.assertIn("retrain_ia_gridsearch", task_ids)

        task_data = next(t for t in DEFAULT_TASKS if t["task_id"] == "retrain_ia_gridsearch")
        self.assertEqual(task_data["categoria"], "Inteligência Artificial")
        self.assertEqual(task_data["intervalo_valor"], 30)
        self.assertEqual(task_data["intervalo_unidade"], "dias")

if __name__ == "__main__":
    unittest.main()

