# Pipeline de Inteligência Artificial e Classificação de TAGs

> **Módulo Principal:** [`src/tag_classifier.py`](../../src/tag_classifier.py)  
> **Testes Unitários:** [`tests/unit/test_ml_champion_pipeline.py`](../../tests/unit/test_ml_champion_pipeline.py)

---

## 🎯 1. Visão Geral do Motor de IA

O Sistema Bancada STI utiliza Machine Learning supervisionado para categorizar automaticamente chamados de TI (atribuição de TAGs técnicas como `IMPRESSORA`, `HARDWARE`, `REDE`, `SISTEMAS`, `SENHAS`, `TELEFONIA`, etc.).

O pipeline foi projetado com arquitetura em dois níveis:
1. **Nível 1 — Treinamento Rápido de Rotina (Modelo Campeão):** Instanciação direta da melhor arquitetura encontrada pelo GridSearchCV. Roda a cada ciclo do orquestrador em ~0.1s a 0.5s, eliminando o antigo gargalo de minutos por ciclo.
2. **Nível 2 — Recalibração Profunda Periódica (GridSearchCV):** Varredura exaustiva de hiperparâmetros e algoritmos concorrentes executada mensalmente pelo `BancadaCronDaemon` ou sob demanda manual pela interface.

---

## 🏎️ 2. Aceleração de Pré-processamento spaCy em Lote

O pré-processamento textual utiliza o modelo `pt_core_news_sm` do spaCy, aplicando as seguintes otimizações de alta performance:

- **Desativação de Componentes Pesados:**
  ```python
  nlp = spacy.load("pt_core_news_sm", disable=["parser", "ner"])
  ```
- **Processamento em Lote (`nlp.pipe`):**
  A função `clean_texts_batch(texts, batch_size=256)` processa centenas de textos simultaneamente no loop interno C do spaCy, em vez de lematizar texto a texto via chamadas individuais.
- **Preservação de Jargões e Entidades Técnicas de TI:**
  Substituições Regex inteligentes protegem numerações de IP (`enderecoip`), componentes de hardware (`ssd`, `memoriaram`, `placadevideo`), suprimentos (`toner`, `fusor`) e códigos de erro antes da remoção de pontuação.

---

## 🧬 3. Enriquecimento de Features (`compose_ticket_text`)

Para evitar predições incorretas causadas por descrições curtas (ex: *"não funciona"* ou *"trocar"*), a função `compose_ticket_text()` unifica múltiplas dimensões do chamado:

```python
def compose_ticket_text(row: dict) -> str:
    titulo = str(row.get("Título") or "").strip()
    unidade = str(row.get("Unidade") or "").strip()
    descricao = str(row.get("Descrição") or "").strip()
    comentarios = str(row.get("Comentarios_Limpos") or "").strip()
    
    parts = []
    if titulo:
        parts.append(titulo)
    if unidade:
        parts.append(unidade)
    if descricao:
        parts.append(descricao)
    if comentarios:
        parts.append(comentarios)
        
    return " ".join(parts)
```

Essa junção fornece ao classificador contexto crucial (por exemplo, a presença da unidade *"Cartório da 1ª Vara"* ou *"Gráfica"* ajuda a desambiguar chamados de impressora ou rede).

---

## 🏆 4. Arquitetura do Modelo Campeão (`train_champion_model`)

Resultados empíricos de sintonia contínua revelaram que o modelo vencedor com maior assertividade (*F1-Score ponderado ~ 0.8423*) é composto por:

```python
from sklearn.pipeline import Pipeline
from sklearn.feature_extraction.text import TfidfVectorizer
from sklearn.svm import LinearSVC

champion_pipeline = Pipeline([
    ('tfidf', TfidfVectorizer(
        ngram_range=(1, 2),
        max_df=0.7,
        min_df=1,
        strip_accents='unicode'
    )),
    ('clf', LinearSVC(
        C=1.0,
        class_weight='balanced',
        random_state=42,
        dual=False
    ))
])
```

### Características Técnicas:
- **`ngram_range=(1, 2)`:** Captura uni-gramas e bi-gramas essenciais (ex: *"tela azul"*, *"cabo de rede"*, *"troca de toner"*).
- **`class_weight='balanced'`:** Compensa desbalanceamento natural de classes (classes com poucos chamados não são ignoradas pelo classificador).
- **`dual=False`:** Preferencial para problemas com mais instâncias do que features, reduzindo o tempo de convergência.

---

## ⚙️ 5. Recalibração Profunda Mensal (`--tune` / Cron Daemon)

O ajuste fino exaustivo roda via `retrain_and_tune_model_job()`:
- Compara 4 famílias de algoritmos: `LinearSVC`, `RandomForestClassifier`, `MultinomialNB` e `ComplementNB`.
- Validação cruzada com `StratifiedKFold(n_splits=5)`.
- Paralelismo com `n_jobs` dinâmico baseado nos núcleos de CPU disponíveis no contêiner.
- Salva o modelo treinado em `models/champion_model.joblib` e métricas de desempenho no banco de dados.

### Como Disparar Manualmente via Terminal:
```bash
docker exec bancada_streamlit_app python src/tag_classifier.py --tune
```

