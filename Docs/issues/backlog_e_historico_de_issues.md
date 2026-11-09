# Backlog, Histórico de Entregas e Roadmap

> **Data de Atualização:** Novembro de 2026  
> **Sistema:** Bancada STI — Ministério Público do Estado de Mato Grosso do Sul (MPMS)  
> 📜 *Para consultar o histórico detalhado dos ciclos de desenvolvimento anteriores (#1 a #12), consulte [historico_sprints_legadas.md](historico_sprints_legadas.md).*

---

## 🏆 1. Histórico de Melhorias e Entregas Recentes

### 🚀 Marco: Modernização de Arquitetura e Engenharia de IA (Nov/2026)
- [x] **Scraper HTTP Rápido do OTRS (`src/scrapers/otrs_scraper_fast.py`):**
  - Substituição da automação pesada do Selenium por sessão HTTP nativa multithreaded (`requests` + `BeautifulSoup`).
  - Redução de tempo de extração de chamados de minutos para menos de 10 segundos (~24x mais veloz).
  - Captura live de notas e comentários para chamados em andamento via paralelismo (`ThreadPoolExecutor`).
- [x] **Aceleração do Pipeline de IA e Modelo Campeão (`src/tag_classifier.py`):**
  - Otimização do spaCy em lote (`clean_texts_batch`) com desativação de componentes lentos.
  - Fixação do modelo campeão `LinearSVC` com `TfidfVectorizer` (F1 ~ 0.8423) para rotinas ordinárias, baixando o tempo de treino para frações de segundo.
  - Implementação de tarefa agendada mensal (`retrain_ia_gridsearch`) no Cron Daemon para recalibração profunda via `GridSearchCV`.
  - Enriquecimento de features combinando Título, Unidade, Descrição e Comentários (`compose_ticket_text`).
- [x] **Transição para Banco de Dados Híbrido PostgreSQL & SQLite:**
  - Migração com sucesso de mais de 16.000 registros para contêiner PostgreSQL 15 (`bancada_postgres_db`).
  - Gerenciador de conexão com pool resiliente e fallback automático em `src/database/connection.py`.
- [x] **Cofre de Segurança Resiliente (`src/crypto_utils.py`):**
  - Proteção contra falhas em trocas de ambiente: tokens incompatíveis não travam a autenticação e caem graciosamente para keyring e variáveis de ambiente.
- [x] **Suíte de Testes Automatizados com 100% de Cobertura (90 testes):**
  - Criação de `tests/unit/test_ml_champion_pipeline.py`.
  - Runner mestre `tests/run_all.py` rodando 100% sem falhas no contêiner Docker.
- [x] **Padronização para Agentes Autônomos (`AGENTS.md`):**
  - Definição do Ciclo de Engenharia Obrigatório de 4 etapas: Avaliar -> Planejar -> Testar -> Documentar.
  - Estruturação completa da documentação no padrão unificado `Docs/`.

---

## 📌 2. Backlog e Próximos Passos (Roadmap)

### Curto Prazo (Operacional & Performance):
- [ ] Implementar exportação direta de termos de doação em PDF vetorial via Playwright no módulo de redistribuição.
- [ ] Adicionar filtros de data relativa dinâmicos ("Últimos 7 dias", "Mês atual") nos gráficos analíticos do dashboard.
- [ ] Criar alerta visual na interface quando houver chamados VIP ou com mais de 48h sem movimentação.

### Médio Prazo (Infraestrutura & Integrações):
- [ ] Integração com webhook do Microsoft Teams para envio de relatórios de plantão de TI nos finais de semana.
- [ ] Automatizar auditoria remota periódica de estações de trabalho via WMI/SCCM para detecção proativa de discos com falhas iminentes.
- [ ] Expansão do catálogo de FAQs técnicos com busca semântica vetorial.

