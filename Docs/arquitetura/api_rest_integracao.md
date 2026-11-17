# 🔌 API REST de Integração (`src/api/`)

A **API REST da Bancada STI** foi projetada para permitir comunicação aberta, padronizada e segura entre o ecossistema conteinerizado (WSL / RedHat) e clientes externos como o **AmbientesHub** (`C:\AmbientesVirtuais`).

---

## 🏗️ 1. Arquitetura e Isolamento

- **Container dedicado (`bancada_api_server`):** Executa em processo e container separados do `bancada_streamlit_app`, garantindo que falhas ou picos de carga no Streamlit não derrubem a API e vice-versa.
- **Stack:** Starlette (ASGI) + Uvicorn + Pydantic v2.
- **Hardening de Banco:** O PostgreSQL opera com `max_connections=100`, `statement_timeout=15000` (15s), `idle_in_transaction_session_timeout=30000` (30s) e `lock_timeout=5000` (5s).

---

## 🔐 2. Segurança e Variáveis de Ambiente

| Variável | Padrão | Descrição |
| :--- | :--- | :--- |
| `BANCADA_API_PORT` | `8510` | Porta interna/externa do servidor Uvicorn |
| `BANCADA_API_BIND` | `127.0.0.1` | Endereço de bind no host (restrito a localhost por padrão) |
| `BANCADA_API_KEY` | `""` | Chave obrigatória (>= 16 chars) para `/api/v1/*` (**Fail-Closed**) |
| `BANCADA_API_RATE_LIMIT` | `180` | Limite de requisições por minuto por IP |
| `BANCADA_API_DOCS_ENABLED` | `"true"` | Habilita `/openapi.json` e `/docs` (Swagger UI) |
| `BANCADA_API_CORS_ORIGINS` | `""` | Lista separada por vírgula de origens CORS permitidas |

- Todas as rotas `/api/v1/*` exigem `X-API-Key` ou `Authorization: Bearer <token>` validados com `hmac.compare_digest`.
- Se `BANCADA_API_KEY` estiver ausente ou tiver menos de 16 caracteres, a API recusa todas as chamadas protegidas com HTTP `503`.

---

## 📡 3. Principais Endpoints

- **Saúde e OpenAPI:** `GET /health`, `GET /openapi.json`, `GET /docs`, `GET /api/v1/info`
- **Chamados (Leitura + Edição):**
  - `GET /api/v1/chamados` (filtros: `search`, `base`, `status`, `tag`, `localidades`, `limit`, `offset`)
  - `GET /api/v1/chamados/localidades` e `GET /api/v1/chamados/filters`
  - `GET /api/v1/chamados/{chamado_id}`
  - `PATCH` / `PUT /api/v1/chamados/{chamado_id}` (atualiza campos e marca `tag_manual` / `dados_manuais`)
  - `GET` / `POST /api/v1/chamados/{chamado_id}/comentarios` e `POST /api/v1/chamados/comentarios/batch`
- **Cron Daemon:**
  - `GET /api/v1/cron`
  - `POST /api/v1/cron/{task_id}/trigger` (define `run_requested = 1`)
  - `PATCH /api/v1/cron/{task_id}` (ativa/desativa ou ajusta agendamento)
- **Infraestrutura e Catálogos:** `GET /api/v1/oxe`, `GET /api/v1/impressoras`, `GET /api/v1/ramais`, `GET /api/v1/unidades`, `GET /api/v1/sccm`.
