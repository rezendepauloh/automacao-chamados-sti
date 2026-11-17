# -*- coding: utf-8 -*-
"""
Especificação OpenAPI 3.0.3 e Swagger UI para documentação aberta da API Bancada STI.
"""

from typing import Any, Dict


def build_openapi_spec() -> Dict[str, Any]:
    """Retorna o schema OpenAPI 3.0.3 completo dos endpoints da Bancada STI."""
    sec_ref = [{"ApiKeyAuth": []}, {"BearerAuth": []}]
    return {
        "openapi": "3.0.3",
        "info": {
            "title": "Bancada STI — API de Integração",
            "version": "1.0.0",
            "description": (
                "API REST segura e aberta (JSON/OpenAPI) para leitura e edição de chamados, "
                "controle de rotinas do Cron, Central Telefônica OXE, PaperCut, Ramais, Unidades e SCCM."
            ),
        },
        "components": {
            "securitySchemes": {
                "ApiKeyAuth": {"type": "apiKey", "in": "header", "name": "X-API-Key"},
                "BearerAuth": {"type": "http", "scheme": "bearer"},
            }
        },
        "paths": {
            "/health": {
                "get": {
                    "summary": "Liveness e conectividade com banco",
                    "responses": {"200": {"description": "Serviço operacional"}},
                }
            },
            "/api/v1/info": {
                "get": {
                    "summary": "Metadados e módulos da API",
                    "security": sec_ref,
                    "responses": {"200": {"description": "Metadados retornados"}},
                }
            },
            "/api/v1/chamados": {
                "get": {
                    "summary": "Listar e filtrar chamados (com KPIs)",
                    "security": sec_ref,
                    "parameters": [
                        {"name": "search", "in": "query", "schema": {"type": "string"}},
                        {"name": "base", "in": "query", "schema": {"type": "string"}},
                        {"name": "status", "in": "query", "schema": {"type": "string", "default": "Aberto"}},
                        {"name": "tag", "in": "query", "schema": {"type": "string"}},
                        {"name": "localidades", "in": "query", "schema": {"type": "string"}},
                        {"name": "limit", "in": "query", "schema": {"type": "integer", "default": 150}},
                        {"name": "offset", "in": "query", "schema": {"type": "integer", "default": 0}},
                    ],
                    "responses": {"200": {"description": "Lista paginada de chamados e summary"}},
                }
            },
            "/api/v1/chamados/localidades": {
                "get": {
                    "summary": "Listar localidades físicas distintas",
                    "security": sec_ref,
                    "responses": {"200": {"description": "Lista de localidades"}},
                }
            },
            "/api/v1/chamados/filters": {
                "get": {
                    "summary": "Opções dinâmicas de TAGs e Localidades",
                    "security": sec_ref,
                    "responses": {"200": {"description": "Tags e localidades filtradas"}},
                }
            },
            "/api/v1/chamados/{chamado_id}": {
                "get": {
                    "summary": "Detalhe do chamado e comentários",
                    "security": sec_ref,
                    "parameters": [{"name": "chamado_id", "in": "path", "required": True, "schema": {"type": "string"}}],
                    "responses": {"200": {"description": "Chamado encontrado"}, "404": {"description": "Não encontrado"}},
                },
                "patch": {
                    "summary": "Editar parcialmente campos do chamado",
                    "security": sec_ref,
                    "parameters": [{"name": "chamado_id", "in": "path", "required": True, "schema": {"type": "string"}}],
                    "responses": {"200": {"description": "Chamado atualizado"}, "404": {"description": "Não encontrado"}},
                },
                "put": {
                    "summary": "Atualizar campos do chamado",
                    "security": sec_ref,
                    "parameters": [{"name": "chamado_id", "in": "path", "required": True, "schema": {"type": "string"}}],
                    "responses": {"200": {"description": "Chamado atualizado"}, "404": {"description": "Não encontrado"}},
                },
            },
            "/api/v1/chamados/{chamado_id}/comentarios": {
                "get": {
                    "summary": "Listar comentários de um chamado",
                    "security": sec_ref,
                    "parameters": [{"name": "chamado_id", "in": "path", "required": True, "schema": {"type": "string"}}],
                    "responses": {"200": {"description": "Lista de comentários"}},
                },
                "post": {
                    "summary": "Adicionar comentário técnico em um chamado",
                    "security": sec_ref,
                    "parameters": [{"name": "chamado_id", "in": "path", "required": True, "schema": {"type": "string"}}],
                    "responses": {"201": {"description": "Comentário registrado"}},
                },
            },
            "/api/v1/chamados/comentarios/batch": {
                "post": {
                    "summary": "Buscar comentários em lote por IDs",
                    "security": sec_ref,
                    "responses": {"200": {"description": "Dicionário de comentários por chamado_id"}},
                }
            },
            "/api/v1/cron": {
                "get": {
                    "summary": "Listar agendamentos, logs e métricas do Cron Daemon",
                    "security": sec_ref,
                    "responses": {"200": {"description": "Schedules e logs"}},
                }
            },
            "/api/v1/cron/{task_id}/trigger": {
                "post": {
                    "summary": "Solicitar execução imediata de uma rotina do Cron",
                    "security": sec_ref,
                    "parameters": [{"name": "task_id", "in": "path", "required": True, "schema": {"type": "string"}}],
                    "responses": {"200": {"description": "Execução solicitada"}},
                }
            },
            "/api/v1/cron/{task_id}": {
                "patch": {
                    "summary": "Ativar, desativar ou configurar rotina do Cron",
                    "security": sec_ref,
                    "parameters": [{"name": "task_id", "in": "path", "required": True, "schema": {"type": "string"}}],
                    "responses": {"200": {"description": "Rotina atualizada"}},
                }
            },
            "/api/v1/oxe": {
                "get": {
                    "summary": "Consultar ramais da Central Telefônica Alcatel OXE",
                    "security": sec_ref,
                    "responses": {"200": {"description": "Lista de ramais OXE"}},
                }
            },
            "/api/v1/impressoras": {
                "get": {
                    "summary": "Consultar impressoras e dispositivos do PaperCut",
                    "security": sec_ref,
                    "responses": {"200": {"description": "Lista de impressoras"}},
                }
            },
            "/api/v1/ramais": {
                "get": {
                    "summary": "Consultar catálogo de ramais MPMS",
                    "security": sec_ref,
                    "responses": {"200": {"description": "Lista de ramais"}},
                }
            },
            "/api/v1/unidades": {
                "get": {
                    "summary": "Consultar catálogo de unidades MPMS",
                    "security": sec_ref,
                    "responses": {"200": {"description": "Lista de unidades"}},
                }
            },
            "/api/v1/sccm": {
                "get": {
                    "summary": "Consultar inventário de estações SCCM",
                    "security": sec_ref,
                    "responses": {"200": {"description": "Lista de dispositivos"}},
                }
            },
        },
    }


SWAGGER_HTML = """<!DOCTYPE html>
<html lang="pt-BR">
<head>
  <meta charset="utf-8" />
  <meta name="viewport" content="width=device-width, initial-scale=1" />
  <title>Bancada STI — Documentação OpenAPI</title>
  <link rel="stylesheet" href="https://unpkg.com/swagger-ui-dist@5/swagger-ui.css" />
  <style>
    body { margin: 0; font-family: system-ui, -apple-system, sans-serif; background: #0f172a; color: #e2e8f0; }
    .fallback-header { padding: 16px 24px; background: #1e293b; border-bottom: 1px solid #334155; }
    .fallback-header code { background: #0f172a; padding: 2px 6px; border-radius: 4px; color: #38bdf8; }
  </style>
</head>
<body>
  <div class="fallback-header">
    <strong>Bancada STI REST API v1</strong> — Autenticação via header <code>X-API-Key</code> ou <code>Authorization: Bearer</code>.
    Schema JSON: <a href="/openapi.json" style="color:#38bdf8">/openapi.json</a>
  </div>
  <div id="swagger-ui"></div>
  <script src="https://unpkg.com/swagger-ui-dist@5/swagger-ui-bundle.js"></script>
  <script>
    if (window.SwaggerUIBundle) {
      window.SwaggerUIBundle({
        url: '/openapi.json',
        dom_id: '#swagger-ui',
        presets: [SwaggerUIBundle.presets.apis]
      });
    }
  </script>
</body>
</html>
"""
