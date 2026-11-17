# -*- coding: utf-8 -*-
"""
Aplicação ASGI (Starlette) da API REST do Sistema Bancada STI.
Projetada para execução isolada do Streamlit (container separado ou worker dedicado),
com autenticação Fail-Closed, Rate Limiting e validação Pydantic v2.
"""

import json
import logging
import os
from typing import List

from pydantic import ValidationError
from starlette.applications import Starlette
from starlette.middleware import Middleware
from starlette.middleware.base import BaseHTTPMiddleware
from starlette.middleware.cors import CORSMiddleware
from starlette.requests import Request
from starlette.responses import HTMLResponse, JSONResponse
from starlette.routing import Route

from src.api.openapi import SWAGGER_HTML, build_openapi_spec
from src.api.schemas import (
    ChamadoUpdatePayload,
    ComentarioCreatePayload,
    ComentariosBatchPayload,
    CronUpdatePayload,
)
from src.api.security import SlidingWindowRateLimiter, verify_api_key_headers
from src.api import service

logger = logging.getLogger(__name__)

rate_limiter = SlidingWindowRateLimiter()


def _is_docs_enabled() -> bool:
    return os.getenv("BANCADA_API_DOCS_ENABLED", "true").strip().lower() in ("1", "true", "yes", "sim")


def _parse_localidades_query(request: Request) -> List[str]:
    vals = request.query_params.getlist("localidades")
    result: List[str] = []
    for raw in vals:
        for part in str(raw).split(","):
            clean = part.strip()
            if clean:
                result.append(clean)
    return result


async def _parse_json_body(request: Request) -> dict:
    try:
        data = await request.json()
        if not isinstance(data, dict):
            raise ValueError("O corpo JSON deve ser um objeto.")
        return data
    except json.JSONDecodeError as exc:
        raise ValueError(f"JSON inválido: {exc.msg}") from exc


class SecurityAndRateLimitMiddleware(BaseHTTPMiddleware):
    """Middleware central: cabeçalhos de segurança, Rate Limiting e Auth Fail-Closed em /api/*."""

    async def dispatch(self, request: Request, call_next):
        path = request.url.path
        client_ip = request.client.host if request.client else "unknown"

        # Rate limiting aplicado em todas as rotas de API
        if path.startswith("/api/"):
            allowed, remaining, retry_after = rate_limiter.is_allowed(client_ip)
            if not allowed:
                return JSONResponse(
                    {"detail": "Limite de requisições excedido. Aguarde antes de tentar novamente."},
                    status_code=429,
                    headers={"Retry-After": str(retry_after)},
                )

            ok, status_code, err_msg = verify_api_key_headers(request.headers)
            if not ok:
                return JSONResponse({"detail": err_msg}, status_code=status_code)

        response = await call_next(request)
        response.headers["X-Content-Type-Options"] = "nosniff"
        if path.startswith("/api/"):
            response.headers["Cache-Control"] = "no-store"
        return response


async def health_endpoint(request: Request):
    status = service.get_health_status()
    http_code = 200 if status["db_ok"] else 503
    return JSONResponse(status, status_code=http_code)


async def openapi_endpoint(request: Request):
    if not _is_docs_enabled():
        return JSONResponse({"detail": "Documentação desabilitada."}, status_code=404)
    return JSONResponse(build_openapi_spec())


async def docs_endpoint(request: Request):
    if not _is_docs_enabled():
        return JSONResponse({"detail": "Documentação desabilitada."}, status_code=404)
    return HTMLResponse(SWAGGER_HTML)


async def info_endpoint(request: Request):
    return JSONResponse(
        {
            "service": "bancada-api",
            "version": "1.0.0",
            "db_type": os.getenv("DB_TYPE", "sqlite").lower(),
            "endpoints": [
                "/health",
                "/api/v1/chamados",
                "/api/v1/chamados/{id}",
                "/api/v1/chamados/{id}/comentarios",
                "/api/v1/chamados/comentarios/batch",
                "/api/v1/cron",
                "/api/v1/cron/{task_id}/trigger",
                "/api/v1/cron/{task_id}",
                "/api/v1/oxe",
                "/api/v1/impressoras",
                "/api/v1/ramais",
                "/api/v1/unidades",
                "/api/v1/sccm",
            ],
        }
    )


async def chamados_list_endpoint(request: Request):
    qp = request.query_params
    try:
        limit = int(qp.get("limit", "150"))
        offset = int(qp.get("offset", "0"))
    except ValueError:
        return JSONResponse({"detail": "Parâmetros 'limit' e 'offset' devem ser numéricos."}, status_code=422)

    data = service.query_chamados(
        search=qp.get("search"),
        base_filter=qp.get("base"),
        status_filter=qp.get("status", "Aberto"),
        tag_filter=qp.get("tag"),
        localidades_filter=_parse_localidades_query(request),
        limit=limit,
        offset=offset,
    )
    return JSONResponse(data)


async def chamados_localidades_endpoint(request: Request):
    locs = service.get_distinct_localidades()
    return JSONResponse({"localidades": locs})


async def chamados_filters_endpoint(request: Request):
    qp = request.query_params
    data = service.get_chamados_filters(
        search=qp.get("search"),
        base_filter=qp.get("base"),
        status_filter=qp.get("status", "Aberto"),
        tag_filter=qp.get("tag"),
        localidades_filter=_parse_localidades_query(request),
    )
    return JSONResponse(data)


async def chamado_detail_endpoint(request: Request):
    cid = request.path_params["chamado_id"]
    chamado = service.get_chamado_by_id(cid)
    if not chamado:
        return JSONResponse({"detail": f"Chamado #{cid} não encontrado."}, status_code=404)
    comentarios = service.get_comentarios_chamado(cid)
    return JSONResponse({"chamado": chamado, "comentarios": comentarios})


async def chamado_update_endpoint(request: Request):
    cid = request.path_params["chamado_id"]
    try:
        raw = await _parse_json_body(request)
        payload = ChamadoUpdatePayload.model_validate(raw)
    except ValueError as exc:
        return JSONResponse({"detail": str(exc)}, status_code=422)
    except ValidationError as exc:
        return JSONResponse({"detail": exc.errors()}, status_code=422)

    updated = service.update_chamado(cid, payload)
    if not updated:
        return JSONResponse({"detail": f"Chamado #{cid} não encontrado."}, status_code=404)
    return JSONResponse({"success": True, "message": f"Chamado #{cid} atualizado com sucesso.", "chamado": updated})


async def chamado_comentarios_get_endpoint(request: Request):
    cid = request.path_params["chamado_id"]
    items = service.get_comentarios_chamado(cid)
    return JSONResponse({"chamado_id": cid, "comentarios": items})


async def chamado_comentarios_post_endpoint(request: Request):
    cid = request.path_params["chamado_id"]
    try:
        raw = await _parse_json_body(request)
        payload = ComentarioCreatePayload.model_validate(raw)
    except ValueError as exc:
        return JSONResponse({"detail": str(exc)}, status_code=422)
    except ValidationError as exc:
        return JSONResponse({"detail": exc.errors()}, status_code=422)

    created = service.add_comentario_chamado(cid, payload)
    if not created:
        return JSONResponse({"detail": f"Chamado #{cid} não encontrado."}, status_code=404)
    return JSONResponse({"success": True, "comentario": created}, status_code=201)


async def chamados_comentarios_batch_endpoint(request: Request):
    try:
        raw = await _parse_json_body(request)
        payload = ComentariosBatchPayload.model_validate(raw)
    except ValueError as exc:
        return JSONResponse({"detail": str(exc)}, status_code=422)
    except ValidationError as exc:
        return JSONResponse({"detail": exc.errors()}, status_code=422)

    data = service.get_comentarios_batch(payload.chamado_ids)
    return JSONResponse({"comentarios": data})


async def cron_list_endpoint(request: Request):
    data = service.query_cron_data()
    return JSONResponse(data)


async def cron_trigger_endpoint(request: Request):
    task_id = request.path_params["task_id"]
    ok, msg, code = service.trigger_cron_task(task_id)
    return JSONResponse({"success": ok, "task_id": task_id, "message": msg}, status_code=code)


async def cron_update_endpoint(request: Request):
    task_id = request.path_params["task_id"]
    try:
        raw = await _parse_json_body(request)
        payload = CronUpdatePayload.model_validate(raw)
    except ValueError as exc:
        return JSONResponse({"detail": str(exc)}, status_code=422)
    except ValidationError as exc:
        return JSONResponse({"detail": exc.errors()}, status_code=422)

    ok, msg, code = service.update_cron_task(task_id, payload)
    return JSONResponse({"success": ok, "task_id": task_id, "message": msg}, status_code=code)


async def oxe_list_endpoint(request: Request):
    qp = request.query_params
    try:
        limit = int(qp.get("limit", "2000"))
    except ValueError:
        return JSONResponse({"detail": "Parâmetro 'limit' deve ser numérico."}, status_code=422)
    data = service.query_oxe_ramais(search=qp.get("search"), limit=limit)
    return JSONResponse(data)


async def impressoras_list_endpoint(request: Request):
    qp = request.query_params
    data = service.query_impressoras(search=qp.get("search"))
    return JSONResponse(data)


async def ramais_list_endpoint(request: Request):
    qp = request.query_params
    try:
        limit = int(qp.get("limit", "500"))
    except ValueError:
        return JSONResponse({"detail": "Parâmetro 'limit' deve ser numérico."}, status_code=422)
    data = service.query_ramais_mpms(search=qp.get("search"), limit=limit)
    return JSONResponse(data)


async def unidades_list_endpoint(request: Request):
    qp = request.query_params
    try:
        limit = int(qp.get("limit", "500"))
    except ValueError:
        return JSONResponse({"detail": "Parâmetro 'limit' deve ser numérico."}, status_code=422)
    data = service.query_unidades_mpms(search=qp.get("search"), limit=limit)
    return JSONResponse(data)


async def sccm_list_endpoint(request: Request):
    qp = request.query_params
    try:
        limit = int(qp.get("limit", "200"))
    except ValueError:
        return JSONResponse({"detail": "Parâmetro 'limit' deve ser numérico."}, status_code=422)
    data = service.query_sccm_devices(search=qp.get("search"), limit=limit)
    return JSONResponse(data)


async def global_exception_handler(request: Request, exc: Exception):
    logger.exception(f"Erro não tratado na API ({request.method} {request.url.path}): {exc}")
    return JSONResponse(
        {"detail": "Erro interno ao processar requisição no servidor da Bancada."},
        status_code=500,
    )


def create_app() -> Starlette:
    routes = [
        Route("/health", health_endpoint, methods=["GET"]),
        Route("/openapi.json", openapi_endpoint, methods=["GET"]),
        Route("/docs", docs_endpoint, methods=["GET"]),
        Route("/api/v1/info", info_endpoint, methods=["GET"]),
        Route("/api/v1/chamados", chamados_list_endpoint, methods=["GET"]),
        Route("/api/v1/chamados/localidades", chamados_localidades_endpoint, methods=["GET"]),
        Route("/api/v1/chamados/filters", chamados_filters_endpoint, methods=["GET"]),
        Route("/api/v1/chamados/comentarios/batch", chamados_comentarios_batch_endpoint, methods=["POST"]),
        Route("/api/v1/chamados/{chamado_id}", chamado_detail_endpoint, methods=["GET"]),
        Route("/api/v1/chamados/{chamado_id}", chamado_update_endpoint, methods=["PATCH", "PUT"]),
        Route("/api/v1/chamados/{chamado_id}/comentarios", chamado_comentarios_get_endpoint, methods=["GET"]),
        Route("/api/v1/chamados/{chamado_id}/comentarios", chamado_comentarios_post_endpoint, methods=["POST"]),
        Route("/api/v1/cron", cron_list_endpoint, methods=["GET"]),
        Route("/api/v1/cron/{task_id}/trigger", cron_trigger_endpoint, methods=["POST"]),
        Route("/api/v1/cron/{task_id}", cron_update_endpoint, methods=["PATCH"]),
        Route("/api/v1/oxe", oxe_list_endpoint, methods=["GET"]),
        Route("/api/v1/impressoras", impressoras_list_endpoint, methods=["GET"]),
        Route("/api/v1/ramais", ramais_list_endpoint, methods=["GET"]),
        Route("/api/v1/unidades", unidades_list_endpoint, methods=["GET"]),
        Route("/api/v1/sccm", sccm_list_endpoint, methods=["GET"]),
    ]

    middleware = [Middleware(SecurityAndRateLimitMiddleware)]
    cors_origins = [
        o.strip()
        for o in os.getenv("BANCADA_API_CORS_ORIGINS", "").split(",")
        if o.strip()
    ]
    if cors_origins:
        middleware.append(
            Middleware(
                CORSMiddleware,
                allow_origins=cors_origins,
                allow_methods=["GET", "POST", "PUT", "PATCH", "OPTIONS"],
                allow_headers=["X-API-Key", "Authorization", "Content-Type"],
            )
        )

    return Starlette(
        debug=False,
        routes=routes,
        middleware=middleware,
        exception_handlers={Exception: global_exception_handler},
    )


app = create_app()
