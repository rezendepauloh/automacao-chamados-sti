# -*- coding: utf-8 -*-
"""
Testes unitários e funcionais da API REST de Integração (src.api.*).
Valida autenticação Fail-Closed, Rate Limiting, validação Pydantic,
leitura/edição de chamados, comentários, Cron, OXE, PaperCut, Ramais, Unidades e SCCM.
"""

import asyncio
import json
import os
import sqlite3
import sys
import tempfile
import unittest
from pathlib import Path
from typing import Dict, List, Optional, Tuple
from unittest.mock import patch

ROOT_DIR = Path(__file__).resolve().parent.parent.parent
if str(ROOT_DIR) not in sys.path:
    sys.path.insert(0, str(ROOT_DIR))

import tests.test_helpers  # noqa: F401
from src.api.app import app, rate_limiter
from src.database.cron_db import setup_cron_tables
from src.database.impressoras_db import setup_impressoras_table
from src.database.ramais_db import setup_ramais_table
from src.database.sccm_db import setup_sccm_tables
from src.database.tickets_db import setup_database
from src.database.unidades_db import setup_unidades_tables

TEST_API_KEY = "unit_test_super_secure_api_key_2026"


def call_asgi(
    method: str,
    path: str,
    query_string: str = "",
    headers: Optional[Dict[str, str]] = None,
    json_body: Optional[dict] = None,
    client_ip: str = "127.0.0.1",
) -> Tuple[int, Dict[str, str], dict]:
    """Executa uma requisição ASGI diretamente contra a aplicação Starlette sem rede."""
    body_bytes = json.dumps(json_body).encode("utf-8") if json_body is not None else b""
    raw_headers: List[Tuple[bytes, bytes]] = []
    if json_body is not None:
        raw_headers.append((b"content-type", b"application/json"))
    for k, v in (headers or {}).items():
        raw_headers.append((k.lower().encode("latin-1"), v.encode("latin-1")))

    scope = {
        "type": "http",
        "asgi": {"version": "3.0"},
        "http_version": "1.1",
        "method": method.upper(),
        "scheme": "http",
        "path": path,
        "raw_path": path.encode("latin-1"),
        "query_string": query_string.encode("latin-1"),
        "headers": raw_headers,
        "client": (client_ip, 50000),
        "server": ("127.0.0.1", 8502),
    }

    sent_body = False
    response_status = 500
    response_headers: Dict[str, str] = {}
    response_chunks: List[bytes] = []

    async def receive():
        nonlocal sent_body
        if not sent_body:
            sent_body = True
            return {"type": "http.request", "body": body_bytes, "more_body": False}
        return {"type": "http.disconnect"}

    async def send(message):
        nonlocal response_status, response_headers
        if message["type"] == "http.response.start":
            response_status = message["status"]
            response_headers = {
                k.decode("latin-1").lower(): v.decode("latin-1")
                for k, v in message.get("headers", [])
            }
        elif message["type"] == "http.response.body":
            response_chunks.append(message.get("body", b""))

    asyncio.run(app(scope, receive, send))
    raw_text = b"".join(response_chunks).decode("utf-8", errors="replace")
    try:
        parsed = json.loads(raw_text) if raw_text else {}
    except json.JSONDecodeError:
        parsed = {"raw": raw_text}
    return response_status, response_headers, parsed


class TestBancadaRestApi(unittest.TestCase):
    def setUp(self):
        rate_limiter.reset()
        self.temp_dir = tempfile.TemporaryDirectory()
        self.db_path = Path(self.temp_dir.name) / "test_api.db"

        def _connect():
            conn = sqlite3.connect(self.db_path)
            return conn

        self.patchers = [
            patch("src.database.connection.get_connection", side_effect=_connect),
            patch("src.database.tickets_db.get_connection", side_effect=_connect),
            patch("src.database.cron_db.get_connection", side_effect=_connect),
            patch("src.database.impressoras_db.get_connection", side_effect=_connect),
            patch("src.database.ramais_db.get_connection", side_effect=_connect),
            patch("src.database.unidades_db.get_connection", side_effect=_connect),
            patch("src.database.sccm_db.get_connection", side_effect=_connect),
            patch("src.api.service.get_connection", side_effect=_connect),
            patch("src.database.connection.DB_TYPE", "sqlite"),
            patch("src.database.cron_db.DB_TYPE", "sqlite"),
            patch("src.database.sccm_db.DB_TYPE", "sqlite"),
            patch("src.database.cron_db._cron_tables_initialized", False),
            patch.dict(
                os.environ,
                {
                    "DB_TYPE": "sqlite",
                    "BANCADA_API_KEY": TEST_API_KEY,
                    "BANCADA_API_RATE_LIMIT": "180",
                    "BANCADA_API_DOCS_ENABLED": "true",
                },
                clear=False,
            ),
        ]
        for p in self.patchers:
            p.start()

        self._seed_test_database()

    def tearDown(self):
        for p in self.patchers:
            p.stop()
        self.temp_dir.cleanup()

    def _seed_test_database(self):
        setup_database()
        setup_cron_tables(force=True)
        setup_impressoras_table()
        setup_ramais_table()
        setup_unidades_tables()
        setup_sccm_tables()

        conn = sqlite3.connect(self.db_path)
        cur = conn.cursor()

        cur.execute(
            """
            INSERT INTO chamados (
                id, data_criacao, titulo, cidade_predio, unidade, localidade_fisica,
                usuario, id_cliente, descricao, tag, ip_origem, status, data_atualizacao,
                base, andamento, link, hostname, tag_manual, dados_manuais
            ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
            """,
            (
                "2026001",
                "2026-11-17 09:00:00",
                "Sem acesso à impressora",
                "Campo Grande - Sede",
                "STI",
                "Sede",
                "usuario.teste",
                "usuario.teste",
                "Impressora não imprime",
                "Impressora",
                "10.0.0.10",
                "Aberto",
                "2026-11-17 09:30:00",
                "OTRS",
                "Em triagem",
                "https://otrs.local/ticket/2026001",
                "ESTACAO-01",
                0,
                0,
            ),
        )
        cur.execute(
            """
            INSERT INTO chamados (
                id, data_criacao, titulo, cidade_predio, unidade, localidade_fisica,
                usuario, id_cliente, descricao, tag, ip_origem, status, data_atualizacao,
                base, andamento, link, hostname, tag_manual, dados_manuais
            ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
            """,
            (
                "INC-500",
                "2026-11-16 14:00:00",
                "Troca de monitor",
                "Dourados",
                "1a PJ",
                "Dourados",
                "outro.usuario",
                "outro.usuario",
                "Monitor piscando",
                "Hardware",
                "10.0.1.20",
                "Fechado",
                "2026-11-16 16:00:00",
                "CitSmart",
                "Substituído",
                "https://citsmart.local/500",
                "ESTACAO-02",
                0,
                0,
            ),
        )
        cur.execute(
            "INSERT INTO comentarios (chamado_id, data, autor, texto) VALUES (?, ?, ?, ?)",
            ("2026001", "2026-11-17 09:10:00", "Analista", "Verificando fila PaperCut"),
        )
        cur.execute(
            """
            CREATE TABLE IF NOT EXISTS central_telefonica (
                ramal TEXT,
                nome TEXT,
                localidade TEXT,
                setor TEXT,
                tipo TEXT,
                ip TEXT,
                mac TEXT
            )
            """
        )
        cur.execute(
            "INSERT INTO central_telefonica VALUES ('4100', 'Recepcao STI', 'Sede', 'STI', '8028', '10.10.1.5', '00:11:22:33:44:55')"
        )
        cur.execute(
            """
            INSERT INTO impressoras (nome, servidor, tipo, modelo, localizacao, ip_host, status, total_paginas, filas_relacionadas, detalhes_extra, data_atualizacao)
            VALUES ('IMP-STI-01', 'PRINTSRV01', 'MFD', 'Lexmark MX', 'Sede', '10.10.2.15', 'OK', 12345, 'Fila_STI', '', '2026-11-17 08:00:00')
            """
        )
        cur.execute(
            "INSERT INTO ramais_mpms (localidade, setor_nome, telefone_ramal, tipo, data_atualizacao) VALUES ('Sede', 'Bancada STI', '4100', 'Interno', '2026-11-17')"
        )
        cur.execute(
            "INSERT INTO unidades (cidade, tipo, setor, sigla, titular, unidade_predio, telefone, url, origem, data_atualizacao) VALUES ('Campo Grande', 'Sede', 'STI', 'STI', 'Titular Teste', 'Predio Sede', '4100', 'https://mpms.local', 'web', '2026-11-17')"
        )
        cur.execute(
            "INSERT INTO sccm_cache_devices (resource_id, name, last_logon_user, ip_addresses, model, operating_system, client_active) VALUES ('1001', 'ESTACAO-01', 'usuario.teste', '10.0.0.10', 'Lenovo M70q', 'Windows 11', 1)"
        )
        conn.commit()
        conn.close()

    def test_health_and_docs_endpoints(self):
        status, headers, data = call_asgi("GET", "/health")
        self.assertEqual(status, 200)
        self.assertEqual(data["status"], "ok")
        self.assertTrue(data["db_ok"])
        self.assertEqual(headers.get("x-content-options") or headers.get("x-content-type-options"), "nosniff")

        status_spec, _, spec = call_asgi("GET", "/openapi.json")
        self.assertEqual(status_spec, 200)
        self.assertEqual(spec["openapi"], "3.0.3")
        self.assertIn("/api/v1/chamados/{chamado_id}", spec["paths"])

    def test_auth_fail_closed_and_invalid_credentials(self):
        # 1. Sem BANCADA_API_KEY configurada -> 503 Fail-Closed
        with patch.dict(os.environ, {"BANCADA_API_KEY": ""}, clear=False):
            status, _, data = call_asgi("GET", "/api/v1/info", headers={"X-API-Key": TEST_API_KEY})
            self.assertEqual(status, 503)
            self.assertIn("Fail-Closed", data["detail"])

        # 2. Chave muito curta -> 503 Fail-Closed
        with patch.dict(os.environ, {"BANCADA_API_KEY": "curta"}, clear=False):
            status, _, _ = call_asgi("GET", "/api/v1/info", headers={"X-API-Key": "curta"})
            self.assertEqual(status, 503)

        # 3. Sem header -> 401
        status, _, _ = call_asgi("GET", "/api/v1/info")
        self.assertEqual(status, 401)

        # 4. Chave errada -> 403
        status, _, _ = call_asgi("GET", "/api/v1/info", headers={"X-API-Key": "wrong_key_1234567890"})
        self.assertEqual(status, 403)

        # 5. Via Authorization: Bearer -> 200
        status, _, data = call_asgi("GET", "/api/v1/info", headers={"Authorization": f"Bearer {TEST_API_KEY}"})
        self.assertEqual(status, 200)
        self.assertEqual(data["service"], "bancada-api")

    def test_rate_limiting_blocks_excessive_requests(self):
        with patch.dict(os.environ, {"BANCADA_API_RATE_LIMIT": "3"}, clear=False):
            rate_limiter.reset()
            for _ in range(3):
                st, _, _ = call_asgi("GET", "/api/v1/info", headers={"X-API-Key": TEST_API_KEY}, client_ip="10.9.9.9")
                self.assertEqual(st, 200)

            st_blocked, headers, _ = call_asgi("GET", "/api/v1/info", headers={"X-API-Key": TEST_API_KEY}, client_ip="10.9.9.9")
            self.assertEqual(st_blocked, 429)
            self.assertIn("retry-after", headers)

    def test_chamados_query_filters_and_edit_lifecycle(self):
        auth = {"X-API-Key": TEST_API_KEY}

        # Listagem filtrada por status Aberto
        st, _, res = call_asgi("GET", "/api/v1/chamados", query_string="status=Aberto", headers=auth)
        self.assertEqual(st, 200)
        self.assertEqual(res["summary"]["total"], 2)
        self.assertEqual(res["summary"]["abertos"], 1)
        self.assertEqual(len(res["items"]), 1)
        self.assertEqual(res["items"][0]["id"], "2026001")

        # Localidades e filtros dinâmicos
        st_loc, _, locs = call_asgi("GET", "/api/v1/chamados/localidades", headers=auth)
        self.assertEqual(st_loc, 200)
        self.assertIn("Sede", locs["localidades"])

        st_flt, _, flt = call_asgi("GET", "/api/v1/chamados/filters", query_string="status=Todos", headers=auth)
        self.assertEqual(st_flt, 200)
        self.assertIn("Impressora", flt["tags"])

        # Edição PATCH parcial e completa
        st_upd, _, upd = call_asgi(
            "PATCH",
            "/api/v1/chamados/2026001",
            headers=auth,
            json_body={
                "status": "Em Atendimento",
                "hostname": "ESTACAO-99",
                "tag": "Rede",
                "unidade": "STI-Manutencao",
                "andamento": "Atualizado via API REST",
            },
        )
        self.assertEqual(st_upd, 200)
        self.assertTrue(upd["success"])
        self.assertEqual(upd["chamado"]["hostname"], "ESTACAO-99")
        self.assertEqual(upd["chamado"]["tag"], "Rede")
        self.assertTrue(upd["chamado"]["tag_manual"])
        self.assertTrue(upd["chamado"]["dados_manuais"])

        # Validação rejeita hostname com caracteres de shell
        st_bad, _, _ = call_asgi(
            "PATCH",
            "/api/v1/chamados/2026001",
            headers=auth,
            json_body={"hostname": "host;rm -rf /"},
        )
        self.assertEqual(st_bad, 422)

        # Chamado inexistente -> 404
        st_404, _, _ = call_asgi("PATCH", "/api/v1/chamados/9999999", headers=auth, json_body={"status": "Fechado"})
        self.assertEqual(st_404, 404)

    def test_chamados_comentarios_get_post_and_batch(self):
        auth = {"X-API-Key": TEST_API_KEY}

        st_post, _, created = call_asgi(
            "POST",
            "/api/v1/chamados/2026001/comentarios",
            headers=auth,
            json_body={"autor": "AmbientesHub", "texto": "Driver reinstalado remotamente"},
        )
        self.assertEqual(st_post, 201)
        self.assertEqual(created["comentario"]["autor"], "AmbientesHub")

        st_batch, _, batch = call_asgi(
            "POST",
            "/api/v1/chamados/comentarios/batch",
            headers=auth,
            json_body={"chamado_ids": ["2026001", "INC-500"]},
        )
        self.assertEqual(st_batch, 200)
        self.assertEqual(len(batch["comentarios"]["2026001"]), 2)

    def test_cron_query_trigger_and_toggle(self):
        auth = {"X-API-Key": TEST_API_KEY}

        st, _, cron_data = call_asgi("GET", "/api/v1/cron", headers=auth)
        self.assertEqual(st, 200)
        self.assertGreater(cron_data["summary"]["total"], 0)

        # Disparo remoto
        st_trig, _, trig = call_asgi("POST", "/api/v1/cron/sync_oxe/trigger", headers=auth)
        self.assertEqual(st_trig, 200)
        self.assertTrue(trig["success"])

        # Alterna ativo=false
        st_patch, _, patched = call_asgi(
            "PATCH",
            "/api/v1/cron/sync_oxe",
            headers=auth,
            json_body={"ativo": False},
        )
        self.assertEqual(st_patch, 200)
        self.assertTrue(patched["success"])

        # Confirma persistência
        _, _, cron_after = call_asgi("GET", "/api/v1/cron", headers=auth)
        oxe_row = next(s for s in cron_after["schedules"] if s["task_id"] == "sync_oxe")
        self.assertFalse(oxe_row["ativo"])
        self.assertTrue(oxe_row["run_requested"])

    def test_oxe_impressoras_ramais_unidades_sccm_endpoints(self):
        auth = {"X-API-Key": TEST_API_KEY}

        st_oxe, _, oxe = call_asgi("GET", "/api/v1/oxe", query_string="search=4100", headers=auth)
        self.assertEqual(st_oxe, 200)
        self.assertEqual(oxe["summary"]["total"], 1)
        self.assertEqual(oxe["items"][0]["ramal"], "4100")

        st_imp, _, imp = call_asgi("GET", "/api/v1/impressoras", headers=auth)
        self.assertEqual(st_imp, 200)
        self.assertEqual(imp["summary"]["ok_count"], 1)
        self.assertEqual(imp["summary"]["total_paginas_soma"], 12345)

        st_ram, _, ram = call_asgi("GET", "/api/v1/ramais", query_string="search=Bancada", headers=auth)
        self.assertEqual(st_ram, 200)
        self.assertEqual(ram["count"], 1)

        st_uni, _, uni = call_asgi("GET", "/api/v1/unidades", query_string="search=STI", headers=auth)
        self.assertEqual(st_uni, 200)
        self.assertEqual(uni["count"], 1)

        st_sccm, _, sccm = call_asgi("GET", "/api/v1/sccm", query_string="search=ESTACAO-01", headers=auth)
        self.assertEqual(st_sccm, 200)
        self.assertEqual(sccm["count"], 1)


if __name__ == "__main__":
    unittest.main()
