# -*- coding: utf-8 -*-
"""
Camada de serviço da API REST Bancada STI.
Centraliza consultas e edições parametrizadas com compatibilidade híbrida (PostgreSQL e SQLite).
"""

import logging
from contextlib import contextmanager
from datetime import datetime
from typing import Any, Dict, Iterable, List, Optional, Tuple

from src.database.connection import get_connection
from src.api.schemas import ChamadoUpdatePayload, ComentarioCreatePayload, CronUpdatePayload

logger = logging.getLogger(__name__)


@contextmanager
def _db_cursor():
    """Abre conexão do pool e garante fechamento de cursor/conexão ao finalizar."""
    conn = get_connection()
    cursor = conn.cursor()
    try:
        yield conn, cursor
    finally:
        try:
            cursor.close()
        except Exception:
            pass
        try:
            conn.close()
        except Exception:
            pass


def _clean_str(val: Any, default: str = "") -> str:
    if val is None:
        return default
    s = str(val).strip()
    if s.lower() in ("nan", "none", "null", "<na>"):
        return default
    return s


def _normalize_ticket_id(cid: str) -> str:
    s = str(cid or "").strip()
    if s.endswith(".0"):
        s = s[:-2]
    return s


def _has_real_value(val: Optional[str]) -> bool:
    if not val:
        return False
    t = val.strip()
    return t != "-" and t.lower() not in ("n/a", "nan", "none", "null")


def _first_non_empty(fields: Dict[str, str], *keys: str) -> str:
    lower_map = {k.lower(): v for k, v in fields.items()}
    for k in keys:
        v = lower_map.get(k.lower(), "")
        if _has_real_value(v):
            return v
    return ""


def _build_chamados_where(
    search: Optional[str] = None,
    base_filter: Optional[str] = None,
    status_filter: Optional[str] = "Aberto",
    tag_filter: Optional[str] = None,
    localidades_filter: Optional[List[str]] = None,
) -> Tuple[str, List[Any]]:
    clauses: List[str] = []
    params: List[Any] = []

    if base_filter and base_filter.strip() and base_filter.strip().lower() != "todas as bases":
        clauses.append("LOWER(COALESCE(base, '')) LIKE ?")
        params.append(f"%{base_filter.strip().lower()}%")

    if status_filter and status_filter.strip() and status_filter.strip().lower() != "todos":
        clauses.append("LOWER(COALESCE(status, '')) = ?")
        params.append(status_filter.strip().lower())

    if tag_filter and tag_filter.strip() and tag_filter.strip().lower() != "todas as tags":
        clauses.append("LOWER(COALESCE(tag, '')) = ?")
        params.append(tag_filter.strip().lower())

    clean_locs = [
        loc.strip().lower()
        for loc in (localidades_filter or [])
        if loc and str(loc).strip()
    ]
    if clean_locs:
        unique_locs = list(dict.fromkeys(clean_locs))
        placeholders = ", ".join(["?"] * len(unique_locs))
        clauses.append(f"LOWER(TRIM(COALESCE(localidade_fisica, ''))) IN ({placeholders})")
        params.extend(unique_locs)

    if search and search.strip():
        q = f"%{search.strip().lower()}%"
        clauses.append(
            "("
            "LOWER(COALESCE(id, '')) LIKE ? OR "
            "LOWER(COALESCE(titulo, '')) LIKE ? OR "
            "LOWER(COALESCE(usuario, '')) LIKE ? OR "
            "LOWER(COALESCE(unidade, '')) LIKE ? OR "
            "LOWER(COALESCE(tag, '')) LIKE ? OR "
            "LOWER(COALESCE(hostname, '')) LIKE ?"
            ")"
        )
        params.extend([q, q, q, q, q, q])

    where_sql = f"WHERE {' AND '.join(clauses)}" if clauses else ""
    return where_sql, params


def get_health_status() -> Dict[str, Any]:
    """Verifica disponibilidade do serviço e do banco de dados sem expor credenciais."""
    db_ok = False
    try:
        with _db_cursor() as (_, cursor):
            cursor.execute("SELECT 1")
            row = cursor.fetchone()
            db_ok = bool(row and row[0] == 1)
    except Exception as exc:
        logger.warning(f"Falha no healthcheck de banco da API: {exc}")
        db_ok = False

    return {
        "status": "ok" if db_ok else "degraded",
        "service": "bancada-api",
        "db_ok": db_ok,
        "timestamp": datetime.now().isoformat(timespec="seconds"),
    }


def get_distinct_localidades() -> List[str]:
    with _db_cursor() as (_, cursor):
        cursor.execute(
            """
            SELECT DISTINCT TRIM(localidade_fisica) AS loc
            FROM chamados
            WHERE TRIM(COALESCE(localidade_fisica, '')) <> ''
            ORDER BY loc ASC
            """
        )
        return [_clean_str(r[0]) for r in cursor.fetchall() if _clean_str(r[0])]


def get_chamados_filters(
    search: Optional[str] = None,
    base_filter: Optional[str] = None,
    status_filter: Optional[str] = "Aberto",
    tag_filter: Optional[str] = None,
    localidades_filter: Optional[List[str]] = None,
) -> Dict[str, List[str]]:
    tags: List[str] = []
    locs: List[str] = []

    with _db_cursor() as (_, cursor):
        where_tags, params_tags = _build_chamados_where(
            search=search,
            base_filter=base_filter,
            status_filter=status_filter,
            tag_filter=None,
            localidades_filter=localidades_filter,
        )
        prefix_tags = f"{where_tags} AND" if where_tags else "WHERE"
        cursor.execute(
            f"SELECT DISTINCT TRIM(tag) FROM chamados {prefix_tags} TRIM(COALESCE(tag, '')) <> '' ORDER BY 1 ASC",
            tuple(params_tags),
        )
        tags = [_clean_str(r[0]) for r in cursor.fetchall() if _clean_str(r[0])]

        where_locs, params_locs = _build_chamados_where(
            search=search,
            base_filter=base_filter,
            status_filter=status_filter,
            tag_filter=tag_filter,
            localidades_filter=None,
        )
        prefix_locs = f"{where_locs} AND" if where_locs else "WHERE"
        cursor.execute(
            f"SELECT DISTINCT TRIM(localidade_fisica) FROM chamados {prefix_locs} TRIM(COALESCE(localidade_fisica, '')) <> '' ORDER BY 1 ASC",
            tuple(params_locs),
        )
        locs = [_clean_str(r[0]) for r in cursor.fetchall() if _clean_str(r[0])]

    return {"tags": tags, "localidades": locs}


def query_chamados(
    search: Optional[str] = None,
    base_filter: Optional[str] = None,
    status_filter: Optional[str] = "Aberto",
    tag_filter: Optional[str] = None,
    localidades_filter: Optional[List[str]] = None,
    limit: int = 150,
    offset: int = 0,
) -> Dict[str, Any]:
    safe_limit = max(1, min(int(limit), 500))
    safe_offset = max(0, int(offset))

    items: List[Dict[str, Any]] = []
    summary = {"total": 0, "abertos": 0, "otrs_count": 0, "citsmart_count": 0}

    with _db_cursor() as (_, cursor):
        cursor.execute(
            """
            SELECT
                COUNT(*),
                COALESCE(SUM(CASE WHEN LOWER(COALESCE(status, '')) = 'aberto' THEN 1 ELSE 0 END), 0),
                COALESCE(SUM(CASE WHEN LOWER(COALESCE(base, '')) LIKE '%otrs%' THEN 1 ELSE 0 END), 0),
                COALESCE(SUM(CASE WHEN LOWER(COALESCE(base, '')) LIKE '%citsmart%' THEN 1 ELSE 0 END), 0)
            FROM chamados
            """
        )
        sum_row = cursor.fetchone()
        if sum_row:
            summary = {
                "total": int(sum_row[0] or 0),
                "abertos": int(sum_row[1] or 0),
                "otrs_count": int(sum_row[2] or 0),
                "citsmart_count": int(sum_row[3] or 0),
            }

        where_sql, params = _build_chamados_where(
            search=search,
            base_filter=base_filter,
            status_filter=status_filter,
            tag_filter=tag_filter,
            localidades_filter=localidades_filter,
        )
        query_sql = f"""
            SELECT
                COALESCE(id, ''),
                COALESCE(data_criacao, ''),
                COALESCE(titulo, ''),
                COALESCE(cidade_predio, ''),
                COALESCE(unidade, ''),
                COALESCE(localidade_fisica, ''),
                COALESCE(usuario, ''),
                COALESCE(id_cliente, ''),
                COALESCE(descricao, ''),
                COALESCE(tag, ''),
                COALESCE(ip_origem, ''),
                COALESCE(status, 'Aberto'),
                COALESCE(data_atualizacao, ''),
                COALESCE(base, ''),
                COALESCE(andamento, ''),
                COALESCE(link, ''),
                COALESCE(hostname, ''),
                COALESCE(tag_manual, 0),
                COALESCE(dados_manuais, 0)
            FROM chamados
            {where_sql}
            ORDER BY COALESCE(NULLIF(data_atualizacao, ''), data_criacao) DESC
            LIMIT ? OFFSET ?
        """
        cursor.execute(query_sql, tuple(params + [safe_limit, safe_offset]))
        for r in cursor.fetchall():
            items.append(
                {
                    "id": _clean_str(r[0]),
                    "data_criacao": _clean_str(r[1]),
                    "titulo": _clean_str(r[2]),
                    "cidade_predio": _clean_str(r[3]),
                    "unidade": _clean_str(r[4]),
                    "localidade_fisica": _clean_str(r[5]),
                    "usuario": _clean_str(r[6]),
                    "id_cliente": _clean_str(r[7]),
                    "descricao": _clean_str(r[8]),
                    "tag": _clean_str(r[9]),
                    "ip_origem": _clean_str(r[10]),
                    "status": _clean_str(r[11], "Aberto"),
                    "data_atualizacao": _clean_str(r[12]),
                    "base": _clean_str(r[13]),
                    "andamento": _clean_str(r[14]),
                    "link": _clean_str(r[15]),
                    "hostname": _clean_str(r[16]),
                    "tag_manual": bool(r[17]),
                    "dados_manuais": bool(r[18]),
                }
            )

    return {"items": items, "summary": summary, "count": len(items)}


def get_chamado_by_id(chamado_id: str) -> Optional[Dict[str, Any]]:
    cid = _normalize_ticket_id(chamado_id)
    if not cid:
        return None

    with _db_cursor() as (_, cursor):
        cursor.execute(
            """
            SELECT
                COALESCE(id, ''),
                COALESCE(data_criacao, ''),
                COALESCE(titulo, ''),
                COALESCE(cidade_predio, ''),
                COALESCE(unidade, ''),
                COALESCE(localidade_fisica, ''),
                COALESCE(usuario, ''),
                COALESCE(id_cliente, ''),
                COALESCE(descricao, ''),
                COALESCE(tag, ''),
                COALESCE(ip_origem, ''),
                COALESCE(status, 'Aberto'),
                COALESCE(data_atualizacao, ''),
                COALESCE(base, ''),
                COALESCE(andamento, ''),
                COALESCE(link, ''),
                COALESCE(hostname, ''),
                COALESCE(tag_manual, 0),
                COALESCE(dados_manuais, 0)
            FROM chamados
            WHERE id = ?
            """,
            (cid,),
        )
        r = cursor.fetchone()
        if not r:
            return None
        return {
            "id": _clean_str(r[0]),
            "data_criacao": _clean_str(r[1]),
            "titulo": _clean_str(r[2]),
            "cidade_predio": _clean_str(r[3]),
            "unidade": _clean_str(r[4]),
            "localidade_fisica": _clean_str(r[5]),
            "usuario": _clean_str(r[6]),
            "id_cliente": _clean_str(r[7]),
            "descricao": _clean_str(r[8]),
            "tag": _clean_str(r[9]),
            "ip_origem": _clean_str(r[10]),
            "status": _clean_str(r[11], "Aberto"),
            "data_atualizacao": _clean_str(r[12]),
            "base": _clean_str(r[13]),
            "andamento": _clean_str(r[14]),
            "link": _clean_str(r[15]),
            "hostname": _clean_str(r[16]),
            "tag_manual": bool(r[17]),
            "dados_manuais": bool(r[18]),
        }


def update_chamado(chamado_id: str, payload: ChamadoUpdatePayload) -> Optional[Dict[str, Any]]:
    cid = _normalize_ticket_id(chamado_id)
    if not cid:
        return None

    data = payload.model_dump(exclude_none=True)
    now = datetime.now().strftime("%Y-%m-%d %H:%M:%S")

    field_map = {
        "status": "status",
        "titulo": "titulo",
        "hostname": "hostname",
        "ip_origem": "ip_origem",
        "tag": "tag",
        "cidade_predio": "cidade_predio",
        "unidade": "unidade",
        "localidade_fisica": "localidade_fisica",
        "descricao": "descricao",
        "andamento": "andamento",
    }

    set_clauses: List[str] = []
    params: List[Any] = []

    for py_field, db_col in field_map.items():
        if py_field in data:
            set_clauses.append(f"{db_col} = ?")
            params.append(str(data[py_field]).strip())

    if "tag_manual" in data:
        set_clauses.append("tag_manual = ?")
        params.append(1 if data["tag_manual"] else 0)
    elif "tag" in data:
        set_clauses.append("tag_manual = 1")

    if "dados_manuais" in data:
        set_clauses.append("dados_manuais = ?")
        params.append(1 if data["dados_manuais"] else 0)
    elif any(k in data for k in ("cidade_predio", "unidade", "localidade_fisica")):
        set_clauses.append("dados_manuais = 1")

    if not set_clauses:
        return get_chamado_by_id(cid)

    set_clauses.append("data_atualizacao = ?")
    params.append(now)
    params.append(cid)

    with _db_cursor() as (conn, cursor):
        cursor.execute(
            f"UPDATE chamados SET {', '.join(set_clauses)} WHERE id = ?",
            tuple(params),
        )
        updated = cursor.rowcount
        conn.commit()
        if updated == 0:
            # Verifica se o registro realmente não existe (caso rowcount varie por driver)
            cursor.execute("SELECT 1 FROM chamados WHERE id = ?", (cid,))
            if not cursor.fetchone():
                return None

    return get_chamado_by_id(cid)


def get_comentarios_chamado(chamado_id: str) -> List[Dict[str, str]]:
    cid = _normalize_ticket_id(chamado_id)
    if not cid:
        return []

    with _db_cursor() as (_, cursor):
        cursor.execute(
            """
            SELECT COALESCE(data, ''), COALESCE(autor, ''), COALESCE(texto, '')
            FROM comentarios
            WHERE chamado_id = ?
            ORDER BY id ASC
            """,
            (cid,),
        )
        return [
            {"data": _clean_str(r[0]), "autor": _clean_str(r[1]), "texto": _clean_str(r[2])}
            for r in cursor.fetchall()
        ]


def get_comentarios_batch(chamado_ids: Iterable[str]) -> Dict[str, List[Dict[str, str]]]:
    clean_ids = list(dict.fromkeys(_normalize_ticket_id(cid) for cid in chamado_ids if _normalize_ticket_id(cid)))[:500]
    result: Dict[str, List[Dict[str, str]]] = {}
    if not clean_ids:
        return result

    placeholders = ", ".join(["?"] * len(clean_ids))
    with _db_cursor() as (_, cursor):
        cursor.execute(
            f"""
            SELECT COALESCE(chamado_id, ''), COALESCE(data, ''), COALESCE(autor, ''), COALESCE(texto, '')
            FROM comentarios
            WHERE chamado_id IN ({placeholders})
            ORDER BY id ASC
            """,
            tuple(clean_ids),
        )
        for r in cursor.fetchall():
            cid = _clean_str(r[0])
            result.setdefault(cid, []).append(
                {"data": _clean_str(r[1]), "autor": _clean_str(r[2]), "texto": _clean_str(r[3])}
            )
    return result


def add_comentario_chamado(chamado_id: str, payload: ComentarioCreatePayload) -> Optional[Dict[str, str]]:
    cid = _normalize_ticket_id(chamado_id)
    if not cid:
        return None

    now = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    data_str = _clean_str(payload.data, now) or now

    with _db_cursor() as (conn, cursor):
        cursor.execute("SELECT 1 FROM chamados WHERE id = ?", (cid,))
        if not cursor.fetchone():
            return None

        cursor.execute(
            """
            INSERT INTO comentarios (chamado_id, data, autor, texto)
            VALUES (?, ?, ?, ?)
            """,
            (cid, data_str, payload.autor.strip(), payload.texto.strip()),
        )
        cursor.execute(
            "UPDATE chamados SET data_atualizacao = ? WHERE id = ?",
            (now, cid),
        )
        conn.commit()

    return {"chamado_id": cid, "data": data_str, "autor": payload.autor.strip(), "texto": payload.texto.strip()}


def query_cron_data() -> Dict[str, Any]:
    schedules: List[Dict[str, Any]] = []
    logs: List[Dict[str, Any]] = []

    with _db_cursor() as (_, cursor):
        cursor.execute(
            """
            SELECT
                COALESCE(task_id, ''),
                COALESCE(nome, ''),
                COALESCE(categoria, ''),
                COALESCE(ativo, 1),
                COALESCE(tipo_agendamento, 'intervalo'),
                COALESCE(intervalo_valor, 2),
                COALESCE(intervalo_unidade, 'horas'),
                COALESCE(horario_fixo, '12:00'),
                COALESCE(apenas_dias_uteis, 0),
                COALESCE(descricao, ''),
                ultima_execucao,
                COALESCE(ultimo_status, 'pendente'),
                COALESCE(ultimo_log, ''),
                COALESCE(run_requested, 0)
            FROM cron_schedules
            ORDER BY COALESCE(categoria, '') ASC, COALESCE(nome, '') ASC
            """
        )
        for r in cursor.fetchall():
            schedules.append(
                {
                    "task_id": _clean_str(r[0]),
                    "nome": _clean_str(r[1]),
                    "categoria": _clean_str(r[2]),
                    "ativo": int(r[3] or 0) == 1,
                    "tipo_agendamento": _clean_str(r[4], "intervalo"),
                    "intervalo_valor": int(r[5] or 2),
                    "intervalo_unidade": _clean_str(r[6], "horas"),
                    "horario_fixo": _clean_str(r[7], "12:00"),
                    "apenas_dias_uteis": int(r[8] or 0) == 1,
                    "descricao": _clean_str(r[9]),
                    "ultima_execucao": _clean_str(r[10]),
                    "ultimo_status": _clean_str(r[11], "pendente"),
                    "ultimo_log": _clean_str(r[12]),
                    "run_requested": int(r[13] or 0) == 1,
                }
            )

        cursor.execute(
            """
            SELECT
                l.id,
                COALESCE(s.nome, l.task_id),
                l.inicio,
                COALESCE(l.duracao_segundos, 0),
                COALESCE(l.status, ''),
                COALESCE(l.mensagem, '')
            FROM cron_logs l
            LEFT JOIN cron_schedules s ON l.task_id = s.task_id
            ORDER BY l.id DESC
            LIMIT 25
            """
        )
        for r in cursor.fetchall():
            logs.append(
                {
                    "id": int(r[0]),
                    "rotina": _clean_str(r[1]),
                    "inicio": _clean_str(r[2]),
                    "duracao_segundos": str(r[3] if r[3] is not None else 0),
                    "status": _clean_str(r[4]),
                    "mensagem": _clean_str(r[5]),
                }
            )

    summary = {
        "total": len(schedules),
        "ativas": sum(1 for s in schedules if s["ativo"]),
        "executando": sum(
            1 for s in schedules if s["ultimo_status"].lower() == "executando" or s["run_requested"]
        ),
        "com_erro": sum(1 for s in schedules if s["ultimo_status"].lower() == "erro"),
    }

    return {"schedules": schedules, "logs": logs, "summary": summary}


def trigger_cron_task(task_id: str) -> Tuple[bool, str, int]:
    clean_id = (task_id or "").strip()
    if not clean_id:
        return False, "ID da tarefa inválido.", 400

    with _db_cursor() as (conn, cursor):
        cursor.execute("SELECT ultimo_status FROM cron_schedules WHERE task_id = ?", (clean_id,))
        row = cursor.fetchone()
        if not row:
            return False, f"Rotina '{clean_id}' não encontrada.", 404

        current_status = _clean_str(row[0]).lower()
        if current_status == "executando":
            return False, f"A tarefa '{clean_id}' já está em execução.", 409

        cursor.execute(
            """
            UPDATE cron_schedules
            SET run_requested = 1,
                updated_at = CURRENT_TIMESTAMP
            WHERE task_id = ?
            """,
            (clean_id,),
        )
        conn.commit()

    return True, f"Execução solicitada para '{clean_id}'. O container iniciará em segundos.", 200


def update_cron_task(task_id: str, payload: CronUpdatePayload) -> Tuple[bool, str, int]:
    clean_id = (task_id or "").strip()
    if not clean_id:
        return False, "ID da tarefa inválido.", 400

    data = payload.model_dump(exclude_none=True)
    if not data:
        return False, "Nenhum campo informado para atualização.", 400

    set_clauses: List[str] = []
    params: List[Any] = []
    for col in ("tipo_agendamento", "intervalo_valor", "intervalo_unidade", "horario_fixo"):
        if col in data:
            set_clauses.append(f"{col} = ?")
            params.append(data[col])
    if "ativo" in data:
        set_clauses.append("ativo = ?")
        params.append(1 if data["ativo"] else 0)
    if "apenas_dias_uteis" in data:
        set_clauses.append("apenas_dias_uteis = ?")
        params.append(1 if data["apenas_dias_uteis"] else 0)

    set_clauses.append("updated_at = CURRENT_TIMESTAMP")
    params.append(clean_id)

    with _db_cursor() as (conn, cursor):
        cursor.execute("SELECT 1 FROM cron_schedules WHERE task_id = ?", (clean_id,))
        if not cursor.fetchone():
            return False, f"Rotina '{clean_id}' não encontrada.", 404

        cursor.execute(
            f"UPDATE cron_schedules SET {', '.join(set_clauses)} WHERE task_id = ?",
            tuple(params),
        )
        conn.commit()

    return True, f"Rotina '{clean_id}' atualizada com sucesso.", 200


def query_oxe_ramais(search: Optional[str] = None, limit: int = 2000) -> Dict[str, Any]:
    safe_limit = max(1, min(int(limit), 5000))
    all_items: List[Dict[str, Any]] = []

    with _db_cursor() as (_, cursor):
        cursor.execute("SELECT * FROM central_telefonica")
        col_names = [desc[0] for desc in cursor.description]
        for row in cursor.fetchall():
            fields = {col_names[i]: _clean_str(row[i]) for i in range(len(col_names))}
            item = {
                "ramal": _first_non_empty(fields, "ramal", "directory_number", "numero", "extensao"),
                "nome": _first_non_empty(fields, "nome", "usuario", "titular", "display_name", "nome_usuario", "name"),
                "localidade": _first_non_empty(fields, "localidade", "entidade", "unidade", "cidade", "predio"),
                "setor": _first_non_empty(fields, "setor", "departamento", "lotacao", "department"),
                "tipo": _first_non_empty(fields, "tipo", "aparelho", "modelo", "set_type", "station_type"),
                "ip": _first_non_empty(fields, "ip", "endereco_ip", "ip_address"),
                "mac": _first_non_empty(fields, "mac", "mac_address", "endereco_mac"),
                "grupo_captura": _first_non_empty(fields, "grupo_de_captura", "grupo_captura", "pickup_group"),
                "cat_rede_publica": _first_non_empty(fields, "cat_rede_pública", "cat_rede_publica", "categoria_rede_publica", "public_network_category"),
                "categoria_dispositivo": _first_non_empty(fields, "categoria_dispositivo", "categoria_do_dispositivo", "categoria"),
                "tipo_estacao": _first_non_empty(fields, "tipo_de_estacao", "tipo_estacao", "station_type", "tipo"),
                "all_fields": fields,
            }
            all_items.append(item)

    all_items.sort(key=lambda x: (0, int(x["ramal"])) if x["ramal"].isdigit() else (1, x["ramal"]))

    summary = {
        "total": len(all_items),
        "com_nome": sum(1 for x in all_items if _has_real_value(x["nome"])),
        "com_ip_ou_mac": sum(1 for x in all_items if _has_real_value(x["ip"]) or _has_real_value(x["mac"])),
    }

    q = (search or "").strip().lower()
    if q:
        filtered = [
            item
            for item in all_items
            if any(q in v.lower() for v in item["all_fields"].values())
        ]
    else:
        filtered = all_items

    return {"items": filtered[:safe_limit], "summary": summary, "count": len(filtered[:safe_limit])}


def _is_printer_status_ok(status: str) -> bool:
    s = (status or "").strip().lower()
    return s == "ok" or "online" in s or "pront" in s or "normal" in s


def query_impressoras(search: Optional[str] = None) -> Dict[str, Any]:
    items: List[Dict[str, Any]] = []
    with _db_cursor() as (_, cursor):
        cursor.execute("SELECT * FROM impressoras ORDER BY COALESCE(nome, '') ASC")
        col_names = [desc[0] for desc in cursor.description]
        for row in cursor.fetchall():
            fields = {col_names[i]: _clean_str(row[i]) for i in range(len(col_names))}
            raw_pages = _first_non_empty(fields, "total_paginas")
            try:
                total_paginas = int(float(raw_pages)) if raw_pages else 0
            except ValueError:
                total_paginas = 0

            items.append(
                {
                    "id": _first_non_empty(fields, "id"),
                    "nome": _first_non_empty(fields, "nome"),
                    "servidor": _first_non_empty(fields, "servidor"),
                    "tipo": _first_non_empty(fields, "tipo"),
                    "modelo": _first_non_empty(fields, "modelo"),
                    "localizacao": _first_non_empty(fields, "localizacao"),
                    "ip_host": _first_non_empty(fields, "ip_host"),
                    "status": _first_non_empty(fields, "status"),
                    "total_paginas": total_paginas,
                    "filas_relacionadas": _first_non_empty(fields, "filas_relacionadas"),
                    "detalhes_extra": _first_non_empty(fields, "detalhes_extra"),
                    "data_atualizacao": _first_non_empty(fields, "data_atualizacao"),
                    "all_fields": fields,
                }
            )

    summary = {
        "total": len(items),
        "ok_count": sum(1 for x in items if _is_printer_status_ok(x["status"])),
        "alert_count": sum(1 for x in items if not _is_printer_status_ok(x["status"])),
        "total_paginas_soma": sum(x["total_paginas"] for x in items),
    }

    q = (search or "").strip().lower()
    if q:
        items = [x for x in items if any(q in v.lower() for v in x["all_fields"].values())]

    return {"items": items, "summary": summary, "count": len(items)}


def query_ramais_mpms(search: Optional[str] = None, limit: int = 500) -> Dict[str, Any]:
    safe_limit = max(1, min(int(limit), 2000))
    items: List[Dict[str, Any]] = []
    with _db_cursor() as (_, cursor):
        if search and search.strip():
            q = f"%{search.strip().lower()}%"
            cursor.execute(
                """
                SELECT id, COALESCE(localidade, ''), COALESCE(setor_nome, ''), COALESCE(telefone_ramal, ''), COALESCE(tipo, ''), COALESCE(data_atualizacao, '')
                FROM ramais_mpms
                WHERE LOWER(COALESCE(localidade, '')) LIKE ?
                   OR LOWER(COALESCE(setor_nome, '')) LIKE ?
                   OR LOWER(COALESCE(telefone_ramal, '')) LIKE ?
                ORDER BY localidade ASC, setor_nome ASC
                LIMIT ?
                """,
                (q, q, q, safe_limit),
            )
        else:
            cursor.execute(
                """
                SELECT id, COALESCE(localidade, ''), COALESCE(setor_nome, ''), COALESCE(telefone_ramal, ''), COALESCE(tipo, ''), COALESCE(data_atualizacao, '')
                FROM ramais_mpms
                ORDER BY localidade ASC, setor_nome ASC
                LIMIT ?
                """,
                (safe_limit,),
            )
        for r in cursor.fetchall():
            items.append(
                {
                    "id": int(r[0]),
                    "localidade": _clean_str(r[1]),
                    "setor_nome": _clean_str(r[2]),
                    "telefone_ramal": _clean_str(r[3]),
                    "tipo": _clean_str(r[4]),
                    "data_atualizacao": _clean_str(r[5]),
                }
            )
    return {"items": items, "count": len(items)}


def query_unidades_mpms(search: Optional[str] = None, limit: int = 500) -> Dict[str, Any]:
    safe_limit = max(1, min(int(limit), 2000))
    items: List[Dict[str, Any]] = []
    with _db_cursor() as (_, cursor):
        if search and search.strip():
            q = f"%{search.strip().lower()}%"
            cursor.execute(
                """
                SELECT id, COALESCE(cidade, ''), COALESCE(tipo, ''), COALESCE(setor, ''), COALESCE(sigla, ''),
                       COALESCE(titular, ''), COALESCE(unidade_predio, ''), COALESCE(telefone, ''), COALESCE(url, ''), COALESCE(origem, '')
                FROM unidades
                WHERE LOWER(COALESCE(cidade, '')) LIKE ?
                   OR LOWER(COALESCE(setor, '')) LIKE ?
                   OR LOWER(COALESCE(sigla, '')) LIKE ?
                   OR LOWER(COALESCE(unidade_predio, '')) LIKE ?
                ORDER BY cidade ASC, setor ASC
                LIMIT ?
                """,
                (q, q, q, q, safe_limit),
            )
        else:
            cursor.execute(
                """
                SELECT id, COALESCE(cidade, ''), COALESCE(tipo, ''), COALESCE(setor, ''), COALESCE(sigla, ''),
                       COALESCE(titular, ''), COALESCE(unidade_predio, ''), COALESCE(telefone, ''), COALESCE(url, ''), COALESCE(origem, '')
                FROM unidades
                ORDER BY cidade ASC, setor ASC
                LIMIT ?
                """,
                (safe_limit,),
            )
        for r in cursor.fetchall():
            items.append(
                {
                    "id": int(r[0]),
                    "cidade": _clean_str(r[1]),
                    "tipo": _clean_str(r[2]),
                    "setor": _clean_str(r[3]),
                    "sigla": _clean_str(r[4]),
                    "titular": _clean_str(r[5]),
                    "unidade_predio": _clean_str(r[6]),
                    "telefone": _clean_str(r[7]),
                    "url": _clean_str(r[8]),
                    "origem": _clean_str(r[9]),
                }
            )
    return {"items": items, "count": len(items)}


def query_sccm_devices(search: Optional[str] = None, limit: int = 200) -> Dict[str, Any]:
    safe_limit = max(1, min(int(limit), 1000))
    items: List[Dict[str, Any]] = []
    with _db_cursor() as (_, cursor):
        if search and search.strip():
            q = f"%{search.strip().lower()}%"
            cursor.execute(
                """
                SELECT COALESCE(resource_id, ''), COALESCE(name, ''), COALESCE(last_logon_user, ''),
                       COALESCE(ip_addresses, ''), COALESCE(mac_addresses, ''), COALESCE(manufacturer, ''),
                       COALESCE(model, ''), COALESCE(operating_system, ''), COALESCE(os_version, ''),
                       COALESCE(client_active, 1), COALESCE(last_active_time, ''), COALESCE(ad_site_name, ''),
                       COALESCE(processor, ''), COALESCE(memory_ram, ''), COALESCE(updated_at, '')
                FROM sccm_cache_devices
                WHERE LOWER(COALESCE(name, '')) LIKE ?
                   OR LOWER(COALESCE(last_logon_user, '')) LIKE ?
                   OR LOWER(COALESCE(ip_addresses, '')) LIKE ?
                   OR LOWER(COALESCE(model, '')) LIKE ?
                ORDER BY name ASC
                LIMIT ?
                """,
                (q, q, q, q, safe_limit),
            )
        else:
            cursor.execute(
                """
                SELECT COALESCE(resource_id, ''), COALESCE(name, ''), COALESCE(last_logon_user, ''),
                       COALESCE(ip_addresses, ''), COALESCE(mac_addresses, ''), COALESCE(manufacturer, ''),
                       COALESCE(model, ''), COALESCE(operating_system, ''), COALESCE(os_version, ''),
                       COALESCE(client_active, 1), COALESCE(last_active_time, ''), COALESCE(ad_site_name, ''),
                       COALESCE(processor, ''), COALESCE(memory_ram, ''), COALESCE(updated_at, '')
                FROM sccm_cache_devices
                ORDER BY name ASC
                LIMIT ?
                """,
                (safe_limit,),
            )
        for r in cursor.fetchall():
            items.append(
                {
                    "resource_id": _clean_str(r[0]),
                    "name": _clean_str(r[1]),
                    "last_logon_user": _clean_str(r[2]),
                    "ip_addresses": _clean_str(r[3]),
                    "mac_addresses": _clean_str(r[4]),
                    "manufacturer": _clean_str(r[5]),
                    "model": _clean_str(r[6]),
                    "operating_system": _clean_str(r[7]),
                    "os_version": _clean_str(r[8]),
                    "client_active": int(r[9] or 0) == 1,
                    "last_active_time": _clean_str(r[10]),
                    "ad_site_name": _clean_str(r[11]),
                    "processor": _clean_str(r[12]),
                    "memory_ram": _clean_str(r[13]),
                    "updated_at": _clean_str(r[14]),
                }
            )
    return {"items": items, "count": len(items)}
