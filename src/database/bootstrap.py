# -*- coding: utf-8 -*-
"""
Bootstrap Unificado do Banco de Dados (PostgreSQL & SQLite).
Garante que todas as tabelas, configurações padrão, rotinas de cron e índices
sejam criados de forma idempotente na primeira inicialização em qualquer nova máquina.
"""

import logging
import time
from typing import Dict, List

from .connection import ensure_database_indexes, get_connection

logger = logging.getLogger(__name__)


def wait_for_database(max_retries: int = 10, delay_seconds: float = 2.0) -> bool:
    """Aguarda o banco de dados estar acessível para conexões (útil em cold start no Docker)."""
    for attempt in range(1, max_retries + 1):
        try:
            conn = get_connection()
            cur = conn.cursor()
            cur.execute("SELECT 1")
            cur.fetchone()
            cur.close()
            conn.close()
            return True
        except Exception as exc:
            if attempt == max_retries:
                logger.error(f"Banco de dados indisponível após {max_retries} tentativas: {exc}")
                return False
            time.sleep(delay_seconds)
    return False


def setup_faq_tables() -> None:
    """Garante a criação das tabelas da base de conhecimento (faqs, faq_videos, faq_imagens)."""
    conn = get_connection()
    cursor = conn.cursor()
    try:
        cursor.execute("""
            CREATE TABLE IF NOT EXISTS faqs (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                titulo TEXT NOT NULL,
                tipo_faq TEXT NOT NULL,
                url TEXT NOT NULL UNIQUE,
                conteudo TEXT,
                data_atualizacao TIMESTAMP DEFAULT CURRENT_TIMESTAMP
            )
        """)
        cursor.execute("""
            CREATE TABLE IF NOT EXISTS faq_videos (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                titulo TEXT NOT NULL,
                categoria TEXT NOT NULL,
                nome_arquivo TEXT NOT NULL,
                url TEXT NOT NULL UNIQUE,
                caminho_relativo TEXT,
                tamanho_bytes INTEGER DEFAULT 0,
                data_atualizacao TIMESTAMP DEFAULT CURRENT_TIMESTAMP
            )
        """)
        cursor.execute("""
            CREATE TABLE IF NOT EXISTS faq_imagens (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                titulo TEXT NOT NULL,
                categoria TEXT NOT NULL,
                nome_arquivo TEXT NOT NULL,
                url TEXT NOT NULL UNIQUE,
                caminho_relativo TEXT,
                tamanho_bytes INTEGER DEFAULT 0,
                data_atualizacao TIMESTAMP DEFAULT CURRENT_TIMESTAMP
            )
        """)
        conn.commit()
    finally:
        cursor.close()
        conn.close()


def initialize_database(wait_retries: int = 10, delay_seconds: float = 2.0) -> Dict[str, List[str]]:
    """
    Inicializa todas as tabelas, configurações iniciais, agendamentos de cron e índices
    do Sistema Bancada STI de forma idempotente.
    """
    summary: Dict[str, List[str]] = {"initialized": [], "errors": []}

    if not wait_for_database(max_retries=wait_retries, delay_seconds=delay_seconds):
        summary["errors"].append("database_connection_unavailable")
        return summary

    from .ad_db import setup_ad_tables
    from .cron_db import setup_cron_tables
    from .donations_db import setup_donations_table
    from .events_db import setup_eventos_manuais_table
    from .ferias_db import setup_ferias_table
    from .garantia_db import setup_garantia_tables
    from .impressoras_db import setup_impressoras_table
    from .map_db import setup_map_tables
    from .notifications_db import setup_notifications_table
    from .plantoes_db import setup_plantoes_tables
    from .ramais_db import setup_ramais_config_table, setup_ramais_table
    from .sccm_db import setup_sccm_tables
    from .senhas_db import setup_senhas_table
    from .settings_db import seed_settings_from_env_if_empty, setup_settings_table
    from .tickets_db import setup_database
    from .unidades_db import setup_unidades_tables
    from .viagens_db import setup_viagens_table
    from .whatsapp_db import setup_whatsapp_tables

    steps = [
        ("system_settings", setup_settings_table),
        ("seed_settings", seed_settings_from_env_if_empty),
        ("cron_schedules_and_logs", setup_cron_tables),
        ("chamados_and_comentarios", setup_database),
        ("unidades", setup_unidades_tables),
        ("ramais", setup_ramais_table),
        ("ramais_config", setup_ramais_config_table),
        ("plantoes", setup_plantoes_tables),
        ("ferias", setup_ferias_table),
        ("viagens", setup_viagens_table),
        ("garantia", setup_garantia_tables),
        ("equipamentos_doados", setup_donations_table),
        ("impressoras", setup_impressoras_table),
        ("eventos_manuais", setup_eventos_manuais_table),
        ("mapa", setup_map_tables),
        ("notificacoes", setup_notifications_table),
        ("senhas_cofre", setup_senhas_table),
        ("whatsapp", setup_whatsapp_tables),
        ("ad_cache", setup_ad_tables),
        ("sccm_cache", setup_sccm_tables),
        ("faqs", setup_faq_tables),
        ("database_indexes", ensure_database_indexes),
    ]

    for name, fn in steps:
        try:
            fn()
            summary["initialized"].append(name)
        except Exception as exc:
            logger.warning(f"Falha ao inicializar módulo '{name}' no banco: {exc}")
            summary["errors"].append(f"{name}: {exc}")

    return summary

