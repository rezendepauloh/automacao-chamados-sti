import os
from datetime import datetime
import pandas as pd
from typing import List, Dict, Any, Optional
from .connection import get_connection, DB_TYPE

def setup_ad_tables():
    """
    Cria as tabelas de cache do Active Directory caso não existam.
    Suporta PostgreSQL e SQLite.
    """
    conn = get_connection()
    cursor = conn.cursor()

    is_pg = DB_TYPE in ["postgres", "postgresql"]
    id_pk = "SERIAL PRIMARY KEY" if is_pg else "INTEGER PRIMARY KEY AUTOINCREMENT"

    # 1. Tabela de Unidades Organizacionais (OUs)
    cursor.execute(f"""
    CREATE TABLE IF NOT EXISTS ad_cache_ous (
        dn TEXT PRIMARY KEY,
        name TEXT,
        parent_dn TEXT,
        description TEXT,
        canonical_name TEXT,
        updated_at TEXT
    )
    """)

    # 2. Tabela de Usuários / Contas
    cursor.execute(f"""
    CREATE TABLE IF NOT EXISTS ad_cache_users (
        sam_account_name TEXT PRIMARY KEY,
        display_name TEXT,
        mail TEXT,
        department TEXT,
        title TEXT,
        dn TEXT,
        parent_ou_dn TEXT,
        is_active BOOLEAN,
        user_account_control INTEGER,
        when_created TEXT,
        last_logon TEXT,
        updated_at TEXT
    )
    """)

    # 3. Tabela de Grupos de Segurança
    cursor.execute(f"""
    CREATE TABLE IF NOT EXISTS ad_cache_groups (
        sam_account_name TEXT PRIMARY KEY,
        display_name TEXT,
        dn TEXT,
        description TEXT,
        group_type TEXT,
        member_count INTEGER DEFAULT 0,
        updated_at TEXT
    )
    """)

    # 4. Tabela de Relação Grupo <-> Membro
    cursor.execute(f"""
    CREATE TABLE IF NOT EXISTS ad_cache_group_members (
        id {id_pk},
        group_dn TEXT,
        member_dn TEXT,
        updated_at TEXT
    )
    """)

    # 5. Metadados de Sincronização
    cursor.execute(f"""
    CREATE TABLE IF NOT EXISTS ad_sync_meta (
        id INTEGER PRIMARY KEY,
        last_sync_at TEXT,
        status TEXT,
        total_ous INTEGER DEFAULT 0,
        total_users INTEGER DEFAULT 0,
        total_groups INTEGER DEFAULT 0,
        error_message TEXT
    )
    """)

    conn.commit()
    conn.close()


def save_ad_cache(
    ous: List[Dict[str, Any]],
    users: List[Dict[str, Any]],
    groups: List[Dict[str, Any]],
    memberships: List[Dict[str, str]],
    status: str = "success",
    error_message: Optional[str] = None
) -> bool:
    """
    Sobrescreve/atualiza os dados de cache do Active Directory de forma transacional.
    """
    setup_ad_tables()
    conn = get_connection()
    cursor = conn.cursor()
    now_str = datetime.now().strftime("%Y-%m-%d %H:%M:%S")

    try:
        # Limpa tabelas de cache para consistência total da árvore
        cursor.execute("DELETE FROM ad_cache_ous")
        cursor.execute("DELETE FROM ad_cache_users")
        cursor.execute("DELETE FROM ad_cache_groups")
        cursor.execute("DELETE FROM ad_cache_group_members")

        # Inserção de OUs
        for ou in ous:
            cursor.execute("""
                INSERT INTO ad_cache_ous (dn, name, parent_dn, description, canonical_name, updated_at)
                VALUES (%s, %s, %s, %s, %s, %s)
            """ if DB_TYPE in ["postgres", "postgresql"] else """
                INSERT INTO ad_cache_ous (dn, name, parent_dn, description, canonical_name, updated_at)
                VALUES (?, ?, ?, ?, ?, ?)
            """, (
                ou.get("dn", ""),
                ou.get("name", ""),
                ou.get("parent_dn", ""),
                ou.get("description", ""),
                ou.get("canonical_name", ""),
                now_str
            ))

        # Inserção de Usuários
        for u in users:
            cursor.execute("""
                INSERT INTO ad_cache_users (
                    sam_account_name, display_name, mail, department, title, dn,
                    parent_ou_dn, is_active, user_account_control, when_created, last_logon, updated_at
                )
                VALUES (%s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s)
            """ if DB_TYPE in ["postgres", "postgresql"] else """
                INSERT INTO ad_cache_users (
                    sam_account_name, display_name, mail, department, title, dn,
                    parent_ou_dn, is_active, user_account_control, when_created, last_logon, updated_at
                )
                VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
            """, (
                u.get("sam_account_name", ""),
                u.get("display_name", ""),
                u.get("mail", ""),
                u.get("department", ""),
                u.get("title", ""),
                u.get("dn", ""),
                u.get("parent_ou_dn", ""),
                bool(u.get("is_active", True)),
                int(u.get("user_account_control", 512)),
                str(u.get("when_created", "")),
                str(u.get("last_logon", "")),
                now_str
            ))

        # Inserção de Grupos
        for g in groups:
            cursor.execute("""
                INSERT INTO ad_cache_groups (sam_account_name, display_name, dn, description, group_type, member_count, updated_at)
                VALUES (%s, %s, %s, %s, %s, %s, %s)
            """ if DB_TYPE in ["postgres", "postgresql"] else """
                INSERT INTO ad_cache_groups (sam_account_name, display_name, dn, description, group_type, member_count, updated_at)
                VALUES (?, ?, ?, ?, ?, ?, ?)
            """, (
                g.get("sam_account_name", ""),
                g.get("display_name", ""),
                g.get("dn", ""),
                g.get("description", ""),
                g.get("group_type", ""),
                int(g.get("member_count", 0)),
                now_str
            ))

        # Inserção de Associações Grupo <-> Membros
        for m in memberships:
            cursor.execute("""
                INSERT INTO ad_cache_group_members (group_dn, member_dn, updated_at)
                VALUES (%s, %s, %s)
            """ if DB_TYPE in ["postgres", "postgresql"] else """
                INSERT INTO ad_cache_group_members (group_dn, member_dn, updated_at)
                VALUES (?, ?, ?)
            """, (
                m.get("group_dn", ""),
                m.get("member_dn", ""),
                now_str
            ))

        # Atualiza Metadados de Sincronização (ID fixo = 1)
        cursor.execute("DELETE FROM ad_sync_meta WHERE id = 1")
        cursor.execute("""
            INSERT INTO ad_sync_meta (id, last_sync_at, status, total_ous, total_users, total_groups, error_message)
            VALUES (%s, %s, %s, %s, %s, %s, %s)
        """ if DB_TYPE in ["postgres", "postgresql"] else """
            INSERT INTO ad_sync_meta (id, last_sync_at, status, total_ous, total_users, total_groups, error_message)
            VALUES (?, ?, ?, ?, ?, ?, ?)
        """, (
            1,
            now_str,
            status,
            len(ous),
            len(users),
            len(groups),
            error_message or ""
        ))

        conn.commit()
        return True
    except Exception as e:
        conn.rollback()
        raise e
    finally:
        conn.close()


def get_ad_sync_meta() -> Dict[str, Any]:
    """Retorna os metadados da última sincronização do Active Directory."""
    setup_ad_tables()
    conn = get_connection()
    cursor = conn.cursor()
    try:
        cursor.execute("SELECT last_sync_at, status, total_ous, total_users, total_groups, error_message FROM ad_sync_meta WHERE id = 1")
        row = cursor.fetchone()
        if row:
            return {
                "last_sync_at": row[0],
                "status": row[1],
                "total_ous": row[2],
                "total_users": row[3],
                "total_groups": row[4],
                "error_message": row[5]
            }
        return {
            "last_sync_at": None,
            "status": "never_run",
            "total_ous": 0,
            "total_users": 0,
            "total_groups": 0,
            "error_message": ""
        }
    finally:
        conn.close()


def get_ad_ous() -> pd.DataFrame:
    """Retorna DataFrame de todas as OUs salvas no cache."""
    setup_ad_tables()
    conn = get_connection()
    try:
        return pd.read_sql_query("SELECT dn, name, parent_dn, description, canonical_name FROM ad_cache_ous ORDER BY name ASC", conn)
    finally:
        conn.close()


def get_ad_users_df(status_filter: str = "Todos", department: str = "Todos", search: str = "") -> pd.DataFrame:
    """
    Retorna DataFrame de usuários com filtros aplicados.
    status_filter: 'Todos', 'Ativos', 'Desativados'
    """
    setup_ad_tables()
    conn = get_connection()
    try:
        query = "SELECT sam_account_name, display_name, mail, department, title, is_active, parent_ou_dn, when_created, last_logon FROM ad_cache_users WHERE 1=1"
        params = []
        is_pg = DB_TYPE in ["postgres", "postgresql"]
        ph = "%s" if is_pg else "?"

        if status_filter == "Ativos":
            query += f" AND is_active = {('TRUE' if is_pg else '1')}"
        elif status_filter == "Desativados":
            query += f" AND is_active = {('FALSE' if is_pg else '0')}"

        if department and department != "Todos":
            query += f" AND department = {ph}"
            params.append(department)

        if search:
            query += f" AND (LOWER(sam_account_name) LIKE {ph} OR LOWER(display_name) LIKE {ph} OR LOWER(mail) LIKE {ph})"
            s_val = f"%{search.lower()}%"
            params.extend([s_val, s_val, s_val])

        query += " ORDER BY display_name ASC"

        if params:
            return pd.read_sql_query(query, conn, params=params)
        return pd.read_sql_query(query, conn)
    finally:
        conn.close()


def get_ad_departments() -> List[str]:
    """Retorna a lista de departamentos distintos existentes no cache de usuários."""
    setup_ad_tables()
    conn = get_connection()
    try:
        df = pd.read_sql_query("SELECT DISTINCT department FROM ad_cache_users WHERE department IS NOT NULL AND department != '' ORDER BY department ASC", conn)
        return [str(d) for d in df["department"].dropna().tolist()]
    finally:
        conn.close()


def get_ad_groups_df(search: str = "") -> pd.DataFrame:
    """Retorna DataFrame de grupos de segurança com filtro opcional de busca."""
    setup_ad_tables()
    conn = get_connection()
    try:
        is_pg = DB_TYPE in ["postgres", "postgresql"]
        ph = "%s" if is_pg else "?"
        query = "SELECT sam_account_name, display_name, dn, description, group_type, member_count FROM ad_cache_groups WHERE 1=1"
        params = []
        if search:
            query += f" AND (LOWER(sam_account_name) LIKE {ph} OR LOWER(display_name) LIKE {ph} OR LOWER(description) LIKE {ph})"
            s_val = f"%{search.lower()}%"
            params.extend([s_val, s_val, s_val])
        query += " ORDER BY sam_account_name ASC"

        if params:
            return pd.read_sql_query(query, conn, params=params)
        return pd.read_sql_query(query, conn)
    finally:
        conn.close()


def get_group_members(group_dn: str) -> List[Dict[str, Any]]:
    """Retorna a lista de membros de um grupo específico a partir do cache."""
    setup_ad_tables()
    conn = get_connection()
    cursor = conn.cursor()
    try:
        is_pg = DB_TYPE in ["postgres", "postgresql"]
        ph = "%s" if is_pg else "?"
        cursor.execute(f"""
            SELECT u.sam_account_name, u.display_name, u.mail, u.department, u.is_active
            FROM ad_cache_group_members gm
            LEFT JOIN ad_cache_users u ON LOWER(gm.member_dn) = LOWER(u.dn)
            WHERE LOWER(gm.group_dn) = LOWER({ph})
            ORDER BY u.display_name ASC
        """, (group_dn,))
        rows = cursor.fetchall()
        result = []
        for r in rows:
            result.append({
                "sam_account_name": r[0] or "N/D",
                "display_name": r[1] or "Membro Externo / Objeto",
                "mail": r[2] or "",
                "department": r[3] or "",
                "is_active": r[4] if r[4] is not None else True
            })
        return result
    finally:
        conn.close()
