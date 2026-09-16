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
        pager TEXT,
        telephone_number TEXT,
        mobile TEXT,
        office TEXT,
        company TEXT,
        manager TEXT,
        description TEXT,
        updated_at TEXT
    )
    """)

    # Migrações seguras (caso a tabela já tenha sido criada anteriormente)
    new_cols = [
        ("pager", "TEXT"),
        ("telephone_number", "TEXT"),
        ("mobile", "TEXT"),
        ("office", "TEXT"),
        ("company", "TEXT"),
        ("manager", "TEXT"),
        ("description", "TEXT")
    ]
    for col_name, col_type in new_cols:
        try:
            cursor.execute(f"ALTER TABLE ad_cache_users ADD COLUMN {col_name} {col_type}")
            conn.commit()
        except Exception:
            conn.rollback()

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

    # Índices essenciais para consultas instantâneas em dezenas de milhares de registros
    cursor.execute("CREATE INDEX IF NOT EXISTS idx_ad_members_group ON ad_cache_group_members (group_dn)")
    cursor.execute("CREATE INDEX IF NOT EXISTS idx_ad_members_member ON ad_cache_group_members (member_dn)")
    cursor.execute("CREATE INDEX IF NOT EXISTS idx_ad_users_dn ON ad_cache_users (dn)")
    cursor.execute("CREATE INDEX IF NOT EXISTS idx_ad_groups_dn ON ad_cache_groups (dn)")

    # 5. Tabela de Computadores & Servidores
    cursor.execute(f"""
    CREATE TABLE IF NOT EXISTS ad_cache_computers (
        name TEXT PRIMARY KEY,
        dns_hostname TEXT,
        operating_system TEXT,
        os_version TEXT,
        description TEXT,
        managed_by TEXT,
        dn TEXT,
        parent_ou_dn TEXT,
        is_active BOOLEAN,
        user_account_control INTEGER,
        when_created TEXT,
        last_logon TEXT,
        updated_at TEXT
    )
    """)
    cursor.execute("CREATE INDEX IF NOT EXISTS idx_ad_comp_dn ON ad_cache_computers (dn)")
    cursor.execute("CREATE INDEX IF NOT EXISTS idx_ad_comp_name ON ad_cache_computers (name)")
    cursor.execute("CREATE INDEX IF NOT EXISTS idx_ad_comp_os ON ad_cache_computers (operating_system)")

    # 6. Metadados de Sincronização
    cursor.execute(f"""
    CREATE TABLE IF NOT EXISTS ad_sync_meta (
        id INTEGER PRIMARY KEY,
        last_sync_at TEXT,
        status TEXT,
        total_ous INTEGER DEFAULT 0,
        total_users INTEGER DEFAULT 0,
        total_groups INTEGER DEFAULT 0,
        total_computers INTEGER DEFAULT 0,
        error_message TEXT
    )
    """)

    # Migração segura para total_computers caso a tabela ad_sync_meta já exista
    try:
        cursor.execute("ALTER TABLE ad_sync_meta ADD COLUMN total_computers INTEGER DEFAULT 0")
        conn.commit()
    except Exception:
        conn.rollback()

    conn.commit()
    conn.close()


def save_ad_cache(
    ous: List[Dict[str, Any]],
    users: List[Dict[str, Any]],
    groups: List[Dict[str, Any]],
    memberships: List[Dict[str, str]],
    computers: Optional[List[Dict[str, Any]]] = None,
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
    computers = computers or []

    try:
        # Limpa tabelas de cache para consistência total da árvore
        cursor.execute("DELETE FROM ad_cache_ous")
        cursor.execute("DELETE FROM ad_cache_users")
        cursor.execute("DELETE FROM ad_cache_groups")
        cursor.execute("DELETE FROM ad_cache_group_members")
        cursor.execute("DELETE FROM ad_cache_computers")

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
                    parent_ou_dn, is_active, user_account_control, when_created, last_logon,
                    pager, telephone_number, mobile, office, company, manager, description, updated_at
                )
                VALUES (%s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s)
            """ if DB_TYPE in ["postgres", "postgresql"] else """
                INSERT INTO ad_cache_users (
                    sam_account_name, display_name, mail, department, title, dn,
                    parent_ou_dn, is_active, user_account_control, when_created, last_logon,
                    pager, telephone_number, mobile, office, company, manager, description, updated_at
                )
                VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
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
                str(u.get("pager", "")),
                str(u.get("telephone_number", "")),
                str(u.get("mobile", "")),
                str(u.get("office", "")),
                str(u.get("company", "")),
                str(u.get("manager", "")),
                str(u.get("description", "")),
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

        # Inserção de Computadores & Servidores
        for c in computers:
            cursor.execute("""
                INSERT INTO ad_cache_computers (
                    name, dns_hostname, operating_system, os_version, description, managed_by,
                    dn, parent_ou_dn, is_active, user_account_control, when_created, last_logon, updated_at
                )
                VALUES (%s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s)
            """ if DB_TYPE in ["postgres", "postgresql"] else """
                INSERT INTO ad_cache_computers (
                    name, dns_hostname, operating_system, os_version, description, managed_by,
                    dn, parent_ou_dn, is_active, user_account_control, when_created, last_logon, updated_at
                )
                VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
            """, (
                c.get("name", ""),
                c.get("dns_hostname", ""),
                c.get("operating_system", ""),
                c.get("os_version", ""),
                c.get("description", ""),
                c.get("managed_by", ""),
                c.get("dn", ""),
                c.get("parent_ou_dn", ""),
                bool(c.get("is_active", True)),
                int(c.get("user_account_control", 4096)),
                str(c.get("when_created", "")),
                str(c.get("last_logon", "")),
                now_str
            ))

        # Atualiza Metadados de Sincronização (ID fixo = 1)
        cursor.execute("DELETE FROM ad_sync_meta WHERE id = 1")
        cursor.execute("""
            INSERT INTO ad_sync_meta (id, last_sync_at, status, total_ous, total_users, total_groups, total_computers, error_message)
            VALUES (%s, %s, %s, %s, %s, %s, %s, %s)
        """ if DB_TYPE in ["postgres", "postgresql"] else """
            INSERT INTO ad_sync_meta (id, last_sync_at, status, total_ous, total_users, total_groups, total_computers, error_message)
            VALUES (?, ?, ?, ?, ?, ?, ?, ?)
        """, (
            1,
            now_str,
            status,
            len(ous),
            len(users),
            len(groups),
            len(computers),
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
        cursor.execute("SELECT last_sync_at, status, total_ous, total_users, total_groups, total_computers, error_message FROM ad_sync_meta WHERE id = 1")
        row = cursor.fetchone()
        if row:
            return {
                "last_sync_at": row[0],
                "status": row[1],
                "total_ous": row[2],
                "total_users": row[3],
                "total_groups": row[4],
                "total_computers": row[5] or 0,
                "error_message": row[6] or ""
            }
        return {
            "last_sync_at": None,
            "status": "never_run",
            "total_ous": 0,
            "total_users": 0,
            "total_groups": 0,
            "total_computers": 0,
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
        query = """
            SELECT sam_account_name, display_name, mail, department, title, is_active, parent_ou_dn,
                   when_created, last_logon, dn, user_account_control,
                   pager, telephone_number, mobile, office, company, manager, description
            FROM ad_cache_users WHERE 1=1
        """
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
            SELECT u.sam_account_name, u.display_name, u.mail, u.department, u.is_active, u.pager, u.telephone_number
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
                "is_active": r[4] if r[4] is not None else True,
                "pager": r[5] or "-",
                "telephone_number": r[6] or "-"
            })
        return result
    finally:
        conn.close()


def get_user_groups(user_dn: str) -> List[Dict[str, Any]]:
    """Retorna os grupos de segurança dos quais o usuário é membro (associação direta)."""
    if not user_dn:
        return []
    setup_ad_tables()
    conn = get_connection()
    cursor = conn.cursor()
    try:
        is_pg = DB_TYPE in ["postgres", "postgresql"]
        ph = "%s" if is_pg else "?"
        # Compara de forma direta para usar o índice B-Tree imediatamente; se não achar tenta case-insensitive
        cursor.execute(f"""
            SELECT g.sam_account_name, g.display_name, g.description, g.member_count, g.dn
            FROM ad_cache_group_members gm
            JOIN ad_cache_groups g ON gm.group_dn = g.dn
            WHERE gm.member_dn = {ph}
            ORDER BY g.sam_account_name ASC
        """, (user_dn,))
        rows = cursor.fetchall()

        if not rows:
            cursor.execute(f"""
                SELECT g.sam_account_name, g.display_name, g.description, g.member_count, g.dn
                FROM ad_cache_group_members gm
                JOIN ad_cache_groups g ON LOWER(gm.group_dn) = LOWER(g.dn)
                WHERE LOWER(gm.member_dn) = LOWER({ph})
                ORDER BY g.sam_account_name ASC
            """, (user_dn,))
            rows = cursor.fetchall()

        result = []
        for r in rows:
            result.append({
                "sam_account_name": r[0] or "",
                "display_name": r[1] or r[0] or "",
                "description": r[2] or "",
                "member_count": r[3] or 0,
                "dn": r[4] or ""
            })
        return result
    finally:
        conn.close()


def get_ad_computers_df(
    status_filter: str = "Todos",
    os_filter: Any = "Todos",
    search: str = "",
    machine_type: str = "Todos",
    stale_days: str = "Todos"
) -> pd.DataFrame:
    """
    Retorna DataFrame de computadores/servidores cadastrados no Active Directory.
    status_filter: 'Todos', 'Ativos', 'Desativados'
    os_filter: 'Todos', nome de SO específico (str) ou lista de nomes de SOs (List[str])
    machine_type: 'Todos', 'Servidores', 'Estações de Trabalho'
    stale_days: 'Todos', '> 30 dias', '> 90 dias', '> 180 dias', 'Sem Logon Registrado'
    """
    setup_ad_tables()
    conn = get_connection()
    try:
        query = """
            SELECT name, dns_hostname, operating_system, os_version, description, managed_by,
                   dn, parent_ou_dn, is_active, user_account_control, when_created, last_logon
            FROM ad_cache_computers WHERE 1=1
        """
        params = []
        is_pg = DB_TYPE in ["postgres", "postgresql"]
        ph = "%s" if is_pg else "?"

        if status_filter == "Ativos":
            query += f" AND is_active = {('TRUE' if is_pg else '1')}"
        elif status_filter == "Desativados":
            query += f" AND is_active = {('FALSE' if is_pg else '0')}"

        if machine_type == "Servidores":
            query += f" AND LOWER(operating_system) LIKE {ph}"
            params.append("%server%")
        elif machine_type == "Estações de Trabalho":
            query += f" AND (LOWER(operating_system) NOT LIKE {ph} OR operating_system IS NULL OR operating_system = '')"
            params.append("%server%")

        if stale_days == "Sem Logon Registrado":
            query += " AND (last_logon IS NULL OR last_logon = '' OR last_logon = 'None')"
        elif stale_days in ["> 30 dias", "> 90 dias", "> 180 dias"]:
            days_map = {"> 30 dias": 30, "> 90 dias": 90, "> 180 dias": 180}
            d_val = days_map[stale_days]
            cutoff_date = (datetime.now() - pd.Timedelta(days=d_val)).strftime("%Y-%m-%d %H:%M:%S")
            query += f" AND last_logon IS NOT NULL AND last_logon != '' AND last_logon != 'None' AND last_logon < {ph}"
            params.append(cutoff_date)

        if os_filter and os_filter != "Todos":
            if isinstance(os_filter, (list, tuple, set)):
                os_list = [str(item) for item in os_filter if item and item != "Todos"]
                if os_list:
                    placeholders = ", ".join([ph] * len(os_list))
                    query += f" AND operating_system IN ({placeholders})"
                    params.extend(os_list)
            else:
                query += f" AND operating_system = {ph}"
                params.append(str(os_filter))

        if search:
            query += f" AND (LOWER(name) LIKE {ph} OR LOWER(dns_hostname) LIKE {ph} OR LOWER(description) LIKE {ph} OR LOWER(managed_by) LIKE {ph})"
            s_val = f"%{search.lower()}%"
            params.extend([s_val, s_val, s_val, s_val])

        query += " ORDER BY name ASC"

        if params:
            return pd.read_sql_query(query, conn, params=params)
        return pd.read_sql_query(query, conn)
    finally:
        conn.close()


def get_ad_operating_systems() -> List[str]:
    """Retorna lista de Sistemas Operacionais distintos encontrados nos computadores do AD."""
    setup_ad_tables()
    conn = get_connection()
    try:
        df = pd.read_sql_query(
            "SELECT DISTINCT operating_system FROM ad_cache_computers WHERE operating_system IS NOT NULL AND operating_system != '' ORDER BY operating_system ASC",
            conn
        )
        return [str(o) for o in df["operating_system"].dropna().tolist()]
    finally:
        conn.close()


def get_ad_ou_stats() -> Dict[str, Dict[str, Any]]:
    """
    Retorna estatísticas agregadas de usuários e computadores por OU para enriquecer a árvore GoJS.
    Estrutura: {ou_dn: {'user_count': int, 'comp_count': int, 'active_users': int, 'sample_users': list}}
    """
    setup_ad_tables()
    conn = get_connection()
    try:
        # Usuários agregados por OU pai
        df_users = pd.read_sql_query("""
            SELECT parent_ou_dn, sam_account_name, display_name, mail, telephone_number, title, is_active
            FROM ad_cache_users
            WHERE parent_ou_dn IS NOT NULL AND parent_ou_dn != ''
        """, conn)

        # Computadores agregados por OU pai
        df_comps = pd.read_sql_query("""
            SELECT parent_ou_dn, COUNT(*) as qtd_comps
            FROM ad_cache_computers
            WHERE parent_ou_dn IS NOT NULL AND parent_ou_dn != ''
            GROUP BY parent_ou_dn
        """, conn)

        stats = {}
        comp_map = dict(zip(df_comps["parent_ou_dn"], df_comps["qtd_comps"]))

        if not df_users.empty:
            grouped = df_users.groupby("parent_ou_dn")
            for ou_dn, group in grouped:
                u_count = len(group)
                active_count = int(group["is_active"].sum())
                # Pega até 3 usuários de amostra com dados de contato
                sample_users = []
                for _, u_row in group.head(3).iterrows():
                    sample_users.append({
                        "name": u_row["display_name"] or u_row["sam_account_name"],
                        "title": u_row["title"] or "",
                        "mail": u_row["mail"] or "",
                        "phone": u_row["telephone_number"] or ""
                    })

                stats[ou_dn] = {
                    "user_count": u_count,
                    "active_users": active_count,
                    "comp_count": comp_map.get(ou_dn, 0),
                    "sample_users": sample_users
                }

        # Inclui OUs que só tem computadores
        for ou_dn, c_count in comp_map.items():
            if ou_dn not in stats:
                stats[ou_dn] = {
                    "user_count": 0,
                    "active_users": 0,
                    "comp_count": c_count,
                    "sample_users": []
                }

        return stats
    finally:
        conn.close()


def get_ou_members(ou_dn: str) -> List[Dict[str, Any]]:
    """Retorna todos os usuários alocados diretamente nesta OU."""
    if not ou_dn:
        return []
    setup_ad_tables()
    conn = get_connection()
    try:
        is_pg = DB_TYPE in ["postgres", "postgresql"]
        ph = "%s" if is_pg else "?"
        query = f"""
            SELECT sam_account_name, display_name, mail, department, title, is_active,
                   pager, telephone_number, mobile, office, company, manager, description,
                   when_created, last_logon, user_account_control, dn, parent_ou_dn
            FROM ad_cache_users
            WHERE LOWER(parent_ou_dn) = LOWER({ph})
            ORDER BY display_name ASC
        """
        df = pd.read_sql_query(query, conn, params=[ou_dn])
        return df.fillna("").to_dict(orient="records")
    finally:
        conn.close()


def get_ou_computers(ou_dn: str) -> List[Dict[str, Any]]:
    """Retorna todos os computadores/servidores alocados diretamente nesta OU."""
    if not ou_dn:
        return []
    setup_ad_tables()
    conn = get_connection()
    try:
        is_pg = DB_TYPE in ["postgres", "postgresql"]
        ph = "%s" if is_pg else "?"
        query = f"""
            SELECT name, dns_hostname, operating_system, os_version, description, managed_by,
                   dn, parent_ou_dn, is_active, user_account_control, when_created, last_logon
            FROM ad_cache_computers
            WHERE LOWER(parent_ou_dn) = LOWER({ph})
            ORDER BY name ASC
        """
        df = pd.read_sql_query(query, conn, params=[ou_dn])
        return df.fillna("").to_dict(orient="records")
    finally:
        conn.close()


def get_all_ou_entities_compact() -> Dict[str, Any]:
    """
    Retorna todos os usuários e computadores agrupados por parent_ou_dn de forma ultra compacta e rápida,
    para permitir buscas instantâneas e abertura de modais na Árvore GoJS sem chamadas de rede adicionais.
    """
    setup_ad_tables()
    conn = get_connection()
    try:
        cur = conn.cursor()
        users_rows = cur.execute("""
            SELECT parent_ou_dn, sam_account_name, display_name, mail, title, telephone_number,
                   pager, department, is_active, user_account_control, office, manager, when_created, last_logon
            FROM ad_cache_users
            ORDER BY display_name ASC
        """).fetchall()

        comps_rows = cur.execute("""
            SELECT parent_ou_dn, name, dns_hostname, operating_system, os_version, managed_by,
                   description, is_active, user_account_control, when_created, last_logon
            FROM ad_cache_computers
            ORDER BY name ASC
        """).fetchall()

        users_by_ou = {}
        for r in users_rows:
            ou = str(r[0] or "").strip()
            users_by_ou.setdefault(ou, []).append({
                "sam": r[1] or "",
                "name": r[2] or r[1] or "",
                "mail": r[3] or "",
                "title": r[4] or "",
                "phone": r[5] or "",
                "pager": r[6] or "",
                "dept": r[7] or "",
                "active": bool(r[8]),
                "uac": r[9],
                "office": r[10] or "",
                "manager": r[11] or "",
                "created": str(r[12] or ""),
                "logon": str(r[13] or "")
            })

        comps_by_ou = {}
        for r in comps_rows:
            ou = str(r[0] or "").strip()
            comps_by_ou.setdefault(ou, []).append({
                "name": r[1] or "",
                "dns": r[2] or "",
                "os": r[3] or "",
                "os_ver": r[4] or "",
                "managed_by": r[5] or "",
                "desc": r[6] or "",
                "active": bool(r[7]),
                "uac": r[8],
                "created": str(r[9] or ""),
                "logon": str(r[10] or "")
            })

        return {"users": users_by_ou, "comps": comps_by_ou}
    finally:
        conn.close()


