import re
import datetime
from typing import Dict, Any, List, Optional, Tuple
import ldap3
from ldap3.core.exceptions import LDAPException

from src.config import USERNAME, PASSWORD, DOMINIO, DOMINIO_MMC, DOMINIO_CURTO
from src.terminal import log, GREEN, RED, YELLOW, CYAN
from src.database.ad_db import save_ad_cache, get_ad_sync_meta


def get_ldap_connection(timeout: int = 10) -> Tuple[Optional[ldap3.Connection], Optional[str]]:
    """
    Estabelece conexão e autenticação com o Domain Controller corporativo via ldap3.
    Prioriza autenticação SIMPLE usando o User Principal Name (username@dominio).
    Retorna uma tupla: (connection, error_message).
    """
    if not DOMINIO:
        return None, "Domínio do Active Directory (AD_DOMAIN) não configurado."
    if not USERNAME or not PASSWORD:
        return None, "Credenciais do Active Directory (AD_USER / AD_PASSWORD) incompletas."

    try:
        server = ldap3.Server(DOMINIO, port=389, get_info=ldap3.ALL, connect_timeout=timeout)
        upn = f"{USERNAME}@{DOMINIO}"
        
        conn = ldap3.Connection(
            server,
            user=upn,
            password=PASSWORD,
            authentication=ldap3.SIMPLE,
            auto_bind=False,
            receive_timeout=timeout
        )

        bound = conn.bind()
        if not bound:
            err_msg = f"Falha de autenticação LDAP: {conn.result.get('description', 'Desconhecido')}"
            return None, err_msg

        return conn, None
    except LDAPException as e:
        return None, f"Erro LDAP: {str(e)}"
    except Exception as e:
        return None, f"Erro de conexão com o DC: {str(e)}"


def test_ad_connection() -> Dict[str, Any]:
    """
    Testa a conectividade e autenticação com o Domain Controller e retorna status detalhado.
    """
    start_time = datetime.datetime.now()
    conn, err = get_ldap_connection(timeout=5)
    elapsed = (datetime.datetime.now() - start_time).total_seconds()

    if conn and conn.bound:
        # Pega informações do servidor
        info = {
            "success": True,
            "latency_seconds": round(elapsed, 3),
            "domain": DOMINIO,
            "base_dn": DOMINIO_MMC,
            "user": f"{USERNAME}@{DOMINIO}",
            "server_host": conn.server.host,
            "server_port": conn.server.port,
            "error": None
        }
        conn.unbind()
        return info
    else:
        return {
            "success": False,
            "latency_seconds": round(elapsed, 3),
            "domain": DOMINIO,
            "base_dn": DOMINIO_MMC,
            "user": f"{USERNAME}@{DOMINIO}",
            "server_host": DOMINIO,
            "server_port": 389,
            "error": err or "Não foi possível conectar ou autenticar no Domain Controller."
        }


def _extract_parent_dn(dn: str) -> str:
    """Extrai o DN pai de um Distinguished Name."""
    parts = dn.split(",", 1)
    return parts[1].strip() if len(parts) > 1 else ""


def _is_account_active(uac: Optional[int]) -> bool:
    """
    Verifica se a conta do usuário está ativa através do atributo userAccountControl.
    Bit flag 0x0002 (ACCOUNTDISABLE) = 2 indica conta desativada.
    """
    if uac is None:
        return True
    try:
        uac_val = int(uac)
        return not bool(uac_val & 2)
    except Exception:
        return True


def _clean_str(val: Any) -> str:
    """Remove caracteres de controle invisíveis e não permitidos pelo openpyxl/XML."""
    if val is None:
        return ""
    s = str(val)
    # Remove caracteres de controle ASCII 0x00-0x08, 0x0B-0x0C, 0x0E-0x1F, 0x7F-0x9F
    return re.sub(r'[\x00-\x08\x0b\x0c\x0e-\x1f\x7f-\x9f]', '', s).strip()


def _convert_ad_timestamp(timestamp_val: Any) -> str:
    """Converte atributos como whenCreated ou lastLogonTimestamp para string amigável."""
    if not timestamp_val:
        return ""
    if isinstance(timestamp_val, datetime.datetime):
        return timestamp_val.strftime("%Y-%m-%d %H:%M:%S")
    try:
        # Pode ser int em formato Windows FileTime (100-nanosecond intervals since Jan 1 1601)
        if isinstance(timestamp_val, (int, str)) and str(timestamp_val).isdigit():
            val = int(timestamp_val)
            if val > 116444736000000000:
                epoch = datetime.datetime(1601, 1, 1) + datetime.timedelta(microseconds=val / 10)
                return epoch.strftime("%Y-%m-%d %H:%M:%S")
    except Exception:
        pass
    return str(timestamp_val)


def sync_active_directory_cache(page_size: int = 500, progress_callback: Optional[Any] = None) -> Dict[str, Any]:
    """
    Executa a sincronização completa do Active Directory (OUs, Usuários e Grupos)
    e armazena os dados no banco de dados local para acesso ultra-rápido.
    """
    log("Iniciando sincronização de dados do Active Directory...", symbol="ℹ️", color=CYAN)
    conn, err = get_ldap_connection(timeout=15)
    if not conn:
        log(f"Sincronização abortada: {err}", symbol="❌", color=RED)
        return {"success": False, "error": err}

    try:
        base_dn = DOMINIO_MMC or f"DC={DOMINIO.replace('.', ',DC=')}"

        if progress_callback:
            progress_callback(10, "Consultando Unidades Organizacionais (OUs)...")

        # 1. Busca de OUs
        ous: List[Dict[str, Any]] = []
        ou_generator = conn.extend.standard.paged_search(
            search_base=base_dn,
            search_filter="(objectClass=organizationalUnit)",
            search_scope=ldap3.SUBTREE,
            attributes=["name", "distinguishedName", "description", "canonicalName"],
            paged_size=page_size,
            generator=True
        )

        for resp in ou_generator:
            if resp.get('type') == 'searchResEntry':
                attrs = resp.get('attributes', {})
                dn_val = resp.get('dn', '')
                if not dn_val:
                    continue
                name_val = attrs.get('name') or dn_val.split(",")[0].replace("OU=", "")
                parent_val = _extract_parent_dn(dn_val)
                desc_val = attrs.get('description') or ""
                canon_val = attrs.get('canonicalName') or ""

                ous.append({
                    "dn": dn_val,
                    "name": str(name_val),
                    "parent_dn": parent_val,
                    "description": str(desc_val),
                    "canonical_name": str(canon_val)
                })

        log(f"OUs extraídas: {len(ous)}", symbol="📁", color=CYAN)

        if progress_callback:
            progress_callback(35, "Consultando Usuários do domínio...")

        # 2. Busca de Usuários
        users: List[Dict[str, Any]] = []
        # Exclui contas de computador
        user_filter = "(&(objectClass=user)(!(objectClass=computer)))"
        user_attrs = [
            "sAMAccountName", "displayName", "mail", "department", "title",
            "distinguishedName", "userAccountControl", "whenCreated", "lastLogonTimestamp",
            "pager", "telephoneNumber", "mobile", "physicalDeliveryOfficeName", "company", "manager", "description"
        ]

        # Usa paged search do ldap3
        entry_generator = conn.extend.standard.paged_search(
            search_base=base_dn,
            search_filter=user_filter,
            search_scope=ldap3.SUBTREE,
            attributes=user_attrs,
            paged_size=page_size,
            generator=True
        )

        for resp in entry_generator:
            if resp.get('type') == 'searchResEntry':
                attrs = resp.get('attributes', {})
                dn_val = resp.get('dn', '')
                sam = attrs.get('sAMAccountName') or ''
                if not sam:
                    continue
                display_name = attrs.get('displayName') or sam
                mail = attrs.get('mail') or ''
                dept = attrs.get('department') or ''
                title = attrs.get('title') or ''
                uac = attrs.get('userAccountControl')
                when_created = _convert_ad_timestamp(attrs.get('whenCreated'))
                last_logon = _convert_ad_timestamp(attrs.get('lastLogonTimestamp'))
                parent_ou = _extract_parent_dn(dn_val)
                is_act = _is_account_active(uac)

                pager_val = attrs.get('pager') or ''
                tel_val = attrs.get('telephoneNumber') or ''
                mobile_val = attrs.get('mobile') or ''
                office_val = attrs.get('physicalDeliveryOfficeName') or ''
                company_val = attrs.get('company') or ''
                desc_val = attrs.get('description') or ''
                manager_val = attrs.get('manager') or ''

                # Se manager for um DN completo, extrai o CN para facilitar leitura
                if manager_val and "CN=" in str(manager_val).upper():
                    m_parts = [p for p in str(manager_val).split(",") if p.upper().startswith("CN=")]
                    if m_parts:
                        manager_display = m_parts[0][3:].strip()
                    else:
                        manager_display = str(manager_val)
                else:
                    manager_display = str(manager_val)

                users.append({
                    "sam_account_name": _clean_str(sam),
                    "display_name": _clean_str(display_name),
                    "mail": _clean_str(mail),
                    "department": _clean_str(dept),
                    "title": _clean_str(title),
                    "dn": str(dn_val),
                    "parent_ou_dn": str(parent_ou),
                    "is_active": is_act,
                    "user_account_control": int(uac) if uac is not None else 512,
                    "when_created": when_created,
                    "last_logon": last_logon,
                    "pager": _clean_str(pager_val),
                    "telephone_number": _clean_str(tel_val),
                    "mobile": _clean_str(mobile_val),
                    "office": _clean_str(office_val),
                    "company": _clean_str(company_val),
                    "manager": _clean_str(manager_display),
                    "description": _clean_str(desc_val)
                })

        log(f"Usuários extraídos: {len(users)}", symbol="👥", color=CYAN)

        if progress_callback:
            progress_callback(70, "Consultando Grupos de Segurança e Associações...")

        # 3. Busca de Grupos de Segurança
        groups: List[Dict[str, Any]] = []
        memberships: List[Dict[str, str]] = []
        group_filter = "(objectClass=group)"
        group_attrs = ["sAMAccountName", "displayName", "distinguishedName", "description", "groupType", "member"]

        group_generator = conn.extend.standard.paged_search(
            search_base=base_dn,
            search_filter=group_filter,
            search_scope=ldap3.SUBTREE,
            attributes=group_attrs,
            paged_size=page_size,
            generator=True
        )

        for resp in group_generator:
            if resp.get('type') == 'searchResEntry':
                attrs = resp.get('attributes', {})
                dn_val = resp.get('dn', '')
                sam = attrs.get('sAMAccountName') or ''
                if not sam:
                    continue
                display_name = attrs.get('displayName') or sam
                desc = attrs.get('description') or ''
                g_type = str(attrs.get('groupType') or '')
                raw_members = attrs.get('member') or []
                if isinstance(raw_members, str):
                    raw_members = [raw_members]

                member_count = len(raw_members)
                groups.append({
                    "sam_account_name": str(sam),
                    "display_name": str(display_name),
                    "dn": dn_val,
                    "description": str(desc),
                    "group_type": g_type,
                    "member_count": member_count
                })

                for m_dn in raw_members:
                    memberships.append({
                        "group_dn": dn_val,
                        "member_dn": str(m_dn)
                    })

        log(f"Grupos extraídos: {len(groups)}, Relações: {len(memberships)}", symbol="🛡️", color=CYAN)

        if progress_callback:
            progress_callback(85, "Consultando Computadores e Servidores do domínio...")

        # 4. Busca de Computadores & Servidores
        computers: List[Dict[str, Any]] = []
        comp_filter = "(objectClass=computer)"
        comp_attrs = [
            "name", "sAMAccountName", "dNSHostName", "operatingSystem", "operatingSystemVersion",
            "description", "managedBy", "distinguishedName", "userAccountControl", "whenCreated", "lastLogonTimestamp"
        ]

        comp_generator = conn.extend.standard.paged_search(
            search_base=base_dn,
            search_filter=comp_filter,
            search_scope=ldap3.SUBTREE,
            attributes=comp_attrs,
            paged_size=page_size,
            generator=True
        )

        for resp in comp_generator:
            if resp.get('type') == 'searchResEntry':
                attrs = resp.get('attributes', {})
                dn_val = resp.get('dn', '')
                c_name = attrs.get('name') or attrs.get('sAMAccountName') or ''
                if not c_name:
                    continue
                c_name = str(c_name).rstrip('$')
                dns_host = attrs.get('dNSHostName') or ''
                os_name = attrs.get('operatingSystem') or ''
                os_ver = attrs.get('operatingSystemVersion') or ''
                desc_val = attrs.get('description') or ''
                managed_val = attrs.get('managedBy') or ''
                uac = attrs.get('userAccountControl')
                when_created = _convert_ad_timestamp(attrs.get('whenCreated'))
                last_logon = _convert_ad_timestamp(attrs.get('lastLogonTimestamp'))
                parent_ou = _extract_parent_dn(dn_val)
                is_act = _is_account_active(uac)

                # Limpa managedBy caso seja um DN para exibir nome amigável
                if managed_val and "CN=" in str(managed_val).upper():
                    m_parts = [p for p in str(managed_val).split(",") if p.upper().startswith("CN=")]
                    managed_display = m_parts[0][3:].strip() if m_parts else str(managed_val)
                else:
                    managed_display = str(managed_val)

                computers.append({
                    "name": _clean_str(c_name),
                    "dns_hostname": _clean_str(dns_host),
                    "operating_system": _clean_str(os_name),
                    "os_version": _clean_str(os_ver),
                    "description": _clean_str(desc_val),
                    "managed_by": _clean_str(managed_display),
                    "dn": str(dn_val),
                    "parent_ou_dn": str(parent_ou),
                    "is_active": is_act,
                    "user_account_control": int(uac) if uac is not None else 4096,
                    "when_created": when_created,
                    "last_logon": last_logon
                })

        log(f"Computadores extraídos: {len(computers)}", symbol="💻", color=CYAN)

        if progress_callback:
            progress_callback(95, "Gravando dados de cache no banco de dados...")

        # 5. Salvar tudo no banco relacional
        save_ad_cache(
            ous=ous,
            users=users,
            groups=groups,
            memberships=memberships,
            computers=computers,
            status="success",
            error_message=""
        )

        if progress_callback:
            progress_callback(100, "Sincronização concluída com sucesso!")

        log(f"Sincronização concluída: {len(ous)} OUs, {len(users)} Usuários, {len(groups)} Grupos, {len(computers)} Computadores.", symbol="✅", color=GREEN)
        return {
            "success": True,
            "total_ous": len(ous),
            "total_users": len(users),
            "total_groups": len(groups),
            "total_computers": len(computers),
            "total_memberships": len(memberships),
            "error": None
        }

    except Exception as e:
        log(f"Erro durante a sincronização do Active Directory: {e}", symbol="❌", color=RED)
        save_ad_cache([], [], [], [], computers=[], status="error", error_message=str(e))
        return {"success": False, "error": str(e)}
    finally:
        conn.unbind()
