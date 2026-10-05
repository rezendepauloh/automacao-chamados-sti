# -*- coding: utf-8 -*-
"""
Aba: 🔐 Gerenciador de Senhas (Cofre da Bancada)
Interface moderna inspirada nos gerenciadores nativos de navegadores (Chrome / Edge / Firefox)
com proteção criptográfica Fernet / AES, desbloqueio seguro temporizado, cópia para clipboard,
gerador de senhas e modais nativos (@st.dialog).
"""

import time
import secrets
import string
import pandas as pd
import streamlit as st
from datetime import datetime
from typing import Optional, Dict, Any

from src.config import USERNAME, PASSWORD
from src.database.senhas_db import (
    salvar_senha,
    listar_senhas,
    get_senhas_df,
    obter_senha_decifrada,
    obter_credencial_por_id,
    atualizar_senha,
    excluir_senha,
    get_senhas_stats,
    CATEGORIAS_PADRAO
)
from src.services.ad_ldap_service import authenticate_user_credentials
from src.components.metric_cards import render_metric_cards
from src.components.pagination import (
    render_items_per_page_selector,
    paginate_items,
    render_pagination_controls
)

VAULT_TIMEOUT_SECONDS = 300  # 5 minutos de desbloqueio ativo por padrão


def _gerar_senha_forte(tamanho: int = 16) -> str:
    """Gera senha aleatória criptograficamente segura contendo letras, dígitos e símbolos."""
    chars = string.ascii_letters + string.digits + "!@#$%&*-_+="
    # Garante ao menos um caractere de cada categoria
    obrigatorios = [
        secrets.choice(string.ascii_uppercase),
        secrets.choice(string.ascii_lowercase),
        secrets.choice(string.digits),
        secrets.choice("!@#$%&*-_+=")
    ]
    restantes = [secrets.choice(chars) for _ in range(max(4, tamanho) - 4)]
    senha_lista = obrigatorios + restantes
    secrets.SystemRandom().shuffle(senha_lista)
    return "".join(senha_lista)


def _init_vault_session():
    """Inicializa as variáveis de controle do cofre em st.session_state."""
    if "vault_unlocked_until" not in st.session_state:
        st.session_state["vault_unlocked_until"] = 0
    if "vault_revealed_ids" not in st.session_state:
        st.session_state["vault_revealed_ids"] = set()
    if "vault_last_copied_id" not in st.session_state:
        st.session_state["vault_last_copied_id"] = None
    if "vault_auth_pending_action" not in st.session_state:
        st.session_state["vault_auth_pending_action"] = None


def is_vault_unlocked() -> bool:
    """Verifica se o cofre está desbloqueado na sessão ativa."""
    _init_vault_session()
    return time.time() < st.session_state.get("vault_unlocked_until", 0)


def unlock_vault():
    """Destrava o cofre com tempo limite."""
    _init_vault_session()
    st.session_state["vault_unlocked_until"] = time.time() + VAULT_TIMEOUT_SECONDS


def lock_vault():
    """Bloqueia imediatamente o cofre."""
    _init_vault_session()
    st.session_state["vault_unlocked_until"] = 0
    st.session_state["vault_revealed_ids"] = set()
    st.session_state["vault_auth_pending_action"] = None


# -----------------------------------------------------------------------------
# Modais (@st.dialog)
# -----------------------------------------------------------------------------

@st.dialog("🔒 Confirmação de Segurança do Cofre", width="medium")
def modal_autenticar_operador():
    """
    Modal de autenticação estrita no estilo Windows Hello / Browser Security:
    Exige a validação da senha do operador via Active Directory ou senha mestra do sistema.
    """
    st.markdown("#### 🛡️ Autenticação de Operador")
    st.caption(
        "Para revelar ou copiar credenciais confidenciais, confirme sua identidade com sua senha de rede (AD/Windows)."
    )

    padrao_user = USERNAME or "operador"
    auth_user = st.text_input("Usuário de Rede", value=padrao_user, key="dlg_auth_user")
    auth_pass = st.text_input("Senha de Rede / AD", type="password", key="dlg_auth_pass")

    c1, c2 = st.columns([1, 1])
    with c1:
        if st.button("🔓 Desbloquear", type="primary", use_container_width=True):
            if not auth_pass:
                st.error("Informe a senha do operador.")
                return

            sucesso, msg = authenticate_user_credentials(auth_user, auth_pass)
            if sucesso:
                unlock_vault()
                st.toast(f"✅ Cofre destravado por {VAULT_TIMEOUT_SECONDS // 60} minutos!", icon="🔓")
                st.rerun()
            else:
                st.error(f"Falha na autenticação: {msg}")
    with c2:
        if st.button("Cancelar", use_container_width=True):
            st.rerun()


@st.dialog("➕ Nova Credencial no Cofre", width="large")
def modal_adicionar_credencial():
    """Modal para cadastro de uma nova credencial com gerador de senha integrado."""
    st.markdown("### 🔐 Cadastrar Nova Credencial")
    st.caption("Armazenada com criptografia de padrão militar (AES-Fernet) no banco de dados.")

    titulo = st.text_input("Título / Nome do Recurso *", placeholder="Ex: Roteador Central - Sala Técnica", key="new_cred_titulo")
    
    c_cat, c_url = st.columns(2)
    with c_cat:
        categoria = st.selectbox("Categoria", options=CATEGORIAS_PADRAO, key="new_cred_cat")
    with c_url:
        url_sistema = st.text_input("URL / Endereço IP / Host", placeholder="https://10.10.x.x ou srv-db01", key="new_cred_url")

    c_usr, c_pwd = st.columns(2)
    with c_usr:
        usuario = st.text_input("Usuário / Login *", placeholder="admin ou operador", key="new_cred_usr")
    with c_pwd:
        if "new_cred_pwd_val" not in st.session_state:
            st.session_state["new_cred_pwd_val"] = ""
        senha_val = st.text_input("Senha *", value=st.session_state["new_cred_pwd_val"], type="password", key="new_cred_pwd")

    col_btn_gen, col_gen_feedback = st.columns([1, 2])
    with col_btn_gen:
        if st.button("🎲 Gerar Senha Forte", key="btn_gen_pwd", help="Gera uma senha forte e segura aleatória"):
            senha_gerada = _gerar_senha_forte(18)
            st.session_state["new_cred_pwd_val"] = senha_gerada
            st.rerun()
    with col_gen_feedback:
        if st.session_state.get("new_cred_pwd_val"):
            st.code(st.session_state["new_cred_pwd_val"], language="text")

    observacoes = st.text_area("Observações / Instruções Adicionais", placeholder="Porta, método de conexão, VLAN, observações de segurança...", key="new_cred_obs")

    st.markdown("---")
    col_salvar, col_cancelar = st.columns([1, 1])
    with col_salvar:
        if st.button("💾 Salvar Credencial", type="primary", use_container_width=True):
            if not titulo or not titulo.strip():
                st.error("Informe o título do recurso.")
                return
            if not usuario or not usuario.strip():
                st.error("Informe o usuário de acesso.")
                return
            senha_final = senha_val or st.session_state.get("new_cred_pwd_val")
            if not senha_final or not senha_final.strip():
                st.error("Informe ou gere uma senha.")
                return

            try:
                novo_id = salvar_senha(
                    titulo=titulo,
                    categoria=categoria,
                    usuario=usuario,
                    senha_plana=senha_final,
                    url_sistema=url_sistema,
                    observacoes=observacoes
                )
                st.session_state["new_cred_pwd_val"] = ""
                st.toast(f"✅ Credencial #{novo_id} gravada com sucesso!", icon="🔒")
                st.rerun()
            except Exception as e:
                st.error(f"Erro ao salvar credencial: {e}")
    with col_cancelar:
        if st.button("Cancelar", use_container_width=True):
            st.session_state["new_cred_pwd_val"] = ""
            st.rerun()


@st.dialog("✏️ Editar Credencial", width="large")
def modal_editar_credencial(senha_id: int):
    """Modal para edição segura de credencial existente."""
    cred = obter_credencial_por_id(senha_id)
    if not cred:
        st.error("Credencial não encontrada.")
        return

    st.markdown(f"### ✏️ Editar Credencial #{cred['id']} — {cred['titulo']}")
    st.caption("Modifique os campos desejados. Deixe a nova senha em branco para mantê-la inalterada.")

    titulo = st.text_input("Título / Nome do Recurso *", value=cred["titulo"], key=f"edit_titulo_{senha_id}")

    c_cat, c_url = st.columns(2)
    with c_cat:
        idx_cat = CATEGORIAS_PADRAO.index(cred["categoria"]) if cred["categoria"] in CATEGORIAS_PADRAO else 0
        categoria = st.selectbox("Categoria", options=CATEGORIAS_PADRAO, index=idx_cat, key=f"edit_cat_{senha_id}")
    with c_url:
        url_sistema = st.text_input("URL / IP / Host", value=cred["url_sistema"], key=f"edit_url_{senha_id}")

    c_usr, c_pwd = st.columns(2)
    with c_usr:
        usuario = st.text_input("Usuário / Login *", value=cred["usuario"], key=f"edit_usr_{senha_id}")
    with c_pwd:
        nova_senha = st.text_input(
            "Nova Senha (opcional)",
            type="password",
            placeholder="Deixe vazio para manter a atual",
            key=f"edit_pwd_{senha_id}"
        )

    col_btn_gen, col_gen_feedback = st.columns([1, 2])
    with col_btn_gen:
        if st.button("🎲 Gerar Nova Senha", key=f"btn_gen_edit_{senha_id}"):
            st.session_state[f"edit_gen_pwd_{senha_id}"] = _gerar_senha_forte(18)
            st.rerun()
    with col_gen_feedback:
        if st.session_state.get(f"edit_gen_pwd_{senha_id}"):
            st.code(st.session_state[f"edit_gen_pwd_{senha_id}"], language="text")

    observacoes = st.text_area("Observações", value=cred["observacoes"], key=f"edit_obs_{senha_id}")

    st.markdown("---")
    col_salvar, col_del, col_cancel = st.columns([1.5, 1.2, 1])
    with col_salvar:
        if st.button("💾 Salvar Alterações", type="primary", use_container_width=True, key=f"btn_save_{senha_id}"):
            senha_a_gravar = nova_senha or st.session_state.get(f"edit_gen_pwd_{senha_id}") or None
            try:
                atualizar_senha(
                    senha_id=senha_id,
                    titulo=titulo,
                    categoria=categoria,
                    usuario=usuario,
                    senha_plana=senha_a_gravar,
                    url_sistema=url_sistema,
                    observacoes=observacoes
                )
                if f"edit_gen_pwd_{senha_id}" in st.session_state:
                    del st.session_state[f"edit_gen_pwd_{senha_id}"]
                st.toast("✅ Credencial atualizada com sucesso!", icon="💾")
                st.rerun()
            except Exception as e:
                st.error(f"Erro ao atualizar: {e}")
    with col_del:
        if st.button("🗑️ Excluir", use_container_width=True, key=f"btn_del_{senha_id}"):
            excluir_senha(senha_id)
            st.toast("🗑️ Credencial excluída com sucesso!", icon="🗑️")
            st.rerun()
    with col_cancel:
        if st.button("Fechar", use_container_width=True, key=f"btn_cancel_{senha_id}"):
            st.rerun()


# -----------------------------------------------------------------------------
# Renderização Principal da Página
# -----------------------------------------------------------------------------

def render_senhas_page():
    """Renderiza a página visual completa do Gerenciador de Senhas / Cofre da Bancada."""
    _init_vault_session()
    unlocked = is_vault_unlocked()

    # Header / Banner com estilo premium
    status_tag = (
        '<span style="background-color: rgba(16, 185, 129, 0.2); color: #10b981; font-size: 13px; font-weight: 600; padding: 4px 12px; border-radius: 20px; border: 1px solid #10b981;">🔓 Cofre Desbloqueado</span>'
        if unlocked
        else '<span style="background-color: rgba(239, 68, 68, 0.2); color: #f87171; font-size: 13px; font-weight: 600; padding: 4px 12px; border-radius: 20px; border: 1px solid #ef4444;">🔒 Cofre Bloqueado</span>'
    )

    st.markdown(f"""
        <div style="background: var(--metric-bg, #1e293b); padding: 18px 24px; border-radius: 12px; border-left: 6px solid #8b5cf6; border-top: 1px solid var(--metric-border, #2d3139); border-right: 1px solid var(--metric-border, #2d3139); border-bottom: 1px solid var(--metric-border, #2d3139); margin-bottom: 20px; box-shadow: 0 2px 8px rgba(0,0,0,0.08);">
            <div style="display: flex; justify-content: space-between; align-items: center; flex-wrap: wrap; gap: 10px;">
                <h2 style="color: var(--metric-value-color, #ffffff); margin: 0; font-size: 24px; font-weight: 700;">🔐 Gerenciador de Senhas & Cofre da Bancada</h2>
                <div>{status_tag}</div>
            </div>
            <p style="color: var(--metric-title-color, #94a3b8); margin: 6px 0 0 0; font-size: 14px;">
                Central de credenciais de sistemas, switches, roteadores, servidores e ferramentas da equipe da Bancada STI com proteção criptográfica AES-Fernet e revelação estilo browser.
            </p>
        </div>
    """, unsafe_allow_html=True)

    # 1. Cards KPI do Topo
    stats = get_senhas_stats()
    kpi_cards = [
        {
            "title": "TOTAL DE CREDENCIAIS",
            "value": stats["total_senhas"],
            "border_color": "#8b5cf6",
            "subtitle": "registros no cofre"
        },
        {
            "title": "CATEGORIAS ATIVAS",
            "value": stats["total_categorias"],
            "border_color": "#3b82f6",
            "subtitle": "grupos de serviços"
        },
        {
            "title": "SISTEMAS COM LINK / IP",
            "value": stats["com_link"],
            "border_color": "#10b981",
            "subtitle": "acesso rápido disponível"
        },
        {
            "title": "ESTADO DO COFRE",
            "value": "Desbloqueado" if unlocked else "Protegido",
            "border_color": "#10b981" if unlocked else "#ef4444",
            "value_color": "#10b981" if unlocked else "#f87171",
            "subtitle": f"{VAULT_TIMEOUT_SECONDS // 60}m timeout" if unlocked else "Requer senha AD"
        }
    ]
    render_metric_cards(kpi_cards)
    st.markdown("<br>", unsafe_allow_html=True)

    # 2. Barra de Ações Rápidas (Adicionar, Bloquear/Desbloquear)
    col_act_left, col_act_right = st.columns([3, 1.5])
    with col_act_left:
        c_add, c_lock = st.columns([1.5, 1.5])
        with c_add:
            if st.button("➕ Nova Credencial", type="primary", use_container_width=True):
                modal_adicionar_credencial()
        with c_lock:
            if unlocked:
                restante = max(0, int(st.session_state.get("vault_unlocked_until", 0) - time.time()))
                mins = restante // 60
                secs = restante % 60
                if st.button(f"🔒 Trancar Agora ({mins:02d}:{secs:02d})", use_container_width=True, help="Bloqueia imediatamente a visualização de todas as senhas"):
                    lock_vault()
                    st.toast("🔒 Cofre bloqueado com sucesso!", icon="🔒")
                    st.rerun()
            else:
                if st.button("🔓 Desbloquear Cofre", use_container_width=True, help="Autentica com sua credencial de rede para visualizar senhas"):
                    modal_autenticar_operador()

    # 3. Filtros e Pesquisa no Menu Lateral (Sidebar)
    st.sidebar.markdown("### 🔍 Filtros do Cofre")
    busca_query = st.sidebar.text_input(
        "Buscar credencial",
        placeholder="Nome, usuário, URL, notas...",
        key="senhas_search"
    )
    cat_filtro = st.sidebar.selectbox(
        "Categoria",
        options=["Todas"] + CATEGORIAS_PADRAO,
        key="senhas_cat_filtro"
    )

    # Quantidade por página via sidebar selector padrão
    items_per_page = render_items_per_page_selector(key_prefix="senhas", options=[10, 20, 50, "Todos"], default_index=0)

    st.markdown("### 📋 Credenciais Cadastradas")

    # Carregar registros
    credenciais = listar_senhas(
        categoria=None if cat_filtro == "Todas" else cat_filtro,
        busca=busca_query if busca_query else None
    )

    if not credenciais:
        st.info("Nenhuma credencial encontrada para os filtros selecionados.")
        return

    # Paginação
    cred_slice, current_page, total_pages, total_items = paginate_items(
        credenciais,
        page_key="senhas_grid",
        items_per_page=items_per_page
    )

    # 4. Renderização do Grid no estilo Password Manager do Chrome/Edge
    for item in cred_slice:
        s_id = item["id"]
        is_revealed = (s_id in st.session_state["vault_revealed_ids"]) and unlocked

        with st.container():
            st.markdown("""
                <style>
                .vault-card {
                    background: var(--metric-bg, #1e293b);
                    border: 1px solid var(--metric-border, #334155);
                    border-radius: 10px;
                    padding: 14px 18px;
                    margin-bottom: 12px;
                    transition: border-color 0.2s;
                }
                .vault-card:hover {
                    border-color: #8b5cf6;
                }
                </style>
            """, unsafe_allow_html=True)

            c_info, c_user, c_pass, c_actions = st.columns([3, 2, 2.5, 2.5], vertical_alignment="center")

            with c_info:
                st.markdown(f"**{item['titulo']}**")
                url_str = item["url_sistema"]
                if url_str:
                    link_href = url_str if url_str.startswith("http") else f"http://{url_str}"
                    st.markdown(f"<span style='font-size: 0.82rem;'><a href='{link_href}' target='_blank' style='color: #38bdf8; text-decoration: none;'>🌐 {url_str}</a></span>", unsafe_allow_html=True)
                st.caption(f"🏷️ `{item['categoria']}`")

            with c_user:
                st.markdown(f"👤 `{item['usuario']}`")
                if item["observacoes"]:
                    st.caption(f"📝 {item['observacoes'][:40]}..." if len(item['observacoes']) > 40 else f"📝 {item['observacoes']}")

            with c_pass:
                if is_revealed:
                    senha_real = obter_senha_decifrada(s_id) or "---"
                    st.code(senha_real, language="text")
                else:
                    st.markdown("<span style='font-family: monospace; font-size: 1.15rem; letter-spacing: 2px; color: #94a3b8;'>••••••••</span>", unsafe_allow_html=True)

            with c_actions:
                col_btn_view, col_btn_copy, col_btn_edit = st.columns(3)
                
                # Botão Olho (Ver / Ocultar)
                with col_btn_view:
                    if is_revealed:
                        if st.button("🙈", key=f"btn_hide_{s_id}", help="Ocultar senha"):
                            st.session_state["vault_revealed_ids"].discard(s_id)
                            st.rerun()
                    else:
                        if st.button("👁️", key=f"btn_show_{s_id}", help="Revelar senha (requer autenticação)"):
                            if unlocked:
                                st.session_state["vault_revealed_ids"].add(s_id)
                                st.rerun()
                            else:
                                modal_autenticar_operador()

                # Botão Copiar Senha (mostra na interface/toast)
                with col_btn_copy:
                    if st.button("📋", key=f"btn_copy_{s_id}", help="Copiar senha"):
                        if unlocked:
                            senha_real = obter_senha_decifrada(s_id)
                            if senha_real:
                                st.session_state["vault_revealed_ids"].add(s_id)
                                st.toast(f"📋 Senha de '{item['titulo']}' revelada no card acima para cópia!", icon="🔑")
                                st.rerun()
                        else:
                            modal_autenticar_operador()

                # Botão Editar
                with col_btn_edit:
                    if st.button("✏️", key=f"btn_edit_{s_id}", help="Editar credencial"):
                        modal_editar_credencial(s_id)

            st.markdown("---")

    # Controles de Paginação
    render_pagination_controls(
        page_key="senhas_grid",
        current_page=current_page,
        total_pages=total_pages,
        total_items=total_items,
        items_per_page=items_per_page
    )
