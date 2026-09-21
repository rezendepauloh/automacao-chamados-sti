import json
import streamlit as st
import pandas as pd
from datetime import datetime
from src.components.subtabs import render_subtabs
from src.components.metric_cards import render_metric_cards
from src.components.pagination import paginate_items, render_pagination_controls
from src.database.sccm_db import (
    get_sccm_devices_df,
    get_sccm_users_df,
    get_sccm_collections_df,
    setup_sccm_tables
)
from src.services.sccm_service import (
    sync_sccm_devices,
    sync_sccm_users,
    sync_sccm_collections,
    sync_all_sccm
)

# Mapeamento de sub-abas idêntico à hierarquia de Ativos e Conformidade do Console SCCM
SCCM_SUBTABS = {
    "dispositivos": "💻 Dispositivos",
    "usuarios": "👤 Usuários",
    "colecoes_dispositivos": "📁 Coleções de Dispositivos",
    "colecoes_usuarios": "👥 Coleções de Usuários",
    "conformidade": "🛡️ Configurações & Conformidade"
}

@st.dialog("🔍 Ficha Técnica & Ações Rápidas do Computador")
def modal_device_details(device_row: dict):
    """Modal nativo do Streamlit exibindo detalhes da estação e botões de ação remota."""
    name = str(device_row.get("name", "N/A"))
    user = str(device_row.get("last_logon_user", "N/A"))
    ips = str(device_row.get("ip_addresses", "N/A"))
    macs = str(device_row.get("mac_addresses", "N/A"))
    model = str(device_row.get("model", "N/A"))
    mfg = str(device_row.get("manufacturer", "N/A"))
    os_name = str(device_row.get("operating_system", "N/A"))
    os_ver = str(device_row.get("os_version", "N/A"))
    client_ver = str(device_row.get("client_version", "N/A"))
    client_act = bool(device_row.get("client_active", 1))
    site = str(device_row.get("ad_site_name", "N/A"))
    dn = str(device_row.get("distinguished_name", "N/A"))
    last_active = str(device_row.get("last_active_time", "N/A"))

    st.markdown(f"### 🖥️ {name}")
    st.caption(f"Usuário associado: **{user}** | Site AD: **{site}**")

    # --- PAINEL DE DISPARO RÁPIDO VIA BANCADA:// ---
    st.markdown("#### ⚡ Ações Remotas Instantâneas")
    col1, col2, col3, col4 = st.columns(4)

    with col1:
        cmrc_url = f"bancada://run?tool=cmrc&host={name}"
        st.markdown(f"""
        <a href="{cmrc_url}" style="text-decoration:none;">
            <button style="width:100%; height:45px; background-color:#0284c7; color:white; border:none; border-radius:6px; font-weight:600; cursor:pointer;">
                🎮 Controle Remoto
            </button>
        </a>
        """, unsafe_allow_html=True)

    with col2:
        rdp_url = f"bancada://run?tool=rdp&host={name}"
        st.markdown(f"""
        <a href="{rdp_url}" style="text-decoration:none;">
            <button style="width:100%; height:45px; background-color:#16a34a; color:white; border:none; border-radius:6px; font-weight:600; cursor:pointer;">
                🖥️ Conexão RDP
            </button>
        </a>
        """, unsafe_allow_html=True)

    with col3:
        exp_url = f"bancada://run?tool=explorer&host={name}"
        st.markdown(f"""
        <a href="{exp_url}" style="text-decoration:none;">
            <button style="width:100%; height:45px; background-color:#d97706; color:white; border:none; border-radius:6px; font-weight:600; cursor:pointer;">
                📂 Explorer C$
            </button>
        </a>
        """, unsafe_allow_html=True)

    with col4:
        ping_url = f"bancada://run?tool=ping&host={name}"
        st.markdown(f"""
        <a href="{ping_url}" style="text-decoration:none;">
            <button style="width:100%; height:45px; background-color:#475569; color:white; border:none; border-radius:6px; font-weight:600; cursor:pointer;">
                ⚡ Teste Ping
            </button>
        </a>
        """, unsafe_allow_html=True)

    st.markdown("---")

    # --- INFORMAÇÕES DETALHADAS DE HARDWARE E SISTEMA ---
    col_inf1, col_inf2 = st.columns(2)
    with col_inf1:
        st.markdown(f"**Fabricante:** {mfg}")
        st.markdown(f"**Modelo:** {model}")
        st.markdown(f"**Endereço(s) IP:** `{ips}`")
        st.markdown(f"**MAC Address:** `{macs}`")
    with col_inf2:
        st.markdown(f"**Sistema Operacional:** {os_name}")
        st.markdown(f"**Versão de Build:** {os_ver}")
        st.markdown(f"**Versão do Cliente SCCM:** {client_ver}")
        status_txt = "🟢 Ativo" if client_act else "🔴 Inativo"
        st.markdown(f"**Status no SCCM:** {status_txt}")

    if dn and dn != "N/A":
        st.caption(f"**Distinguished Name (DN):** `{dn}`")

    # Exibe JSON bruto em expander caso queira auditar propriedades adicionais
    raw_str = device_row.get("raw_json")
    if raw_str:
        with st.expander("📄 Ver Dados Brutos WMI/CIM"):
            try:
                st.json(json.loads(raw_str))
            except Exception:
                st.code(raw_str)


def render_subtab_dispositivos():
    """Renderiza a listagem de computadores com filtros, métricas e seleção."""
    df = get_sccm_devices_df()

    # Métricas KPI no topo
    total = len(df)
    win11 = len(df[df["operating_system"].str.contains("Windows 11", case=False, na=False)]) if not df.empty else 0
    win10 = len(df[df["operating_system"].str.contains("Windows 10", case=False, na=False)]) if not df.empty else 0
    ativos = len(df[df["client_active"] == 1]) if not df.empty else 0

    render_metric_cards([
        {"title": "TOTAL DE COMPUTADORES", "value": f"{total:,}".replace(",", "."), "border_color": "#0ea5e9", "subtitle": "no SCCM"},
        {"title": "CLIENTES ATIVOS", "value": f"{ativos:,}".replace(",", "."), "border_color": "#10b981", "subtitle": "comunicação recente"},
        {"title": "WINDOWS 11", "value": f"{win11:,}".replace(",", "."), "border_color": "#8b5cf6", "subtitle": "estações"},
        {"title": "WINDOWS 10", "value": f"{win10:,}".replace(",", "."), "border_color": "#f59e0b", "subtitle": "estações"}
    ], cols=4)

    st.markdown("<br>", unsafe_allow_html=True)

    # Filtros e Busca
    col_s, col_f1, col_f2 = st.columns([2, 1, 1])
    with col_s:
        search_txt = st.text_input("🔎 Pesquisar Estação, Usuário ou IP:", placeholder="Ex: PGJ-NT-0123, paulo, 10.111...").strip().lower()
    with col_f1:
        filtro_so = st.selectbox("💻 Sistema Operacional:", ["Todos", "Windows 11", "Windows 10", "Windows Server"])
    with col_f2:
        apenas_ativos = st.checkbox("Apenas Clientes Ativos", value=False)

    filtered_df = df
    if not filtered_df.empty:
        if search_txt:
            filtered_df = filtered_df[
                filtered_df["name"].str.lower().str.contains(search_txt, na=False) |
                filtered_df["last_logon_user"].str.lower().str.contains(search_txt, na=False) |
                filtered_df["ip_addresses"].str.lower().str.contains(search_txt, na=False) |
                filtered_df["model"].str.lower().str.contains(search_txt, na=False)
            ]
        if filtro_so != "Todos":
            filtered_df = filtered_df[filtered_df["operating_system"].str.contains(filtro_so, case=False, na=False)]
        if apenas_ativos:
            filtered_df = filtered_df[filtered_df["client_active"] == 1]

    st.markdown(f"**Exibindo {len(filtered_df)} de {total} computadores.** *Clique em uma linha para abrir a Ficha Técnica e o Controle Remoto.*")

    if filtered_df.empty:
        st.info("Nenhum computador encontrado com os filtros aplicados. Clique em 'Sincronizar SCCM' na barra lateral para buscar dados frescos.")
        return

    # Tabela com seleção de linha para abertura do modal
    display_cols = ["name", "last_logon_user", "ip_addresses", "model", "operating_system", "client_active", "ad_site_name"]
    table_df = filtered_df[display_cols].copy()
    table_df.columns = ["Nome do Computador", "Último Usuário", "IP(s)", "Modelo", "Sistema Operacional", "Ativo", "Site AD"]

    event = st.dataframe(
        table_df,
        column_config={
            "Nome do Computador": st.column_config.TextColumn("Nome do Computador", width="medium"),
            "Último Usuário": st.column_config.TextColumn("Último Logon", width="small"),
            "IP(s)": st.column_config.TextColumn("Endereço IP", width="medium"),
            "Modelo": st.column_config.TextColumn("Modelo do Hardware", width="medium"),
            "Sistema Operacional": st.column_config.TextColumn("Sistema Operacional", width="medium"),
            "Ativo": st.column_config.CheckboxColumn("Ativo", width="small"),
            "Site AD": st.column_config.TextColumn("Site", width="small")
        },
        selection_mode="single-row",
        on_select="rerun",
        hide_index=True,
        width="stretch",
        key="sccm_devices_table"
    )

    if event and event.selection and event.selection.rows:
        sel_idx = event.selection.rows[0]
        device_data = filtered_df.iloc[sel_idx].to_dict()
        modal_device_details(device_data)


def render_subtab_usuarios():
    """Renderiza a listagem de usuários catalogados no SCCM."""
    df = get_sccm_users_df()
    st.markdown(f"### 👤 Usuários do SCCM ({len(df)} catalogados)")

    s_txt = st.text_input("🔎 Pesquisar Usuário por Login, Nome ou DN:", placeholder="Ex: paulo, silva, pgj...").strip().lower()
    filtered = df
    if s_txt and not filtered.empty:
        filtered = filtered[
            filtered["user_name"].str.lower().str.contains(s_txt, na=False) |
            filtered["full_user_name"].str.lower().str.contains(s_txt, na=False) |
            filtered["distinguished_name"].str.lower().str.contains(s_txt, na=False)
        ]

    if filtered.empty:
        st.info("Nenhum usuário localizado. Sincronize a base na barra lateral.")
        return

    cols = ["user_name", "full_user_name", "windows_nt_domain", "distinguished_name"]
    tbl = filtered[cols].copy()
    tbl.columns = ["Login", "Nome Completo", "Domínio", "Distinguished Name"]
    st.dataframe(tbl, hide_index=True, width="stretch")


def render_subtab_colecoes(col_type: str = "Dispositivos"):
    """Renderiza a listagem de coleções (Device Collections / User Collections)."""
    df = get_sccm_collections_df(col_type=col_type)
    st.markdown(f"### 📁 Coleções de {col_type} ({len(df)} coleções)")

    if df.empty:
        st.info(f"Nenhuma coleção de {col_type.lower()} encontrada. Use o botão de sincronização na barra lateral.")
        return

    cols = ["name", "member_count", "collection_id", "comment", "last_refresh_time"]
    tbl = df[cols].copy()
    tbl.columns = ["Nome da Coleção", "Qtd. Membros", "ID da Coleção", "Comentário", "Última Atualização"]
    st.dataframe(tbl, hide_index=True, width="stretch")


def render_subtab_conformidade():
    """Exibe painel de integridade dos agentes SCCM e políticas de conformidade."""
    df_dev = get_sccm_devices_df()
    st.markdown("### 🛡️ Configurações & Conformidade do Endpoint")

    if df_dev.empty:
        st.info("Sincronize o inventário do SCCM para carregar o diagnóstico de conformidade.")
        return

    total = len(df_dev)
    ativos = len(df_dev[df_dev["client_active"] == 1])
    inativos = total - ativos
    taxa = (ativos / total * 100) if total > 0 else 0

    render_metric_cards([
        {"title": "CONFORMIDADE DO AGENTE", "value": f"{taxa:.1f}%", "border_color": "#10b981", "subtitle": "agentes saudáveis"},
        {"title": "AGENTES EM DIA", "value": f"{ativos}", "border_color": "#0ea5e9", "subtitle": "respondendo ao site"},
        {"title": "AGENTES COM ALERTA", "value": f"{inativos}", "border_color": "#ef4444", "subtitle": "sem check-in recente"}
    ], cols=3)

    st.markdown("<br>", unsafe_allow_html=True)
    with st.expander("📋 Ver lista de estações que necessitam de intervenção (Inativos)"):
        df_inativos = df_dev[df_dev["client_active"] == 0]
        if not df_inativos.empty:
            cols = ["name", "last_logon_user", "ip_addresses", "model", "last_active_time"]
            t_in = df_inativos[cols].copy()
            t_in.columns = ["Computador", "Último Usuário", "IP", "Modelo", "Última Atividade"]
            st.dataframe(t_in, hide_index=True, width="stretch")
        else:
            st.success("🎉 Todas as estações gerenciadas estão ativas e em conformidade!")


def render_sccm_page():
    """Página principal do módulo Inventário SCCM."""
    setup_sccm_tables()

    # Cabeçalho com visual corporativo
    st.markdown("""
        <div style="background: var(--metric-bg, #1e293b); padding: 18px 24px; border-radius: 12px; border-left: 6px solid #0284c7; border-top: 1px solid var(--metric-border, #2d3139); border-right: 1px solid var(--metric-border, #2d3139); border-bottom: 1px solid var(--metric-border, #2d3139); margin-bottom: 20px;">
            <div style="display: flex; justify-content: space-between; align-items: center;">
                <h2 style="color: var(--metric-value-color, #ffffff); margin: 0; font-size: 24px; font-weight: 700;">💻 Inventário SCCM (Ativos e Conformidade)</h2>
                <span style="background-color: rgba(2, 132, 199, 0.15); color: #38bdf8; font-size: 13px; font-weight: 600; padding: 4px 12px; border-radius: 20px; border: 1px solid #0284c7;">
                    ⚡ WMI / CIM Integrado
                </span>
            </div>
            <p style="color: var(--metric-title-color, #94a3b8); margin: 6px 0 0 0; font-size: 14px;">
                Gestão unificada de computadores, coleções de dispositivos, usuários e controle remoto oficial (CmRcViewer) direto da bancada.
            </p>
        </div>
    """, unsafe_allow_html=True)

    # Sidebar com ações de sincronização do SCCM
    st.sidebar.markdown("## 🔄 Sincronização SCCM")
    
    # 1. Disparo de sincronização direta via bancada:// (executa com privilégios nativos no Windows)
    bancada_sync_url = "bancada://run?tool=sccm_sync&host=srv-1046.in.mpe.ms.gov.br"
    st.sidebar.markdown(f"""
    <a href="{bancada_sync_url}" style="text-decoration:none;">
        <button style="width:100%; height:42px; background-color:#0284c7; color:white; border:none; border-radius:6px; font-weight:600; cursor:pointer; margin-bottom:10px;">
            ⚡ Sincronizar via Bancada (Windows)
        </button>
    </a>
    """, unsafe_allow_html=True)

    if st.sidebar.button("📥 Importar Dados Coletados", use_container_width=True):
        with st.spinner("Importando arquivo de inventário sccm_inventory.json..."):
            from src.services.sccm_service import import_sccm_inventory_json
            res = import_sccm_inventory_json()
            if res["devices"] > 0 or res["collections"] > 0:
                st.toast(f"✅ Sucesso: {res['devices']} computadores, {res['users']} usuários e {res['collections']} coleções importados!", icon="🎉")
                st.rerun()
            else:
                st.sidebar.warning("⚠️ Nenhum arquivo novo encontrado. Dispare a sincronização acima primeiro.")

    # Sub-abas estilizadas com sincronização por Query Parameter (?subtab=slug)
    selected_title = render_subtabs(SCCM_SUBTABS, default_slug="dispositivos", key="sccm_subtabs_nav")

    st.markdown("<br>", unsafe_allow_html=True)

    if selected_title == "💻 Dispositivos":
        render_subtab_dispositivos()
    elif selected_title == "👤 Usuários":
        render_subtab_usuarios()
    elif selected_title == "📁 Coleções de Dispositivos":
        render_subtab_colecoes(col_type="Dispositivos")
    elif selected_title == "👥 Coleções de Usuários":
        render_subtab_colecoes(col_type="Usuários")
    elif selected_title == "🛡️ Configurações & Conformidade":
        render_subtab_conformidade()
