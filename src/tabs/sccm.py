import json
import streamlit as st
import pandas as pd
from datetime import datetime
from typing import Any, Optional
from src.components.subtabs import render_subtabs
from src.components.metric_cards import render_metric_cards
from src.components.pagination import (
    render_items_per_page_selector,
    paginate_items,
    render_pagination_controls
)
from src.database.sccm_db import (
    get_sccm_devices_df,
    get_sccm_users_df,
    get_sccm_collections_df,
    setup_sccm_tables,
    normalize_hardware_model,
    HARDWARE_MODEL_ALIASES
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


def parse_sccm_datetime(val: Any) -> Optional[datetime]:
    """
    Interpreta timestamps e strings de data do SCCM/WMI para datetime/Timestamp nativo.
    Suporta:
      - WMI CIM DateTime: '20260929080059.657000+***' ou '20260929080059'
      - ISO-8601: '2026-09-29T08:00:59' ou '2026-09-29 08:00:59'
      - Objetos datetime / Timestamp do Python
    Retorna None se nulo ou inválido.
    """
    if val is None or pd.isna(val):
        return None
    if isinstance(val, datetime) or (hasattr(pd, "Timestamp") and isinstance(val, getattr(pd, "Timestamp"))):
        return val

    raw = str(val).strip()
    if not raw or raw.lower() in ["none", "nan", "null", "nat", "-"]:
        return None

    # 1. Padrão WMI / CIM DateTime do SCCM (ex: 20260929080059.657000+***)
    # Primeiros 14 dígitos representam YYYYMMDDHHmmss
    clean_digits = raw.split(".")[0].strip()
    if len(clean_digits) == 14 and clean_digits.isdigit():
        try:
            return datetime.strptime(clean_digits, "%Y%m%d%H%M%S")
        except Exception:
            pass

    # 2. Padrões ISO e formatos comuns
    for fmt in [
        "%Y-%m-%dT%H:%M:%S",
        "%Y-%m-%d %H:%M:%S",
        "%Y-%m-%dT%H:%M:%S.%f",
        "%Y-%m-%d %H:%M:%S.%f",
        "%Y-%m-%d",
        "%d/%m/%Y %H:%M:%S",
        "%d/%m/%Y %H:%M",
        "%d/%m/%Y"
    ]:
        try:
            return datetime.strptime(raw[:19], fmt)
        except Exception:
            continue

    # 3. Tenta parsing flexível via pd.to_datetime
    try:
        dt = pd.to_datetime(raw, errors="coerce")
        if pd.notna(dt):
            return dt.to_pydatetime() if hasattr(dt, "to_pydatetime") else dt
    except Exception:
        pass

    return None


def format_sccm_datetime(val: Any) -> str:
    """
    Formata timestamps e strings de data do SCCM/WMI para o formato brasileiro DD/MM/AAAA HH:MM:SS.
    Suporta:
      - WMI CIM DateTime: '20260929080059.657000+***' ou '20260929080059'
      - ISO-8601: '2026-09-29T08:00:59' ou '2026-09-29 08:00:59'
      - Objetos datetime do Python ou timestamps
    """
    dt = parse_sccm_datetime(val)
    if dt is not None and pd.notna(dt):
        return dt.strftime("%d/%m/%Y %H:%M:%S")

    raw = str(val).strip() if val is not None and not pd.isna(val) else ""
    return "-" if not raw or raw.lower() in ["none", "nan", "null", "nat", "-"] else raw

@st.dialog("🔍 Ficha Técnica & Ações Rápidas do Computador", width="large")
def modal_device_details(device_row: dict):
    """Modal amplo do Streamlit exibindo especificações de hardware, sistema e ações remotas."""
    name = str(device_row.get("name") or "N/A")
    user = str(device_row.get("last_logon_user") or "Não identificado")
    ips = str(device_row.get("ip_addresses") or "N/A")
    macs = str(device_row.get("mac_addresses") or "N/A")
    model = str(device_row.get("model") or "").strip()
    mfg = str(device_row.get("manufacturer") or "").strip()
    proc = str(device_row.get("processor") or "").strip()
    ram = str(device_row.get("memory_ram") or "").strip()
    disks = str(device_row.get("disk_drives") or "").strip()
    os_name = str(device_row.get("operating_system") or "N/A")
    os_ver = str(device_row.get("os_version") or "N/A")
    client_ver = str(device_row.get("client_version") or "N/A")
    client_act = bool(device_row.get("client_active", 1))
    site = str(device_row.get("ad_site_name") or "N/A")
    dn = str(device_row.get("distinguished_name") or "")
    raw_active = device_row.get("last_active_time")
    last_active = format_sccm_datetime(raw_active) if raw_active else "N/A"
    res_id = str(device_row.get("resource_id") or "N/A")

    # Extração defensiva a partir de raw_json se colunas não vierem preenchidas diretamente
    raw_str = device_row.get("raw_json")
    raw_data = {}
    if raw_str:
        try:
            raw_data = json.loads(raw_str) if isinstance(raw_str, str) else raw_str
        except Exception:
            pass

    if not mfg or mfg.lower() in ["none", "nan", "null"]:
        mfg = str(raw_data.get("Manufacturer") or raw_data.get("manufacturer") or "").strip()
    if not model or model.lower() in ["none", "nan", "null"]:
        model = str(raw_data.get("Model") or raw_data.get("model") or "").strip()
    
    # Aplica normalização por aliases caso o modelo seja um código MTM/técnico
    commercial_model = normalize_hardware_model(model)
    model_display = commercial_model
    if commercial_model and commercial_model != model:
        model_display = f"{commercial_model} <span style='font-size:12px; color:#94a3b8;'>({model})</span>"
    if not proc or proc.lower() in ["none", "nan", "null"]:
        proc = str(raw_data.get("Processor") or raw_data.get("processor") or raw_data.get("CPU") or "").strip()
    if not ram or ram.lower() in ["none", "nan", "null"]:
        raw_ram = raw_data.get("MemoryRAM") or raw_data.get("memory_ram") or raw_data.get("TotalPhysicalMemory") or ""
        if isinstance(raw_ram, (int, float)) and raw_ram > 0:
            if raw_ram > 1024 * 1024 * 1024:
                ram = f"{round(raw_ram / (1024**3))} GB"
            elif raw_ram > 1024 * 1024:
                ram = f"{round(raw_ram / (1024**2))} GB"
            else:
                ram = f"{raw_ram} MB"
        else:
            ram = str(raw_ram).strip()
    if not disks or disks.lower() in ["none", "nan", "null"]:
        disks = str(raw_data.get("DiskDrives") or raw_data.get("disk_drives") or raw_data.get("Disks") or "").strip()

    status_badge = "🟢 <span style='color:#10b981;font-weight:600;'>Cliente SCCM Ativo</span>" if client_act else "🔴 <span style='color:#ef4444;font-weight:600;'>Sem Comunicação Recente</span>"

    # --- BANNER DO EQUIPAMENTO ---
    st.markdown(f"""
        <div style="background: rgba(14, 165, 233, 0.08); border-left: 5px solid #0284c7; padding: 14px 20px; border-radius: 8px; margin-bottom: 16px;">
            <div style="display: flex; justify-content: space-between; align-items: center; flex-wrap: wrap; gap: 8px;">
                <div>
                    <h2 style="margin: 0; font-size: 22px; color: var(--text-color, #ffffff); font-weight: 700;">🖥️ {name}</h2>
                    <span style="font-size: 13px; color: #94a3b8;">Último Usuário: <b style="color: #38bdf8;">{user}</b> • Site AD: <b>{site}</b></span>
                </div>
                <div>
                    {status_badge}
                </div>
            </div>
        </div>
    """, unsafe_allow_html=True)

    # --- PAINEL DE DISPARO RÁPIDO VIA BANCADA:// ---
    st.markdown("##### ⚡ Ações Remotas Instantâneas")
    col1, col2, col3, col4, col5 = st.columns(5)

    with col1:
        cmrc_url = f"bancada://run?tool=cmrc&host={name}"
        st.link_button("🎮 Controle Remoto", url=cmrc_url, use_container_width=True, help="Abre o CmRcViewer nativo do SCCM.")

    with col2:
        rdp_url = f"bancada://run?tool=rdp&host={name}"
        st.link_button("🖥️ Conexão RDP", url=rdp_url, use_container_width=True, help="Inicia MSTSC (Área de Trabalho Remota).")

    with col3:
        exp_url = f"bancada://run?tool=explorer&host={name}"
        st.link_button("📂 Explorer C$", url=exp_url, use_container_width=True, help=f"Abre o compartilhamento administrativo \\\\{name}\\c$")

    with col4:
        ping_url = f"bancada://run?tool=ping&host={name}"
        st.link_button("⚡ Teste Ping", url=ping_url, use_container_width=True, help="Dispara ping contínuo na estação.")

    with col5:
        analisador_url = f"bancada://run?tool=analisador&host={name}"
        st.link_button("🛠️ Analisador", url=analisador_url, use_container_width=True, help="Executa o diagnóstico profundo da Bancada STI na máquina.")

    st.markdown("---")

    # --- INFORMAÇÕES TÉCNICAS DETALHADAS EM 3 COLUNAS ---
    col_hw, col_so, col_net = st.columns(3)

    with col_hw:
        st.markdown("#### ⚙️ Hardware & Peças")
        st.markdown(f"**Fabricante:** {mfg or 'Não informado no SCCM'}")
        if model_display:
            st.markdown(f"**Modelo:** {model_display}", unsafe_allow_html=True)
        else:
            st.markdown("**Modelo:** Não informado no SCCM")
        st.markdown(f"**Processador (CPU):** {proc or 'Pendente de sincronização'}")
        st.markdown(f"**Memória RAM:** {ram or 'Pendente de sincronização'}")
        st.markdown(f"**Armazenamento / Discos:** {disks or 'Pendente de sincronização'}")

    with col_so:
        st.markdown("#### 💻 Sistema & Agente")
        st.markdown(f"**Sistema Operacional:** {os_name}")
        st.markdown(f"**Versão de Build:** `{os_ver}`")
        st.markdown(f"**Versão do Cliente SCCM:** `{client_ver}`")
        st.markdown(f"**Último Check-in:** {last_active}")
        st.markdown(f"**Resource ID:** `{res_id}`")

    with col_net:
        st.markdown("#### 🌐 Rede & Domínio")
        st.markdown(f"**Endereço(s) IP:** `{ips}`")
        st.markdown(f"**MAC Address:** `{macs}`")
        st.markdown(f"**Site Active Directory:** `{site}`")
        st.markdown(f"**Domínio:** `MPE (mpe.ms.gov.br)`")

    if dn and dn != "N/A":
        st.markdown("<br>", unsafe_allow_html=True)
        st.caption(f"🌳 **Distinguished Name (DN):** `{dn}`")

    # Atalhos cruzados com o Active Directory
    st.markdown("---")
    c_ad1, c_ad2 = st.columns(2)
    with c_ad1:
        ad_comp_url = f"?tab=active-directory&subtab=computadores&search={name}"
        st.link_button("🌳 Ver Conta da Máquina no AD ↗", url=ad_comp_url, use_container_width=True, help="Abre a ficha desta máquina no Active Directory.")
    with c_ad2:
        if user and user != "Não identificado":
            clean_u = user.split("\\")[-1] if "\\" in user else user
            ad_usr_url = f"?tab=active-directory&subtab=usuarios&search={clean_u}"
            st.link_button(f"👤 Ver Usuário ({clean_u}) no AD ↗", url=ad_usr_url, use_container_width=True, help="Abre a ficha cadastral do usuário no Active Directory.")
        else:
            st.button("👤 Usuário não identificado", disabled=True, use_container_width=True)

    # Exibe JSON bruto em expander caso queira auditar propriedades adicionais
    if raw_str:
        with st.expander("📄 Ver Dados Brutos WMI/CIM do SCCM"):
            try:
                st.json(raw_data if raw_data else json.loads(raw_str))
            except Exception:
                st.code(raw_str)

    st.markdown("<br>", unsafe_allow_html=True)
    if st.button("Fechar Ficha Técnica", key="close_sccm_device_dialog_btn", use_container_width=True):
        st.session_state["last_selected_sccm_dev"] = None
        st.rerun()


def render_subtab_dispositivos(
    search_txt: str = "",
    filtro_so: str = "Todos",
    filtro_modelo: str = "Todos",
    apenas_ativos: bool = False,
    items_per_page: int = 50
):
    """Renderiza a listagem de computadores com filtros da barra lateral, paginação e métricas."""
    df = get_sccm_devices_df()

    # Métricas KPI no topo com SOs normalizados
    total = len(df)
    win11 = len(df[df["operating_system"].str.contains("Windows 11", case=False, na=False)]) if not df.empty else 0
    win10 = len(df[df["operating_system"].str.contains("Windows 10", case=False, na=False)]) if not df.empty else 0
    ativos = len(df[df["client_active"] == 1]) if not df.empty else 0

    render_metric_cards([
        {"title": "TOTAL DE COMPUTADORES", "value": f"{total:,}".replace(",", "."), "border_color": "#0ea5e9", "subtitle": "no inventário SCCM"},
        {"title": "CLIENTES ATIVOS", "value": f"{ativos:,}".replace(",", "."), "border_color": "#10b981", "subtitle": "comunicação recente"},
        {"title": "WINDOWS 11", "value": f"{win11:,}".replace(",", "."), "border_color": "#8b5cf6", "subtitle": "estações corporativas"},
        {"title": "WINDOWS 10", "value": f"{win10:,}".replace(",", "."), "border_color": "#f59e0b", "subtitle": "estações corporativas"}
    ], cols=4)

    st.markdown("<br>", unsafe_allow_html=True)

    # Aplicação de Filtros oriundos do Sidebar
    filtered_df = df
    if not filtered_df.empty:
        if search_txt:
            # Busca unificada inteligente (Nome, Usuário, IP, Modelo de hardware, Fabricante e aliases)
            filtered_df = filtered_df[
                filtered_df["name"].str.lower().str.contains(search_txt, na=False) |
                filtered_df["last_logon_user"].str.lower().str.contains(search_txt, na=False) |
                filtered_df["ip_addresses"].str.lower().str.contains(search_txt, na=False) |
                filtered_df["model"].str.lower().str.contains(search_txt, na=False) |
                filtered_df["manufacturer"].str.lower().str.contains(search_txt, na=False) |
                filtered_df["raw_json"].str.lower().str.contains(search_txt, na=False)
            ]
        if filtro_so != "Todos":
            filtered_df = filtered_df[filtered_df["operating_system"].str.contains(filtro_so, case=False, na=False)]
        if filtro_modelo != "Todos":
            filtered_df = filtered_df[filtered_df["model"].str.contains(filtro_modelo, case=False, na=False, regex=False)]
        if apenas_ativos:
            filtered_df = filtered_df[filtered_df["client_active"] == 1]

    if filtered_df.empty:
        st.info("Nenhum computador encontrado com os filtros aplicados na barra lateral. Clique em 'Sincronizar SCCM' para buscar dados frescos.")
        return

    # Paginação dos itens filtrados
    page_df, current_page, total_pages, total_items = paginate_items(
        filtered_df,
        page_key="sccm_dev",
        items_per_page=items_per_page
    )

    display_cols = ["name", "last_logon_user", "ip_addresses", "model", "operating_system", "client_active", "ad_site_name"]
    table_df = page_df[display_cols].copy()
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
        key=f"sccm_devices_table_p{current_page}"
    )

    if "last_selected_sccm_dev" not in st.session_state:
        st.session_state["last_selected_sccm_dev"] = None

    if event and event.selection and event.selection.rows:
        sel_idx = event.selection.rows[0]
        if st.session_state["last_selected_sccm_dev"] != sel_idx:
            st.session_state["last_selected_sccm_dev"] = sel_idx
            sel_row = page_df.iloc[sel_idx]
            device_data = sel_row.to_dict() if hasattr(sel_row, "to_dict") else dict(sel_row)
            modal_device_details(device_data)
    else:
        st.session_state["last_selected_sccm_dev"] = None

    render_pagination_controls(
        page_key="sccm_dev",
        current_page=current_page,
        total_pages=total_pages,
        total_items=total_items,
        items_per_page=items_per_page
    )


def render_subtab_usuarios(search_txt: str = "", items_per_page: int = 50):
    """Renderiza a listagem de usuários catalogados no SCCM com filtros da barra lateral e paginação."""
    df = get_sccm_users_df()
    total_users = len(df)

    render_metric_cards([
        {"title": "TOTAL DE USUÁRIOS", "value": f"{total_users:,}".replace(",", "."), "border_color": "#0ea5e9", "subtitle": "catalogados no SCCM"},
        {"title": "DOMÍNIO CORPORATIVO", "value": "MPE", "border_color": "#10b981", "subtitle": "mpe.ms.gov.br"}
    ], cols=2)

    st.markdown("<br>", unsafe_allow_html=True)

    filtered = df
    if search_txt and not filtered.empty:
        filtered = filtered[
            filtered["user_name"].str.lower().str.contains(search_txt, na=False) |
            filtered["full_user_name"].str.lower().str.contains(search_txt, na=False) |
            filtered["distinguished_name"].str.lower().str.contains(search_txt, na=False)
        ]

    if filtered.empty:
        st.info("Nenhum usuário localizado com os critérios da barra lateral. Sincronize a base na barra lateral se necessário.")
        return

    page_df, current_page, total_pages, total_items = paginate_items(
        filtered,
        page_key="sccm_usr",
        items_per_page=items_per_page
    )

    cols = ["user_name", "full_user_name", "windows_nt_domain", "distinguished_name"]
    tbl = page_df[cols].copy()
    tbl.columns = ["Login", "Nome Completo", "Domínio", "Distinguished Name"]
    st.dataframe(tbl, hide_index=True, width="stretch", key=f"sccm_users_table_p{current_page}")

    render_pagination_controls(
        page_key="sccm_usr",
        current_page=current_page,
        total_pages=total_pages,
        total_items=total_items,
        items_per_page=items_per_page
    )


def render_subtab_colecoes(
    col_type: str = "Dispositivos",
    search_txt: str = "",
    apenas_com_membros: bool = False,
    items_per_page: int = 50
):
    """Renderiza a listagem de coleções com filtros da barra lateral e paginação."""
    df = get_sccm_collections_df(col_type=col_type)
    total_cols = len(df)
    total_membros = int(df["member_count"].sum()) if not df.empty and "member_count" in df.columns else 0

    render_metric_cards([
        {"title": f"COLEÇÕES DE {col_type.upper()}", "value": f"{total_cols:,}".replace(",", "."), "border_color": "#0ea5e9", "subtitle": "ativas no SCCM"},
        {"title": "TOTAL DE MEMBROS ALOCADOS", "value": f"{total_membros:,}".replace(",", "."), "border_color": "#8b5cf6", "subtitle": "somatória de membros"}
    ], cols=2)

    st.markdown("<br>", unsafe_allow_html=True)

    filtered = df
    if not filtered.empty:
        if search_txt:
            filtered = filtered[
                filtered["name"].str.lower().str.contains(search_txt, na=False) |
                filtered["collection_id"].str.lower().str.contains(search_txt, na=False) |
                filtered["comment"].str.lower().str.contains(search_txt, na=False)
            ]
        if apenas_com_membros and "member_count" in filtered.columns:
            filtered = filtered[filtered["member_count"] > 0]

    if filtered.empty:
        st.info(f"Nenhuma coleção de {col_type.lower()} encontrada. Use o botão de sincronização na barra lateral.")
        return

    pkey = f"sccm_col_{col_type.lower()}"
    page_df, current_page, total_pages, total_items = paginate_items(
        filtered,
        page_key=pkey,
        items_per_page=items_per_page
    )

    cols = ["name", "member_count", "collection_id", "comment", "last_refresh_time"]
    tbl = page_df[cols].copy()
    if "last_refresh_time" in tbl.columns:
        tbl["last_refresh_time"] = tbl["last_refresh_time"].apply(parse_sccm_datetime)
        tbl["last_refresh_time"] = pd.to_datetime(tbl["last_refresh_time"])
    tbl.columns = ["Nome da Coleção", "Qtd. Membros", "ID da Coleção", "Comentário", "Última Atualização"]
    st.dataframe(
        tbl,
        column_config={
            "Nome da Coleção": st.column_config.TextColumn("Nome da Coleção"),
            "Qtd. Membros": st.column_config.NumberColumn("Qtd. Membros"),
            "ID da Coleção": st.column_config.TextColumn("ID da Coleção"),
            "Comentário": st.column_config.TextColumn("Comentário"),
            "Última Atualização": st.column_config.DatetimeColumn("Última Atualização", format="DD/MM/YYYY HH:mm:ss")
        },
        hide_index=True,
        width="stretch",
        key=f"sccm_col_{col_type.lower()}_table_p{current_page}"
    )

    render_pagination_controls(
        page_key=pkey,
        current_page=current_page,
        total_pages=total_pages,
        total_items=total_items,
        items_per_page=items_per_page
    )


def render_subtab_conformidade(filtro_status: str = "Todos", items_per_page: int = 50):
    """Exibe painel de integridade dos agentes SCCM e políticas de conformidade com paginação."""
    df_dev = get_sccm_devices_df()

    if df_dev.empty:
        st.info("Sincronize o inventário do SCCM para carregar o diagnóstico de conformidade.")
        return

    total = len(df_dev)
    ativos = len(df_dev[df_dev["client_active"] == 1])
    inativos = total - ativos
    taxa = (ativos / total * 100) if total > 0 else 0

    render_metric_cards([
        {"title": "CONFORMIDADE DO AGENTE", "value": f"{taxa:.1f}%", "border_color": "#10b981", "subtitle": "agentes saudáveis"},
        {"title": "AGENTES EM DIA", "value": f"{ativos:,}".replace(",", "."), "border_color": "#0ea5e9", "subtitle": "respondendo ao site"},
        {"title": "AGENTES COM ALERTA", "value": f"{inativos:,}".replace(",", "."), "border_color": "#ef4444", "subtitle": "sem check-in recente"}
    ], cols=3)

    st.markdown("<br>", unsafe_allow_html=True)

    # Filtragem com base na sidebar
    df_display = df_dev
    if filtro_status == "Apenas Ativos":
        df_display = df_dev[df_dev["client_active"] == 1]
    elif filtro_status == "Apenas Inativos (Alerta)":
        df_display = df_dev[df_dev["client_active"] == 0]

    page_df, current_page, total_pages, total_items = paginate_items(
        df_display,
        page_key="sccm_conf",
        items_per_page=items_per_page
    )

    cols = ["name", "last_logon_user", "ip_addresses", "model", "operating_system", "last_active_time"]
    t_disp = page_df[cols].copy()
    if "last_active_time" in t_disp.columns:
        t_disp["last_active_time"] = t_disp["last_active_time"].apply(parse_sccm_datetime)
        t_disp["last_active_time"] = pd.to_datetime(t_disp["last_active_time"])
    t_disp.columns = ["Computador", "Último Usuário", "IP", "Modelo", "Sistema Operacional", "Última Atividade"]

    st.dataframe(
        t_disp,
        column_config={
            "Computador": st.column_config.TextColumn("Computador"),
            "Último Usuário": st.column_config.TextColumn("Último Usuário"),
            "IP": st.column_config.TextColumn("IP"),
            "Modelo": st.column_config.TextColumn("Modelo"),
            "Sistema Operacional": st.column_config.TextColumn("Sistema Operacional"),
            "Última Atividade": st.column_config.DatetimeColumn("Última Atividade", format="DD/MM/YYYY HH:mm:ss")
        },
        hide_index=True,
        width="stretch",
        key=f"sccm_conf_table_p{current_page}"
    )

    render_pagination_controls(
        page_key="sccm_conf",
        current_page=current_page,
        total_pages=total_pages,
        total_items=total_items,
        items_per_page=items_per_page
    )


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

    # Sub-abas estilizadas com sincronização por Query Parameter (?subtab=slug)
    selected_title = render_subtabs(SCCM_SUBTABS, default_slug="dispositivos", key="sccm_subtabs_nav")

    # =========================================================================
    # BARRA LATERAL (st.sidebar): FILTROS CONTEXTUAIS DINÂMICOS & PAGINAÇÃO
    # =========================================================================
    st.sidebar.markdown("## 🔍 Filtros & Pesquisa")

    dev_search = ""
    dev_filtro_so = "Todos"
    dev_filtro_modelo = "Todos"
    dev_apenas_ativos = False
    dev_per_page = 50

    usr_search = ""
    usr_per_page = 50

    col_search = ""
    col_has_members = False
    col_per_page = 50

    conf_status = "Todos"
    conf_per_page = 50

    if selected_title == "💻 Dispositivos":
        default_sccm_search = st.query_params.get("search", "")
        dev_search = st.sidebar.text_input(
            "🔎 Buscar Estação, Usuário, IP ou Modelo:",
            value=default_sccm_search,
            placeholder="Ex: PGJ-NT-0123, paulo, ThinkCentre...",
            key="sccm_dev_search"
        ).strip().lower()

        dev_filtro_so = st.sidebar.selectbox(
            "💻 Sistema Operacional:",
            ["Todos", "Windows 11", "Windows 10", "Windows Server"],
            key="sccm_dev_so"
        )

        # Carrega lista dinâmica de modelos existentes no cache do SCCM
        try:
            df_devs_all = get_sccm_devices_df()
            if not df_devs_all.empty and "model" in df_devs_all.columns:
                raw_models = df_devs_all["model"].dropna().astype(str).str.strip()
                valid_models = sorted(list(set([m for m in raw_models if m and m.lower() not in ["none", "nan", "null", ""]])))
                model_options = ["Todos"] + valid_models
            else:
                model_options = ["Todos"]
        except Exception:
            model_options = ["Todos"]

        dev_filtro_modelo = st.sidebar.selectbox(
            "⚙️ Modelo de Hardware:",
            model_options,
            key="sccm_dev_modelo",
            help="Filtra computadores por modelo comercial normalizado ou código de fábrica."
        )

        dev_apenas_ativos = st.sidebar.checkbox(
            "Apenas Clientes Ativos",
            value=False,
            key="sccm_dev_ativos"
        )

        dev_per_page = render_items_per_page_selector(
            key_prefix="sccm_dev",
            options=[10, 25, 50, 100, "Todos"],
            default_index=2,
            label="📄 Computadores por página:"
        )

    elif selected_title == "👤 Usuários":
        usr_search = st.sidebar.text_input(
            "🔎 Buscar Login, Nome ou DN:",
            placeholder="Ex: paulo, silva, pgj...",
            key="sccm_user_search"
        ).strip().lower()

        usr_per_page = render_items_per_page_selector(
            key_prefix="sccm_usr",
            options=[10, 25, 50, 100, "Todos"],
            default_index=2,
            label="📄 Usuários por página:"
        )

    elif selected_title in ["📁 Coleções de Dispositivos", "👥 Coleções de Usuários"]:
        tipo_lbl = "Dispositivos" if "Dispositivos" in selected_title else "Usuários"
        col_search = st.sidebar.text_input(
            f"🔎 Buscar Coleção ({tipo_lbl}):",
            placeholder="Ex: Todos os Sistemas, Windows...",
            key=f"sccm_col_search_{tipo_lbl}"
        ).strip().lower()

        col_has_members = st.sidebar.checkbox(
            "Apenas com Membros (> 0)",
            value=False,
            key=f"sccm_col_members_{tipo_lbl}"
        )

        col_per_page = render_items_per_page_selector(
            key_prefix=f"sccm_col_{tipo_lbl.lower()}",
            options=[10, 25, 50, 100, "Todos"],
            default_index=2,
            label="📄 Coleções por página:"
        )

    elif selected_title == "🛡️ Configurações & Conformidade":
        conf_status = st.sidebar.radio(
            "Status do Agente:",
            ["Todos", "Apenas Ativos", "Apenas Inativos (Alerta)"],
            key="sccm_conf_status"
        )

        conf_per_page = render_items_per_page_selector(
            key_prefix="sccm_conf",
            options=[10, 25, 50, 100, "Todos"],
            default_index=2,
            label="📄 Estações por página:"
        )

    st.sidebar.markdown("---")
    st.sidebar.markdown("## 🏢 Relatórios Especiais")
    from src.components.dmp_patrimonio_report import modal_dmp_patrimonio_report
    if st.sidebar.button("📋 Localizar Patrimônios (DMP)", use_container_width=True, help="Cruza em lote números de patrimônio com dados de usuário, modelo e localização no SCCM e AD."):
        modal_dmp_patrimonio_report()

    st.sidebar.markdown("---")
    st.sidebar.markdown("## 🔄 Sincronização SCCM")

    # Disparo de sincronização direta via bancada:// (executa com privilégios nativos no Windows)
    bancada_sync_url = "bancada://run?tool=sccm_sync&host=srv-1046.in.mpe.ms.gov.br"
    st.sidebar.link_button(
        "⚡ Sincronizar via Bancada (Windows)",
        url=bancada_sync_url,
        type="primary",
        use_container_width=True,
        help="Dispara a coleta do inventário SCCM via WMI/CIM com privilégios locais do Windows."
    )

    if st.sidebar.button("📥 Importar Dados Coletados", use_container_width=True):
        with st.spinner("Importando arquivo de inventário sccm_inventory.json..."):
            from src.services.sccm_service import import_sccm_inventory_json
            res = import_sccm_inventory_json()
            if res["devices"] > 0 or res["collections"] > 0:
                st.toast(f"✅ Sucesso: {res['devices']} computadores, {res['users']} usuários e {res['collections']} coleções importados!", icon="🎉")
                st.rerun()
            else:
                st.sidebar.warning("⚠️ Nenhum arquivo novo encontrado. Dispare a sincronização acima primeiro.")

    st.markdown("<br>", unsafe_allow_html=True)

    # Roteamento das Sub-abas na Área Principal com suporte a paginação
    if selected_title == "💻 Dispositivos":
        render_subtab_dispositivos(
            search_txt=dev_search,
            filtro_so=dev_filtro_so,
            filtro_modelo=dev_filtro_modelo,
            apenas_ativos=dev_apenas_ativos,
            items_per_page=dev_per_page
        )
    elif selected_title == "👤 Usuários":
        render_subtab_usuarios(
            search_txt=usr_search,
            items_per_page=usr_per_page
        )
    elif selected_title == "📁 Coleções de Dispositivos":
        render_subtab_colecoes(
            col_type="Dispositivos",
            search_txt=col_search,
            apenas_com_membros=col_has_members,
            items_per_page=col_per_page
        )
    elif selected_title == "👥 Coleções de Usuários":
        render_subtab_colecoes(
            col_type="Usuários",
            search_txt=col_search,
            apenas_com_membros=col_has_members,
            items_per_page=col_per_page
        )
    elif selected_title == "🛡️ Configurações & Conformidade":
        render_subtab_conformidade(
            filtro_status=conf_status,
            items_per_page=conf_per_page
        )
