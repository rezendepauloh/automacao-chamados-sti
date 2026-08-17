import io
import re
import sys
import platform
import subprocess
import pandas as pd
import streamlit as st
from datetime import datetime
from src.database import get_impressoras_df
from src.scrapers.papercut_scraper import check_papercut_sync_running, read_papercut_last_log_lines, run_papercut_scraper
from src.components.status_banner import render_log_expander
from src.components.pagination import (
    render_items_per_page_selector,
    paginate_items,
    render_pagination_controls
)

def get_printer_url(ip_raw: str) -> str:
    """Retorna URL formatada caso o valor seja um endereço IPv4 válido."""
    if not ip_raw or pd.isna(ip_raw):
        return ""
    s = str(ip_raw).strip()
    if re.match(r"^(\d{1,3}\.){3}\d{1,3}$", s):
        return f"https://{s}"
    elif s.startswith("http://") or s.startswith("https://"):
        return s
    return ""


def ping_host(host: str, count: int = 4, timeout_ms: int = 1000) -> tuple[bool, str]:
    """Executa um ping no host informado (IP ou Hostname) e retorna status (bool) e a saída do terminal (str)."""
    clean_host = str(host).replace("https://", "").replace("http://", "").strip().split("/")[0]
    if not clean_host:
        return False, "Host inválido."

    param = "-n" if platform.system().lower() == "windows" else "-c"
    timeout_param = ["-w", str(timeout_ms)] if platform.system().lower() == "windows" else ["-W", "1"]
    
    command = ["ping", param, str(count)] + timeout_param + [clean_host]

    try:
        output = subprocess.check_output(command, stderr=subprocess.STDOUT, universal_newlines=True, timeout=6)
        is_success = ("0% loss" in output or "0% de perda" in output or "bytes=" in output.lower())
        return is_success, output
    except subprocess.CalledProcessError as e:
        return False, e.output if e.output else "Host inalcançável (Timeout/sem resposta)."
    except Exception as ex:
        return False, f"Erro ao executar o ping: {str(ex)}"


@st.dialog("🖨️ Ficha Técnica do Ativo / Impressora", width="medium")
def show_printer_details(row_data):
    nome = row_data.get('Nome / Ativo', row_data.get('nome', 'N/A'))
    tipo = row_data.get('Tipo de Ativo', row_data.get('tipo', 'N/A'))
    status = row_data.get('Status', row_data.get('status', 'N/A'))
    modelo = row_data.get('Modelo / Fabricante', row_data.get('modelo', 'N/A'))
    local = row_data.get('Localização', row_data.get('localizacao', 'N/A'))
    ip_host = str(row_data.get('IP / Hostname', row_data.get('ip_host', 'N/A')))
    servidor = row_data.get('Servidor', row_data.get('servidor', 'N/A'))
    total_paginas = row_data.get('Total Páginas Impressas', row_data.get('total_paginas', 0))

    st.markdown(f"### 🖨️ {nome}")
    st.markdown("---")

    c1, c2 = st.columns(2)
    with c1:
        st.write(f"**📌 Tipo de Ativo:** {tipo}")
        st.write(f"**🏢 Localização:** {local}")
        st.write(f"**🖥️ Servidor:** {servidor}")
    with c2:
        st.write(f"**🟢 Status:** {status}")
        st.write(f"**⚙️ Modelo:** {modelo}")
        st.write(f"**📊 Páginas Impressas:** {total_paginas}")

    st.markdown("---")
    
    url_web = get_printer_url(ip_host)
    clean_ip = ip_host.replace('https://', '').replace('http://', '').strip().split('/')[0] if ip_host and ip_host != 'N/A' else ""

    if clean_ip and (re.match(r"^(\d{1,3}\.){3}\d{1,3}$", clean_ip) or url_web):
        st.success(f"🌐 **Endereço IP / Host:** `{clean_ip}`")
        
        c_web, c_ping = st.columns([1.2, 1])
        with c_web:
            if url_web:
                st.link_button(
                    label="🌐 Interface Web ↗",
                    url=url_web,
                    type="primary",
                    width='stretch'
                )
        with c_ping:
            do_ping = st.button("📡 Testar Ping", width='stretch', key=f"btn_ping_{clean_ip}")

        if do_ping:
            with st.spinner(f"Disparando 4 pacotes de ping para {clean_ip}..."):
                is_online, ping_output = ping_host(clean_ip)
                if is_online:
                    st.toast(f"✅ Impressora {clean_ip} está ONLINE!", icon="📶")
                else:
                    st.toast(f"⚠️ Impressora {clean_ip} não respondeu ao ping!", icon="❌")
                
                with st.expander("📶 Resultado Detalhado do Ping (Console)", expanded=True):
                    if is_online:
                        st.success("🟢 **Status:** ONLINE / Alcançável")
                    else:
                        st.error("🔴 **Status:** OFFLINE / Inalcançável")
                    st.code(ping_output, language="text")
    else:
        st.info(f"🌐 **IP / Hostname:** `{ip_host if ip_host else 'Não cadastrado'}`")
        st.caption("⚠️ Este dispositivo não possui endereço IP IPv4 configurado para teste de conectividade.")


def render_impressoras_page():
    """
    Renderiza a página principal de gestão de impressoras e dispositivos do PaperCut.
    """
    st.title("🖨️ Gestão de Impressoras & Dispositivos (PaperCut)")
    st.caption("Visualização unificada e controle de filas de impressão e dispositivos multifuncionais (MFDs).")

    papercut_ativo = check_papercut_sync_running()

    render_log_expander(
        "🤖 Robô do PaperCut Rodando em Segundo Plano – Acompanhar Progresso",
        papercut_ativo,
        read_papercut_last_log_lines,
        check_papercut_sync_running,
        "O robô está conectando ao PaperCut, baixando e cruzando os relatórios neste momento. O painel permanece livre para uso!"
    )

    df = get_impressoras_df()

    if df.empty:
        st.info("Nenhuma impressora encontrada no banco de dados local. Utilize o botão '🔄 Sincronizar Impressoras' na barra lateral para iniciar a coleta.")
        return

    # -----------------------------------------------------------------------------
    # FILTROS SIDEBAR & AÇÕES
    # -----------------------------------------------------------------------------
    with st.sidebar:
        st.markdown("## ⚙️ Ações e Coleta")
        if papercut_ativo:
            st.sidebar.button("🤖 Sincronizando PaperCut...", width='stretch', disabled=True)
        else:
            if st.sidebar.button("🔄 Sincronizar Impressoras", type="primary", width='stretch', help="Executa a coleta e unificação de dados do PaperCut em segundo plano."):
                import sys, subprocess, time
                creationflags = subprocess.CREATE_NO_WINDOW if sys.platform == "win32" else 0
                subprocess.Popen([sys.executable, "src/scrapers/papercut_scraper.py"], creationflags=creationflags)
                time.sleep(1.0)
                st.toast("🚀 Sincronização do PaperCut iniciada em segundo plano!", icon="🤖")
                st.rerun()

        st.markdown("---")
        st.markdown("## 🔍 Filtros de Impressoras")
        
        search_query = st.text_input("🔎 Buscar (Nome, Servidor, Modelo, IP)", "").strip().lower()

        servidores_disponiveis = ["Todos"] + sorted([s for s in df['servidor'].dropna().unique() if str(s).strip()])
        selected_servidor = st.selectbox("Servidor de Impressão", servidores_disponiveis)

        tipos_disponiveis = ["Todos"] + sorted(list(df['tipo'].dropna().unique()))
        selected_tipo = st.selectbox("Tipo de Ativo", tipos_disponiveis)

        status_disponiveis = ["Todos"] + sorted(list(df['status'].dropna().unique()))
        selected_status = st.selectbox("Status", status_disponiveis)

        locais_disponiveis = sorted(list(df['localizacao'].dropna().unique()))
        selected_locais = st.multiselect("Localização", locais_disponiveis)

        modelos_disponiveis = sorted(list(df['modelo'].dropna().unique()))
        selected_modelos = st.multiselect("Fabricante / Modelo", modelos_disponiveis)

        items_per_page = render_items_per_page_selector(
            key_prefix="impressoras",
            options=[10, 25, 50, 100, 200, "Todos"],
            default_index=2,
            label="📄 Ativos por página:"
        )

    # -----------------------------------------------------------------------------
    # APLICAÇÃO DOS FILTROS
    # -----------------------------------------------------------------------------
    df_filtered = df.copy()

    if search_query:
        mask = (
            df_filtered['nome'].str.lower().str.contains(search_query, na=False) |
            df_filtered['servidor'].str.lower().str.contains(search_query, na=False) |
            df_filtered['modelo'].str.lower().str.contains(search_query, na=False) |
            df_filtered['localizacao'].str.lower().str.contains(search_query, na=False) |
            df_filtered['ip_host'].str.lower().str.contains(search_query, na=False)
        )
        df_filtered = df_filtered[mask]

    if selected_servidor != "Todos":
        df_filtered = df_filtered[df_filtered['servidor'] == selected_servidor]

    if selected_tipo != "Todos":
        df_filtered = df_filtered[df_filtered['tipo'] == selected_tipo]

    if selected_status != "Todos":
        df_filtered = df_filtered[df_filtered['status'] == selected_status]

    if selected_locais:
        df_filtered = df_filtered[df_filtered['localizacao'].isin(selected_locais)]

    if selected_modelos:
        df_filtered = df_filtered[df_filtered['modelo'].isin(selected_modelos)]


    # -----------------------------------------------------------------------------
    # CARDS KPIS
    # -----------------------------------------------------------------------------
    col1, col2, col3, col4, col5 = st.columns(5)

    total_ativos = len(df_filtered)
    total_filas = len(df_filtered[df_filtered['tipo'] == 'Fila de Impressão'])
    total_mfds = len(df_filtered[df_filtered['tipo'] != 'Fila de Impressão'])
    
    status_lower = df_filtered['status'].str.lower()
    total_ok = len(df_filtered[status_lower.isin(['ok', 'online', 'ativo', 'ready', 'pronto'])])
    total_erros = total_ativos - total_ok
    total_paginas = df_filtered['total_paginas'].sum()

    with col1:
        st.markdown(f"""
            <div class="metric-card" style="border-left-color: #3b82f6;">
                <div class="metric-title">TOTAL DE ATIVOS</div>
                <div class="metric-value">{total_ativos}</div>
            </div>
        """, unsafe_allow_html=True)

    with col2:
        st.markdown(f"""
            <div class="metric-card" style="border-left-color: #8b5cf6;">
                <div class="metric-title">FILAS DE IMPRESSÃO</div>
                <div class="metric-value">{total_filas}</div>
            </div>
        """, unsafe_allow_html=True)

    with col3:
        st.markdown(f"""
            <div class="metric-card" style="border-left-color: #ec4899;">
                <div class="metric-title">DISPOSITIVOS (MFDs)</div>
                <div class="metric-value">{total_mfds}</div>
            </div>
        """, unsafe_allow_html=True)

    with col4:
        st.markdown(f"""
            <div class="metric-card" style="border-left-color: #10b981;">
                <div class="metric-title">DISPOSITIVOS OK</div>
                <div class="metric-value" style="color: #10b981;">{total_ok}</div>
            </div>
        """, unsafe_allow_html=True)

    with col5:
        st.markdown(f"""
            <div class="metric-card" style="border-left-color: #ef4444;">
                <div class="metric-title">ALERTAS / COM ERRO</div>
                <div class="metric-value" style="color: #ef4444;">{total_erros}</div>
            </div>
        """, unsafe_allow_html=True)

    st.markdown("<br>", unsafe_allow_html=True)

    # ABAS INTERNAS DE VISUALIZAÇÃO COM QUERY PARAMETERS (?subtab=slug)
    IMPRESSORAS_SUBTAB_MAP = {
        "tabela": "📋 Tabela Completa",
        "graficos": "📊 Gráficos & Estatísticas"
    }
    IMPRESSORAS_SUBTAB_REVERSE = {v: k for k, v in IMPRESSORAS_SUBTAB_MAP.items()}

    url_subtab = st.query_params.get("subtab", "tabela")
    default_title = IMPRESSORAS_SUBTAB_MAP.get(url_subtab, "📋 Tabela Completa")
    options = list(IMPRESSORAS_SUBTAB_MAP.values())
    default_idx = options.index(default_title) if default_title in options else 0

    selected_subtab = st.radio(
        "Visualização:",
        options=options,
        index=default_idx,
        horizontal=True,
        label_visibility="collapsed",
        key="impressoras_subtab_radio"
    )

    new_slug = IMPRESSORAS_SUBTAB_REVERSE.get(selected_subtab, "tabela")
    if st.query_params.get("subtab") != new_slug:
        st.query_params["subtab"] = new_slug

    st.markdown("<br>", unsafe_allow_html=True)

    if selected_subtab == "📋 Tabela Completa":
        header_col, export_col = st.columns([3, 1])
        with header_col:
            st.subheader(f"Listagem de Impressoras ({len(df_filtered)} registros)")
        with export_col:
            buffer = io.BytesIO()
            with pd.ExcelWriter(buffer, engine='openpyxl') as writer:
                df_filtered.to_excel(writer, index=False, sheet_name='Impressoras PaperCut')
            buffer.seek(0)
            
            st.download_button(
                label="📥 Exportar Excel",
                data=buffer,
                file_name=f"impressoras_papercut_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                width='stretch'
            )

        # Formatação das colunas para exibição amigável
        display_df = df_filtered.copy()
        display_df.rename(columns={
            'nome': 'Nome / Ativo',
            'servidor': 'Servidor',
            'tipo': 'Tipo de Ativo',
            'modelo': 'Modelo / Fabricante',
            'localizacao': 'Localização',
            'ip_host': 'IP / Hostname',
            'status': 'Status',
            'total_paginas': 'Total Páginas Impressas',
            'data_atualizacao': 'Última Atualização'
        }, inplace=True)

        cols_to_show = [
            'Nome / Ativo', 'Tipo de Ativo', 'Status', 'Modelo / Fabricante',
            'Localização', 'IP / Hostname', 'Servidor', 'Total Páginas Impressas', 'Última Atualização'
        ]
        
        cols_to_show = [c for c in cols_to_show if c in display_df.columns]

        # Formata o campo IP / Hostname para URL quando for um IPv4 válido (caso contrário, define None para evitar links quebrados)
        display_df['IP / Hostname'] = display_df['IP / Hostname'].apply(
            lambda x: get_printer_url(x) if get_printer_url(x) else None
        )


        df_page, current_page, total_pages, total_items = paginate_items(
            display_df[cols_to_show],
            page_key="impressoras",
            items_per_page=items_per_page
        )

        if "last_selected_printer" not in st.session_state:
            st.session_state["last_selected_printer"] = None

        selection_event = st.dataframe(
            df_page,
            width='stretch',
            hide_index=True,
            column_config={
                "IP / Hostname": st.column_config.LinkColumn(
                    "IP / Hostname",
                    display_text=r"https?://(.*)",
                    help="Clique para abrir a interface web da impressora"
                ),
                "Total Páginas Impressas": st.column_config.NumberColumn(format="%d"),
                "Última Atualização": st.column_config.DatetimeColumn(format="DD/MM/YYYY HH:mm"),
            },
            on_select="rerun",
            selection_mode="single-row",
            key="tabela_impressoras_datagrid"
        )

        selected_rows = selection_event.selection.rows if hasattr(selection_event, "selection") else []
        
        if selected_rows:
            current_selected = selected_rows[0]
            if st.session_state["last_selected_printer"] != current_selected:
                st.session_state["last_selected_printer"] = current_selected
                row_data = display_df.iloc[(current_page - 1) * items_per_page + current_selected]
                show_printer_details(row_data)
        else:
            st.session_state["last_selected_printer"] = None

        render_pagination_controls(
            page_key="impressoras",
            current_page=current_page,
            total_pages=total_pages,
            total_items=total_items,
            items_per_page=items_per_page
        )

    elif selected_subtab == "📊 Gráficos & Estatísticas":
        st.subheader("📊 Análise Gráfica de Impressoras")
        
        g_col1, g_col2 = st.columns(2)
        
        with g_col1:
            st.markdown("### 📌 Dispositivos por Tipo")
            if not df_filtered.empty and 'tipo' in df_filtered.columns:
                tipo_counts = df_filtered['tipo'].value_counts().reset_index()
                tipo_counts.columns = ['Tipo', 'Quantidade']
                st.bar_chart(tipo_counts, x='Tipo', y='Quantidade', width='stretch')
            else:
                st.info("Sem dados suficientes.")
                
        with g_col2:
            st.markdown("### 🟢 Dispositivos por Status")
            if not df_filtered.empty and 'status' in df_filtered.columns:
                status_counts = df_filtered['status'].value_counts().reset_index()
                status_counts.columns = ['Status', 'Quantidade']
                st.bar_chart(status_counts, x='Status', y='Quantidade', width='stretch')
            else:
                st.info("Sem dados suficientes.")
