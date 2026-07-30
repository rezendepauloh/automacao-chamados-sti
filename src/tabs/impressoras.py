import io
import pandas as pd
import streamlit as st
from datetime import datetime
from src.database import get_impressoras_df
from src.papercut_scraper import run_papercut_scraper


def render_impressoras_page():
    """
    Renderiza a página principal de gestão de impressoras e dispositivos do PaperCut.
    """
    st.markdown("""
        <style>
            .metric-card {
                background-color: #1e222a;
                border-radius: 8px;
                padding: 15px;
                border-left: 4px solid #3b82f6;
                margin-bottom: 10px;
            }
            .metric-title {
                font-size: 0.85rem;
                color: #9ca3af;
                margin-bottom: 4px;
            }
            .metric-value {
                font-size: 1.6rem;
                font-weight: bold;
                color: #f3f4f6;
            }
            .status-ok {
                color: #10b981;
                font-weight: bold;
            }
            .status-error {
                color: #ef4444;
                font-weight: bold;
            }
        </style>
    """, unsafe_allow_html=True)

    st.title("🖨️ Gestão de Impressoras & Dispositivos (PaperCut)")
    st.caption("Visualização unificada e controle de filas de impressão e dispositivos multifuncionais (MFDs).")

    # Carrega dados do banco SQLite
    df = get_impressoras_df()

    if df.empty:
        st.info("ℹ️ Nenhuma impressora cadastrada no banco de dados. Clique abaixo para executar o scraper ou importar os CSVs.")
        if st.button("🚀 Executar Coleta do PaperCut", type="primary"):
            with st.spinner("Conectando ao PaperCut e processando arquivos CSV..."):
                run_papercut_scraper()
                st.rerun()
        return

    # -----------------------------------------------------------------------------
    # SIDEBAR - FILTROS
    # -----------------------------------------------------------------------------
    with st.sidebar:
        st.markdown("### 🔍 Filtros de Impressoras")

        # Busca textual rápida
        search_query = st.text_input(
            "Buscar por Nome, IP, Modelo ou Local",
            placeholder="Ex: PRT-PGJ, Ricoh, 10.10...",
            key="papercut_search"
        ).strip().lower()

        # Filtro de Tipo
        tipos_disponiveis = ["Todos"] + sorted(list(df['tipo'].dropna().unique()))
        selected_tipo = st.selectbox("Tipo de Ativo", tipos_disponiveis)

        # Filtro de Status
        status_disponiveis = ["Todos"] + sorted(list(df['status'].dropna().unique()))
        selected_status = st.selectbox("Status", status_disponiveis)

        # Filtro de Localização
        locais_disponiveis = sorted(list(df['localizacao'].dropna().unique()))
        selected_locais = st.multiselect("Localização / Prédio", locais_disponiveis)

        # Filtro de Modelo
        modelos_disponiveis = sorted(list(df['modelo'].dropna().unique()))
        selected_modelos = st.multiselect("Fabricante / Modelo", modelos_disponiveis)

        st.markdown("---")
        st.markdown("### ⚙️ Ações e Sincronização")
        if st.button("🔄 Recarregar Coleta do PaperCut", use_container_width=True):
            with st.spinner("Sincronizando com o PaperCut..."):
                run_papercut_scraper()
                st.success("Coleta atualizada com sucesso!")
                st.rerun()

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
    
    # Status OK vs Erro
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
                <div class="metric-title">STATUS OPERACIONAL OK</div>
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
            # Exportação Excel/CSV
            buffer = io.BytesIO()
            with pd.ExcelWriter(buffer, engine='openpyxl') as writer:
                df_filtered.to_excel(writer, index=False, sheet_name='Impressoras PaperCut')
            buffer.seek(0)
            
            st.download_button(
                label="📥 Exportar Excel",
                data=buffer,
                file_name=f"impressoras_papercut_{datetime.now().strftime('%Y%m%m_%H%M%S')}.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                use_container_width=True
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
        
        # Filtra apenas colunas existentes
        cols_to_show = [c for c in cols_to_show if c in display_df.columns]

        st.dataframe(
            display_df[cols_to_show],
            use_container_width=True,
            hide_index=True,
            column_config={
                "Total Páginas Impressas": st.column_config.NumberColumn(format="%d"),
                "Última Atualização": st.column_config.DatetimeColumn(format="DD/MM/YYYY HH:mm"),
            }
        )

    elif selected_subtab == "📊 Gráficos & Estatísticas":

        st.subheader("📊 Análise Gráfica de Impressoras")
        
        g_col1, g_col2 = st.columns(2)

        with g_col1:
            st.markdown("#### Distribution por Status")
            status_counts = df_filtered['status'].value_counts()
            st.bar_chart(status_counts)

        with g_col2:
            st.markdown("#### Distribuição por Tipo de Ativo")
            tipo_counts = df_filtered['tipo'].value_counts()
            st.bar_chart(tipo_counts)

        st.markdown("---")
        st.markdown("#### Top 10 Impressoras por Volume de Páginas Impressas")
        df_top_pages = df_filtered.sort_values(by='total_paginas', ascending=False).head(10)
        if not df_top_pages.empty:
            st.bar_chart(data=df_top_pages, x='nome', y='total_paginas')
        else:
            st.info("Sem dados estatísticos de páginas para exibir.")
