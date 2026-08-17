import pandas as pd
import streamlit as st
from src.components.pagination import (
    render_items_per_page_selector,
    paginate_items,
    render_pagination_controls
)
from src.components.status_banner import render_log_expander
from src.syncs.sync_donations import check_donations_sync_running, read_donations_last_log_lines

def render_donations_page():
    """Renderiza a página de Doação & Redistribuição de Máquinas."""
    from src.database import get_donations_data, sync_donations_from_excel
    
    st.title("🖥️ Sistema de Doação & Redistribuição de Máquinas")
    st.write("Acompanhe o inventário de equipamentos destinados a doação, redistribuição, garantia ou baixados.")
    
    donations_ativo = check_donations_sync_running()

    if "was_donations_syncing" not in st.session_state:
        st.session_state["was_donations_syncing"] = False

    if st.session_state["was_donations_syncing"] and not donations_ativo:
        st.session_state["was_donations_syncing"] = False
        st.toast("🎉 Sincronização de doações concluída com sucesso!", icon="✅")
        st.rerun()

    if donations_ativo:
        st.session_state["was_donations_syncing"] = True

    render_log_expander(
        "🤖 Sincronização de Doações em Segundo Plano",
        donations_ativo,
        read_donations_last_log_lines,
        check_donations_sync_running,
        "O robô está lendo a planilha do SharePoint. O painel permanece livre para uso!"
    )
    
    from src.config import DONATIONS_FILE_PATH
    EXCEL_PATH = str(DONATIONS_FILE_PATH)
                    
    df = get_donations_data()
    
    if df.empty:
        st.warning("⚠️ Nenhum dado encontrado no cache local. Por favor, clique em 'Sincronizar Planilha' para carregar os registros.")
        return
        
    df['Ano'] = pd.to_datetime(df['data_movimentacao'], errors='coerce').dt.year
    df['Ano'] = df['Ano'].fillna("Sem Data").astype(str).str.replace(".0", "", regex=False)
        
    st.sidebar.title("🖥️ Painel de Controle")

    if donations_ativo:
        st.sidebar.button("🤖 Atualizando...", type="primary", width='stretch', disabled=True)
    else:
        if st.sidebar.button("🔄 Sincronizar Planilha", type="primary", width='stretch', help="Busca atualizações na planilha do SharePoint em segundo plano."):
            import sys, subprocess
            creationflags = subprocess.CREATE_NO_WINDOW if sys.platform == "win32" else 0
            subprocess.Popen([sys.executable, "src/syncs/sync_donations.py"], creationflags=creationflags)
            st.toast("🚀 Sincronização de doações iniciada em segundo plano!", icon="🤖")
            st.rerun()

    st.sidebar.markdown("---")
    st.sidebar.subheader("🔍 Filtros de Equipamentos")
    
    mov_options = ["Todos"] + sorted(list(df['tipo_movimentacao'].unique()))
    selected_mov = st.sidebar.selectbox("Tipo de Movimentação", mov_options, key="donations_mov")
    
    equip_options = ["Todos"] + sorted(list(df['equipamento'].unique()))
    selected_equip = st.sidebar.selectbox("Tipo de Equipamento", equip_options, key="donations_equip")
    
    model_options = ["Todos"] + sorted(list(df['modelo'].unique()))
    selected_model = st.sidebar.selectbox("Modelo", model_options, key="donations_model")
    
    year_options = ["Todos"] + sorted(list(df['Ano'].unique()), reverse=True)
    selected_year = st.sidebar.selectbox("Ano da Movimentação", year_options, key="donations_year")

    ssd_options = ["Todos"] + sorted(list(df['ssd'].unique()))
    selected_ssd = st.sidebar.selectbox("SSD", ssd_options, key="donations_ssd")
    
    search_query = st.sidebar.text_input("🔎 Buscar (Patrimônio, Modelo, Chamado)", "", key="donations_search").strip()
    
    items_per_page = render_items_per_page_selector(
        key_prefix="redistribuicao",
        options=[10, 25, 50, 100, "Todos"],
        default_index=1,
        label="📄 Equipamentos por página:"
    )

    st.sidebar.markdown("---")
    st.sidebar.subheader("📋 Gerador de Texto (Preparo)")
    
    valid_dates = df[df['data_movimentacao'] != '']['data_movimentacao'].unique()
    valid_dates = sorted(list(valid_dates), reverse=True)
    
    def format_date_br(date_str):
        from datetime import datetime
        try:
            return datetime.strptime(date_str, "%Y-%m-%d").strftime("%d/%m/%Y")
        except:
            return date_str
            
    selected_date_str = st.sidebar.selectbox("Selecione a Data de Preparo", valid_dates, format_func=format_date_br)
    generate_btn = st.sidebar.button("📝 Gerar Texto do Chamado", width='stretch')

    @st.dialog("📋 Texto de Preparo de Chamado", width="large")
    def show_preparo_text(date_str, df_all):
        df_date = df_all[df_all['data_movimentacao'] == date_str]
        
        if df_date.empty:
            st.warning("Nenhum equipamento encontrado nesta data.")
            return
            
        from datetime import datetime
        try:
            dt_obj = datetime.strptime(date_str, "%Y-%m-%d")
            formatted_date = dt_obj.strftime("%d/%m/%Y")
        except:
            formatted_date = date_str
            
        movs = sorted(list(df_date['tipo_movimentacao'].unique()))
        movs_clean = [m.strip().capitalize() for m in movs if m.strip()]
        if len(movs_clean) == 1:
            movs_str = movs_clean[0]
        elif len(movs_clean) > 1:
            movs_str = ", ".join(movs_clean[:-1]) + " e " + movs_clean[-1]
        else:
            movs_str = "Movimentação"
            
        subject = f"[DOAÇÃO] - Preparação de {movs_str} de equipamentos do dia {formatted_date}"
        
        st.write("📋 **Assunto do Chamado:**")
        st.code(subject, language="text")
        st.markdown("---")
            
        html_parts = []
        html_parts.append("<div style='font-family: Arial, Helvetica, sans-serif; color: #000000; line-height: 1.5;'>")
        html_parts.append("<p>Prezados, boa tarde.</p>")
        html_parts.append(f"<p>Na tarde de hoje (<strong>{formatted_date}</strong>), preparamos os seguintes equipamentos, sendo eles:</p>")
        
        for mov_type, grp in df_date.groupby('tipo_movimentacao'):
            html_parts.append(f"<p style='margin-top: 20px; margin-bottom: 8px;'><strong>🔹 Equipamentos para {mov_type.upper()}:</strong></p>")
            
            has_ssd = grp['ssd'].astype(str).str.strip().any()
            has_obs = grp['motivo_baixa'].astype(str).str.strip().any()
            
            table_html = [
                "<table border='2' cellpadding='6' cellspacing='0' style='border-collapse: collapse; width: 100%; border: 2px solid #cccccc; font-family: Arial, Helvetica, sans-serif; font-size: 13px; color: #000000;'>"
            ]
            th_style = "background-color: #2f5597; border: 2px solid #cccccc; padding: 6px 10px; text-align: left;"
            
            headers_html = [
                f"<tr>",
                f"<th style='{th_style}'><span style=\"color:#ffffff\">Patrimônio</span></th>",
                f"<th style='{th_style}'><span style=\"color:#ffffff\">Modelo</span></th>",
                f"<th style='{th_style}'><span style=\"color:#ffffff\">Serial Number PC</span></th>",
                f"<th style='{th_style}'><span style=\"color:#ffffff\">Equipamento</span></th>"
            ]
            if has_ssd:
                headers_html.append(f"<th style='{th_style}'><span style=\"color:#ffffff\">SSD</span></th>")
            if has_obs:
                headers_html.append(f"<th style='{th_style}'><span style=\"color:#ffffff\">Motivo/Obs</span></th>")
            headers_html.append("</tr>")
            
            table_html.append("".join(headers_html))
            
            for idx, (_, row) in enumerate(grp.iterrows()):
                pat = str(row.get('patrimonio', '')).strip()
                mod = str(row.get('modelo', '')).strip()
                ser = str(row.get('serial_number', '')).strip()
                eqp = str(row.get('equipamento', '')).strip()
                ssd = str(row.get('ssd', '')).strip()
                obs = str(row.get('motivo_baixa', '')).strip()
                
                bg_style = "background-color: #d9e1f2;" if idx % 2 == 1 else ""
                
                row_html = [
                    f"<tr>",
                    f"<td style='border: 2px solid #cccccc; padding: 6px 10px; {bg_style}'>{pat}</td>",
                    f"<td style='border: 2px solid #cccccc; padding: 6px 10px; {bg_style}'>{mod}</td>",
                    f"<td style='border: 2px solid #cccccc; padding: 6px 10px; {bg_style}'>{ser}</td>",
                    f"<td style='border: 2px solid #cccccc; padding: 6px 10px; {bg_style}'>{eqp}</td>"
                ]
                if has_ssd:
                    row_html.append(f"<td style='border: 2px solid #cccccc; padding: 6px 10px; {bg_style}'>{ssd}</td>")
                if has_obs:
                    row_html.append(f"<td style='border: 2px solid #cccccc; padding: 6px 10px; {bg_style}'>{obs}</td>")
                row_html.append("</tr>")
                
                table_html.append("".join(row_html))
 
            table_html.append("</table>")
            html_parts.append("".join(table_html))
            
        html_parts.append("</div>")
        full_html = "\n".join(html_parts)
        
        st.write("💡 Selecione o texto abaixo com o mouse, copie (Ctrl+C) e cole diretamente no chamado do OTRS:")
        st.markdown(
            f'<div style="background-color: #ffffff; padding: 20px; border-radius: 6px; border: 1px solid #dddddd; max-height: 400px; overflow-y: auto;">{full_html}</div>', 
            unsafe_allow_html=True
        )
        
        st.markdown("---")
        st.write("💻 Ou copie o código-fonte HTML abaixo (clique no botão **'Código-Fonte'** no OTRS e cole):")
        st.code(full_html, language="html")

    if generate_btn:
        show_preparo_text(selected_date_str, df)

    df_filtered = df.copy()
    if selected_mov != "Todos":
        df_filtered = df_filtered[df_filtered['tipo_movimentacao'] == selected_mov]
    if selected_equip != "Todos":
        df_filtered = df_filtered[df_filtered['equipamento'] == selected_equip]
    if selected_model != "Todos":
        df_filtered = df_filtered[df_filtered['modelo'] == selected_model]
    if selected_year != "Todos":
        df_filtered = df_filtered[df_filtered['Ano'] == selected_year]
    if selected_ssd != "Todos":
        df_filtered = df_filtered[df_filtered['ssd'] == selected_ssd]

    if search_query:
        query_lower = search_query.lower()
        df_filtered = df_filtered[
            df_filtered['patrimonio'].str.lower().str.contains(query_lower) |
            df_filtered['modelo'].str.lower().str.contains(query_lower) |
            df_filtered['chamado'].str.lower().str.contains(query_lower)
        ]

    st.markdown("---")
    kpi_col1, kpi_col2, kpi_col3, kpi_col4, kpi_col5 = st.columns(5)
    
    total_equip = len(df_filtered)
    doados = len(df_filtered[df_filtered['tipo_movimentacao'].str.lower() == 'doação'])
    redistribuicoes = len(df_filtered[df_filtered['tipo_movimentacao'].str.lower() == 'redistribuição'])
    baixas = len(df_filtered[df_filtered['tipo_movimentacao'].str.lower() == 'baixa'])
    garantias = len(df_filtered[df_filtered['tipo_movimentacao'].str.lower() == 'garantia'])
    
    with kpi_col1:
        st.markdown(f"""
            <div class="metric-card" style="border-left-color: #3b82f6;">
                <div class="metric-title">EQUIPAMENTOS</div>
                <div class="metric-value">{total_equip}</div>
            </div>
        """, unsafe_allow_html=True)

    with kpi_col2:
        st.markdown(f"""
            <div class="metric-card" style="border-left-color: #10b981;">
                <div class="metric-title">DOAÇÕES</div>
                <div class="metric-value" style="color: #10b981;">{doados}</div>
            </div>
        """, unsafe_allow_html=True)

    with kpi_col3:
        st.markdown(f"""
            <div class="metric-card" style="border-left-color: #8b5cf6;">
                <div class="metric-title">REDISTRIBUIÇÕES</div>
                <div class="metric-value" style="color: #8b5cf6;">{redistribuicoes}</div>
            </div>
        """, unsafe_allow_html=True)

    with kpi_col4:
        st.markdown(f"""
            <div class="metric-card" style="border-left-color: #ef4444;">
                <div class="metric-title">BAIXAS</div>
                <div class="metric-value" style="color: #ef4444;">{baixas}</div>
            </div>
        """, unsafe_allow_html=True)

    with kpi_col5:
        st.markdown(f"""
            <div class="metric-card" style="border-left-color: #f59e0b;">
                <div class="metric-title">GARANTIAS</div>
                <div class="metric-value" style="color: #f59e0b;">{garantias}</div>
            </div>
        """, unsafe_allow_html=True)


    st.markdown("---")
    
    g_col1, g_col2 = st.columns(2)
    
    with g_col1:
        st.subheader("📊 Distribuição por Movimentação")
        if not df_filtered.empty:
            mov_counts = df_filtered['tipo_movimentacao'].value_counts().reset_index()
            mov_counts.columns = ['Movimentação', 'Quantidade']
            st.bar_chart(data=mov_counts, x='Movimentação', y='Quantidade', width='stretch')
        else:
            st.info("Sem dados para exibir o gráfico.")
            
    with g_col2:
        st.subheader("📅 Histórico de Movimentações por Ano")
        if not df_filtered.empty:
            df_filtered['Ano'] = pd.to_datetime(df_filtered['data_movimentacao'], errors='coerce').dt.year
            df_filtered['Ano'] = df_filtered['Ano'].fillna("Sem Data").astype(str).str.replace(".0", "", regex=False)
            
            ano_counts = df_filtered.groupby(['Ano', 'tipo_movimentacao']).size().unstack(fill_value=0)
            st.bar_chart(ano_counts, width='stretch')
        else:
            st.info("Sem dados para exibir o gráfico.")
            
    st.markdown("---")
    st.subheader("📋 Detalhamento dos Equipamentos")
    
    df_page, current_page, total_pages, total_items = paginate_items(
        df_filtered,
        page_key="redistribuicao",
        items_per_page=items_per_page
    )

    st.dataframe(
        df_page,
        column_config={
            "patrimonio": st.column_config.TextColumn("Patrimônio"),
            "modelo": st.column_config.TextColumn("Modelo"),
            "serial_number": st.column_config.TextColumn("Número de Série"),
            "equipamento": st.column_config.TextColumn("Equipamento"),
            "tipo_movimentacao": st.column_config.TextColumn("Movimentação"),
            "data_movimentacao": st.column_config.DateColumn("Data da Movimentação", format="DD/MM/YYYY"),
            "chamado": st.column_config.TextColumn("Chamado relacionado"),
            "ssd": st.column_config.TextColumn("SSD"),
            "motivo_baixa": st.column_config.TextColumn("Motivo da Baixa"),
        },
        hide_index=True,
        width='stretch'
    )

    render_pagination_controls(
        page_key="redistribuicao",
        current_page=current_page,
        total_pages=total_pages,
        total_items=total_items,
        items_per_page=items_per_page
    )

