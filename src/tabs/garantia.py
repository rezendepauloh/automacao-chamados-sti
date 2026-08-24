import io
import streamlit as st
import streamlit.components.v1 as components
import pandas as pd
from datetime import datetime
from src.database import get_garantia_contratos_df, get_garantia_chamados_df, sync_garantia_from_excel
from src.components.subtabs import render_subtabs
from src.components.calendar import render_master_calendar
from src.components.pagination import (
    render_items_per_page_selector,
    paginate_items,
    render_pagination_controls
)
from src.components.status_banner import render_log_expander
from src.syncs.sync_garantia import check_garantia_sync_running, read_garantia_last_log_lines

def parse_date_to_iso_and_br(date_val):
    if pd.isna(date_val) or not date_val:
        return None, None
    s = str(date_val).strip()
    try:
        import re
        if re.match(r'^\d{4}-\d{2}-\d{2}', s):
            dt = pd.to_datetime(s, errors='coerce')
        else:
            dt = pd.to_datetime(s, dayfirst=True, errors='coerce')
        if pd.isna(dt):
            return None, None
        return dt.strftime('%Y-%m-%d'), dt.strftime('%d/%m/%Y')
    except Exception:
        return None, None

def render_garantia_page():
    st.markdown("# 🛡️ Sistema de Controle de Garantia")
    st.caption("Acompanhe os contratos de garantia, vigências e chamados de manutenção abertos junto aos fornecedores.")
    
    garantia_ativo = check_garantia_sync_running()

    if "was_garantia_syncing" not in st.session_state:
        st.session_state["was_garantia_syncing"] = False

    if st.session_state["was_garantia_syncing"] and not garantia_ativo:
        st.session_state["was_garantia_syncing"] = False
        st.cache_data.clear()
        st.toast("🎉 Sincronização de garantias concluída com sucesso!", icon="✅")
        st.rerun()

    if garantia_ativo:
        st.session_state["was_garantia_syncing"] = True

    render_log_expander(
        "🤖 Sincronização de Garantias em Segundo Plano",
        garantia_ativo,
        read_garantia_last_log_lines,
        check_garantia_sync_running,
        "O robô está lendo a planilha de garantias do OneDrive. O painel permanece livre para uso!"
    )

    st.sidebar.markdown("## ⚙️ Ações e Sincronização")
    if garantia_ativo:
        st.sidebar.button("🤖 Sincronizando...", width='stretch', disabled=True)
    else:
        if st.sidebar.button("🔄 Sincronizar com Excel", type="primary", width='stretch', help="Busca atualizações na planilha de garantia em segundo plano."):
            import sys, subprocess, time
            popen_kwargs = {"creationflags": subprocess.CREATE_NO_WINDOW} if sys.platform == "win32" else {}
            subprocess.Popen([sys.executable, "src/syncs/sync_garantia.py"], **popen_kwargs)
            time.sleep(0.5)
            st.toast("🚀 Sincronização iniciada em segundo plano!", icon="🤖")
            st.rerun()

    st.sidebar.markdown("---")

    df_contratos = get_garantia_contratos_df()
    df_chamados = get_garantia_chamados_df()


    if df_contratos.empty and df_chamados.empty:
        st.warning("Nenhum dado cadastrado no banco SQLite. Clique no botão acima para sincronizar com a planilha.")
        return

    # --- SUBTABS PRINCIPAIS ---
    GARANTIA_SUBTAB_MAP = {
        "contratos": "📜 Contratos de Garantia",
        "chamados": "🛠️ Chamados de Garantia",
        "calendario": "📅 Calendário de Garantias",
        "graficos": "📊 Gráficos & Estatísticas"
    }

    selected_subtab = render_subtabs(GARANTIA_SUBTAB_MAP, default_slug="contratos", key="garantia_subtab_radio")

    st.markdown("<br>", unsafe_allow_html=True)


    # -------------------------------------------------------------------------
    # ABA 1: CONTRATOS DE GARANTIA
    # -------------------------------------------------------------------------
    if selected_subtab == "📜 Contratos de Garantia":
        st.sidebar.markdown("## 🔍 Filtros de Contratos")

        fornecedores = ["Todos"] + sorted([f for f in df_contratos['fornecedor'].dropna().unique() if str(f).strip()])
        selected_forn = st.sidebar.selectbox("Fornecedor / Empresa", fornecedores, key="f_garantia_forn")

        status_opt = ["Todos", "Garantia Ativa", "A Vencer (≤ 30 dias)", "Garantia Vencida", "Não Informada"]
        selected_status_contrato = st.sidebar.selectbox("Status da Vigência", status_opt, key="f_garantia_status_contrato")

        search_contrato = st.sidebar.text_input("🔎 Buscar (Contrato, SAJ, Item, NF)", "", key="f_garantia_search_contrato").strip().lower()

        items_per_page_c = render_items_per_page_selector(
            key_prefix="garantia_contratos",
            options=[10, 25, 50, 100, "Todos"],
            default_index=1,
            label="📄 Contratos por página:"
        )

        df_filtered_c = df_contratos.copy()
        if selected_forn != "Todos":
            df_filtered_c = df_filtered_c[df_filtered_c['fornecedor'] == selected_forn]
        if selected_status_contrato != "Todos":
            df_filtered_c = df_filtered_c[df_filtered_c['status_garantia'] == selected_status_contrato]

        if search_contrato:
            mask = (
                df_filtered_c['contrato'].str.lower().str.contains(search_contrato, na=False) |
                df_filtered_c['pu_saj'].str.lower().str.contains(search_contrato, na=False) |
                df_filtered_c['item'].str.lower().str.contains(search_contrato, na=False) |
                df_filtered_c['fornecedor'].str.lower().str.contains(search_contrato, na=False) |
                df_filtered_c['nota_fiscal'].str.lower().str.contains(search_contrato, na=False)
            )
            df_filtered_c = df_filtered_c[mask]

        # CARDS KPI
        k1, k2, k3, k4 = st.columns(4)
        with k1:
            st.markdown(f"""
                <div class="metric-card" style="border-left-color: #3b82f6;">
                    <div class="metric-title">TOTAL DE CONTRATOS</div>
                    <div class="metric-value">{len(df_filtered_c)}</div>
                </div>
            """, unsafe_allow_html=True)
        with k2:
            st.markdown(f"""
                <div class="metric-card" style="border-left-color: #10b981;">
                    <div class="metric-title">GARANTIAS ATIVAS</div>
                    <div class="metric-value" style="color: #10b981;">{len(df_filtered_c[df_filtered_c['status_garantia'] == 'Garantia Ativa'])}</div>
                </div>
            """, unsafe_allow_html=True)
        with k3:
            st.markdown(f"""
                <div class="metric-card" style="border-left-color: #f59e0b;">
                    <div class="metric-title">VENCENDO EM 30 DIAS</div>
                    <div class="metric-value" style="color: #f59e0b;">{len(df_filtered_c[df_filtered_c['status_garantia'] == 'A Vencer (≤ 30 dias)'])}</div>
                </div>
            """, unsafe_allow_html=True)
        with k4:
            st.markdown(f"""
                <div class="metric-card" style="border-left-color: #ef4444;">
                    <div class="metric-title">GARANTIAS VENCIDAS</div>
                    <div class="metric-value" style="color: #ef4444;">{len(df_filtered_c[df_filtered_c['status_garantia'] == 'Garantia Vencida'])}</div>
                </div>
            """, unsafe_allow_html=True)

        st.markdown("<br>", unsafe_allow_html=True)

        h_col1, h_col2 = st.columns([3, 1])
        with h_col1:
            st.subheader(f"📋 Relação de Contratos de Garantia ({len(df_filtered_c)} registros)")
        with h_col2:
            buffer = io.BytesIO()
            with pd.ExcelWriter(buffer, engine='openpyxl') as writer:
                df_filtered_c.to_excel(writer, index=False, sheet_name='Contratos')
            buffer.seek(0)
            st.download_button(
                label="📥 Exportar Excel",
                data=buffer,
                file_name=f"contratos_garantia_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                width='stretch'
            )

        cols_c = [
            'contrato', 'pu_saj', 'item', 'contratacao_por', 'fornecedor',
            'termo_referencia', 'termo_recebimento', 'nota_fiscal',
            'garantia_inicio', 'garantia_fim', 'status_garantia', 'link_suporte'
        ]
        cols_c = [c for c in cols_c if c in df_filtered_c.columns]

        df_page_c, current_page_c, total_pages_c, total_items_c = paginate_items(
            df_filtered_c[cols_c],
            page_key="garantia_contratos",
            items_per_page=items_per_page_c
        )

        st.dataframe(
            df_page_c,
            column_config={
                "contrato": st.column_config.TextColumn("Contrato"),
                "pu_saj": st.column_config.TextColumn("PU SAJ"),
                "item": st.column_config.TextColumn("Item / Equipamento"),
                "contratacao_por": st.column_config.TextColumn("Contratação por"),
                "fornecedor": st.column_config.TextColumn("Fornecedor"),
                "termo_referencia": st.column_config.TextColumn("Termo de Ref."),
                "termo_recebimento": st.column_config.TextColumn("Termo Rec. Definitivo"),
                "nota_fiscal": st.column_config.TextColumn("Nota Fiscal"),
                "garantia_inicio": st.column_config.DateColumn("Início Garantia", format="DD/MM/YYYY"),
                "garantia_fim": st.column_config.DateColumn("Fim da Garantia", format="DD/MM/YYYY"),
                "status_garantia": st.column_config.TextColumn("Status Vigência"),
                "link_suporte": st.column_config.LinkColumn("Link para Abertura de Chamado", display_text="🔗 Abrir Portal do Fornecedor"),
            },
            hide_index=True,
            width='stretch'
        )

        render_pagination_controls(
            page_key="garantia_contratos",
            current_page=current_page_c,
            total_pages=total_pages_c,
            total_items=total_items_c,
            items_per_page=items_per_page_c
        )

    # -------------------------------------------------------------------------
    # ABA 2: CHAMADOS DE GARANTIA
    # -------------------------------------------------------------------------
    elif selected_subtab == "🛠️ Chamados de Garantia":
        st.sidebar.markdown("## 🔍 Filtros de Chamados")

        status_chamados = ["Todos"] + sorted([s for s in df_chamados['status'].dropna().unique() if str(s).strip()])
        selected_status_ch = st.sidebar.selectbox("Status do Chamado", status_chamados, key="f_garantia_status_ch")

        items_disponiveis = ["Todos"] + sorted([i for i in df_chamados['item'].dropna().unique() if str(i).strip()])
        selected_item_ch = st.sidebar.selectbox("Tipo de Item", items_disponiveis, key="f_garantia_item_ch")

        search_chamado = st.sidebar.text_input("🔎 Buscar (Patrimônio, Serial, Chamado, Defeito)", "", key="f_garantia_search_ch").strip().lower()

        items_per_page_ch = render_items_per_page_selector(
            key_prefix="garantia_chamados",
            options=[10, 25, 50, 100, "Todos"],
            default_index=1,
            label="📄 Chamados por página:"
        )

        df_filtered_ch = df_chamados.copy()
        if selected_status_ch != "Todos":
            df_filtered_ch = df_filtered_ch[df_filtered_ch['status'] == selected_status_ch]
        if selected_item_ch != "Todos":
            df_filtered_ch = df_filtered_ch[df_filtered_ch['item'] == selected_item_ch]

        if search_chamado:
            mask = (
                df_filtered_ch['item'].str.lower().str.contains(search_chamado, na=False) |
                df_filtered_ch['numero_serie'].str.lower().str.contains(search_chamado, na=False) |
                df_filtered_ch['patrimonio'].str.lower().str.contains(search_chamado, na=False) |
                df_filtered_ch['chamado_mpm'].str.lower().str.contains(search_chamado, na=False) |
                df_filtered_ch['chamado_externo'].str.lower().str.contains(search_chamado, na=False) |
                df_filtered_ch['defeito'].str.lower().str.contains(search_chamado, na=False) |
                df_filtered_ch['solucao'].str.lower().str.contains(search_chamado, na=False)
            )
            df_filtered_ch = df_filtered_ch[mask]

        # CARDS KPI
        k1, k2, k3, k4 = st.columns(4)
        with k1:
            st.markdown(f"""
                <div class="metric-card" style="border-left-color: #3b82f6;">
                    <div class="metric-title">TOTAL DE CHAMADOS</div>
                    <div class="metric-value">{len(df_filtered_ch)}</div>
                </div>
            """, unsafe_allow_html=True)
        with k2:
            st.markdown(f"""
                <div class="metric-card" style="border-left-color: #10b981;">
                    <div class="metric-title">CONCLUÍDOS</div>
                    <div class="metric-value" style="color: #10b981;">{len(df_filtered_ch[df_filtered_ch['status'].str.lower().str.contains('conclu', na=False)])}</div>
                </div>
            """, unsafe_allow_html=True)
        with k3:
            st.markdown(f"""
                <div class="metric-card" style="border-left-color: #f59e0b;">
                    <div class="metric-title">EM ATENDIMENTO</div>
                    <div class="metric-value" style="color: #f59e0b;">{len(df_filtered_ch[df_filtered_ch['status'].str.lower().str.contains('atend', na=False)])}</div>
                </div>
            """, unsafe_allow_html=True)
        with k4:
            st.markdown(f"""
                <div class="metric-card" style="border-left-color: #a855f7;">
                    <div class="metric-title">ABRIR DMP / PAUSADOS</div>
                    <div class="metric-value" style="color: #a855f7;">{len(df_filtered_ch[df_filtered_ch['status'].str.lower().str.contains('dmp|paus', na=False)])}</div>
                </div>
            """, unsafe_allow_html=True)

        st.markdown("<br>", unsafe_allow_html=True)

        h_col1, h_col2 = st.columns([3, 1])
        with h_col1:
            st.subheader(f"🛠️ Chamados de Garantia ({len(df_filtered_ch)} registros)")
        with h_col2:
            buffer = io.BytesIO()
            with pd.ExcelWriter(buffer, engine='openpyxl') as writer:
                df_filtered_ch.to_excel(writer, index=False, sheet_name='Chamados')
            buffer.seek(0)
            st.download_button(
                label="📥 Exportar Excel",
                data=buffer,
                file_name=f"chamados_garantia_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                width='stretch'
            )

        cols_ch = [
            'item', 'status', 'patrimonio', 'numero_serie', 'chamado_mpm',
            'chamado_externo', 'defeito', 'solucao', 'nota_no_chamado', 'chamado_dmp'
        ]
        cols_ch = [c for c in cols_ch if c in df_filtered_ch.columns]

        df_page_ch, current_page_ch, total_pages_ch, total_items_ch = paginate_items(
            df_filtered_ch[cols_ch],
            page_key="garantia_chamados",
            items_per_page=items_per_page_ch
        )

        st.dataframe(
            df_page_ch,
            column_config={
                "item": st.column_config.TextColumn("Item / Equipamento"),
                "status": st.column_config.TextColumn("Status"),
                "patrimonio": st.column_config.TextColumn("Patrimônio"),
                "numero_serie": st.column_config.TextColumn("Número de Série"),
                "chamado_mpm": st.column_config.TextColumn("Chamado MPMS"),
                "chamado_externo": st.column_config.TextColumn("Chamado Externo / Fornecedor"),
                "defeito": st.column_config.TextColumn("Defeito Relatado"),
                "solucao": st.column_config.TextColumn("Solução Aplicada"),
                "nota_no_chamado": st.column_config.TextColumn("Nota no Chamado"),
                "chamado_dmp": st.column_config.TextColumn("Chamado DMP"),
            },
            hide_index=True,
            width='stretch'
        )

        render_pagination_controls(
            page_key="garantia_chamados",
            current_page=current_page_ch,
            total_pages=total_pages_ch,
            total_items=total_items_ch,
            items_per_page=items_per_page_ch
        )

    # -------------------------------------------------------------------------
    # ABA 3: CALENDÁRIO DE GARANTIAS
    # -------------------------------------------------------------------------
    elif selected_subtab == "📅 Calendário de Garantias":
        st.sidebar.markdown("## 🔍 Filtros do Calendário")

        fornecedores_cal = ["Todos"] + sorted([f for f in df_contratos['fornecedor'].dropna().unique() if str(f).strip()])
        selected_forn_cal = st.sidebar.selectbox("Fornecedor / Empresa", fornecedores_cal, key="f_garantia_cal_forn")

        tipo_evento_cal = st.sidebar.selectbox(
            "Tipo de Evento",
            ["Todos os Eventos", "🟢 Início de Garantia", "🔴 Fim / Vencimento de Garantia"],
            key="f_garantia_cal_tipo"
        )

        status_cal = st.sidebar.selectbox(
            "Status da Vigência",
            ["Todos", "Garantia Ativa", "A Vencer (≤ 30 dias)", "Garantia Vencida"],
            key="f_garantia_cal_status"
        )

        df_cal = df_contratos.copy()
        if selected_forn_cal != "Todos":
            df_cal = df_cal[df_cal['fornecedor'] == selected_forn_cal]
        if status_cal != "Todos":
            df_cal = df_cal[df_cal['status_garantia'] == status_cal]

        events = []
        for idx, row in df_cal.iterrows():
            contrato = str(row.get('contrato', '')).strip()
            pu_saj = str(row.get('pu_saj', '')).strip()
            item = str(row.get('item', '')).strip()
            fornecedor = str(row.get('fornecedor', '')).strip()
            nota_fiscal = str(row.get('nota_fiscal', '')).strip()
            status_garantia = str(row.get('status_garantia', '')).strip()
            link_suporte = str(row.get('link_suporte', '')).strip()

            # Início da Garantia
            if tipo_evento_cal in ["Todos os Eventos", "🟢 Início de Garantia"]:
                iso_ini, br_ini = parse_date_to_iso_and_br(row.get('garantia_inicio'))
                if iso_ini:
                    events.append({
                        "id": f"garantia_ini_{idx}",
                        "title": f"🟢 Início: {item} ({fornecedor})",
                        "start": iso_ini,
                        "backgroundColor": "#10b981",
                        "borderColor": "#059669",
                        "textColor": "#ffffff",
                        "extendedProps": {
                            "categoria_evento": "garantia",
                            "tipo": "🟢 Início da Garantia",
                            "contrato": contrato,
                            "pu_saj": pu_saj,
                            "item": item,
                            "fornecedor": fornecedor,
                            "nota_fiscal": nota_fiscal,
                            "status_garantia": status_garantia,
                            "data_formatada": br_ini,
                            "link_suporte": link_suporte
                        }
                    })

            # Fim da Garantia
            if tipo_evento_cal in ["Todos os Eventos", "🔴 Fim / Vencimento de Garantia"]:
                iso_fim, br_fim = parse_date_to_iso_and_br(row.get('garantia_fim'))
                if iso_fim:
                    bg_col = "#ef4444" if status_garantia == "Garantia Vencida" else ("#f59e0b" if "30" in status_garantia else "#3b82f6")
                    events.append({
                        "id": f"garantia_fim_{idx}",
                        "title": f"🔴 Fim: {item} ({fornecedor})",
                        "start": iso_fim,
                        "backgroundColor": bg_col,
                        "borderColor": bg_col,
                        "textColor": "#ffffff",
                        "extendedProps": {
                            "categoria_evento": "garantia",
                            "tipo": "🔴 Fim / Vencimento da Garantia",
                            "contrato": contrato,
                            "pu_saj": pu_saj,
                            "item": item,
                            "fornecedor": fornecedor,
                            "nota_fiscal": nota_fiscal,
                            "status_garantia": status_garantia,
                            "data_formatada": br_fim,
                            "link_suporte": link_suporte
                        }
                    })

        st.subheader(f"📅 Agenda de Vigências de Garantia ({len(events)} eventos mapeados)")
        render_master_calendar(events, height_px=750, scrolling_enabled=True)

    # -------------------------------------------------------------------------
    # ABA 4: GRÁFICOS & ESTATÍSTICAS
    # -------------------------------------------------------------------------
    elif selected_subtab == "📊 Gráficos & Estatísticas":

        st.subheader("📊 Análise Gráfica do Controle de Garantia")

        g1, g2 = st.columns(2)

        with g1:
            st.markdown("### 🛠️ Status dos Chamados de Garantia")
            if not df_chamados.empty and 'status' in df_chamados.columns:
                st_counts = df_chamados['status'].value_counts().reset_index()
                st_counts.columns = ['Status do Chamado', 'Quantidade']
                st.bar_chart(data=st_counts, x='Status do Chamado', y='Quantidade', width='stretch')
            else:
                st.info("Sem dados de chamados para exibir gráfico.")

        with g2:
            st.markdown("### 🏢 Contratos por Fornecedor")
            if not df_contratos.empty and 'fornecedor' in df_contratos.columns:
                forn_counts = df_contratos['fornecedor'].value_counts().reset_index()
                forn_counts.columns = ['Fornecedor', 'Quantidade']
                st.bar_chart(data=forn_counts, x='Fornecedor', y='Quantidade', width='stretch')
            else:
                st.info("Sem dados de contratos para exibir gráfico.")
