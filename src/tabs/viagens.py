import io
import os
import sys
import subprocess
import time
from datetime import datetime, timedelta
import pandas as pd
import streamlit as st

from src.components.subtabs import render_subtabs
from src.components.calendar import render_master_calendar
from src.components.pagination import (
    render_items_per_page_selector,
    paginate_items,
    render_pagination_controls
)
from src.components.status_banner import render_log_expander
from src.database import get_viagens_df, sync_viagens_from_excel
from src.syncs.sync_viagens import check_viagens_sync_running, read_viagens_last_log_lines
from src.config import _cfg

@st.dialog("⚙️ Configurar / Enviar Planilha de Viagens")
def modal_config_viagens():
    """Modal (@st.dialog) para gerenciar o link do SharePoint ou envio manual da planilha de viagens."""
    st.markdown("### ✈️ Gestão da Planilha de Viagens da Bancada")
    st.caption("Consulte a URL da planilha no SharePoint, force a sincronização ou envie o arquivo .xlsx diretamente.")

    tab_online, tab_upload = st.tabs(["🌐 Link SharePoint / Atualizar Online", "📥 Envio Direto de Planilha"])

    with tab_online:
        excel_url = (_cfg("VIAGENS_EXCEL_RELATIVE_PATH") or os.getenv("VIAGENS_EXCEL_RELATIVE_PATH", "")).strip()
        st.write("Planilha oficial vinculada no SharePoint:")

        if excel_url.startswith("http://") or excel_url.startswith("https://"):
            st.link_button(
                "🌐 Abrir Planilha no SharePoint (Excel Online) ↗",
                excel_url,
                type="secondary",
                width='stretch',
                help="Abre o arquivo original diretamente no SharePoint / Excel Online em uma nova aba."
            )
            st.markdown("<div style='height: 10px;'></div>", unsafe_allow_html=True)

        if st.button("🚀 Sincronizar pelo Link do SharePoint Agora", type="primary", width='stretch'):
            popen_kwargs = {"creationflags": subprocess.CREATE_NO_WINDOW} if sys.platform == "win32" else {}
            subprocess.Popen([sys.executable, "src/syncs/sync_viagens.py"], **popen_kwargs)
            st.toast("🚀 Sincronização disparada com sucesso!", icon="🤖")
            st.rerun()

    with tab_upload:
        st.write("Faça o upload manual do arquivo Excel (.xlsx) da planilha de Viagens:")
        uploaded_excel = st.file_uploader("Selecione o arquivo Excel (.xlsx)", type=["xlsx", "xls"], key="modal_up_viagens_excel")

        if st.button("⚡ Processar e Gravar no Banco", type="primary", width='stretch'):
            if not uploaded_excel:
                st.warning("Selecione um arquivo Excel primeiro.")
            else:
                with st.spinner("Processando planilha de viagens..."):
                    res = sync_viagens_from_excel(uploaded_excel)
                    if res:
                        st.success("🎉 Planilha de viagens importada com sucesso!")
                        time.sleep(1.2)
                        st.rerun()
                    else:
                        st.error("Não foi possível processar a planilha. Verifique o formato das colunas.")

def render_viagens_page():
    """Página principal de Viagens da Bancada com FullCalendar e Tabela detalhada."""
    st.markdown("# ✈️ Viagens da Bancada STI")
    st.caption("Acompanhe o cronograma de deslocamentos técnicos, comarcas atendidas, chamados vinculados e técnicos escalados.")

    viagens_ativo = check_viagens_sync_running()

    if "was_viagens_syncing" not in st.session_state:
        st.session_state["was_viagens_syncing"] = False

    if st.session_state["was_viagens_syncing"] and not viagens_ativo:
        st.session_state["was_viagens_syncing"] = False
        st.cache_data.clear()
        st.toast("🎉 Sincronização de viagens concluída com sucesso!", icon="✅")
        st.rerun()

    if viagens_ativo:
        st.session_state["was_viagens_syncing"] = True

    render_log_expander(
        "🤖 Sincronização de Viagens em Segundo Plano",
        viagens_ativo,
        read_viagens_last_log_lines,
        check_viagens_sync_running,
        "O robô está atualizando a planilha de viagens do SharePoint. O painel permanece livre para uso!"
    )

    st.sidebar.markdown("## ⚙️ Ações e Sincronização")
    if viagens_ativo:
        st.sidebar.button("🤖 Sincronizando...", width='stretch', disabled=True)
    else:
        if st.sidebar.button("🔄 Sincronizar Planilha", type="primary", width='stretch', help="Busca atualizações na planilha de viagens em segundo plano."):
            popen_kwargs = {"creationflags": subprocess.CREATE_NO_WINDOW} if sys.platform == "win32" else {}
            subprocess.Popen([sys.executable, "src/syncs/sync_viagens.py"], **popen_kwargs)
            time.sleep(0.5)
            st.toast("🚀 Sincronização iniciada em segundo plano!", icon="🤖")
            st.rerun()

    if st.sidebar.button("⚙️ Configurar / Enviar Planilha", width='stretch', help="Gerenciar link do SharePoint ou enviar a planilha de viagens manualmente."):
        modal_config_viagens()

    st.sidebar.markdown("---")

    df_viagens = get_viagens_df()

    if df_viagens.empty:
        st.warning("Nenhum dado de viagem encontrado no banco. Clique no botão acima para sincronizar com a planilha.")
        return

    # --- SUBTABS ---
    VIAGENS_SUBTAB_MAP = {
        "calendario": "📅 Calendário de Viagens",
        "tabela": "📋 Tabela de Viagens"
    }

    selected_subtab = render_subtabs(VIAGENS_SUBTAB_MAP, default_slug="calendario", key="viagens_subtab_radio")
    st.markdown("<br>", unsafe_allow_html=True)

    # -------------------------------------------------------------------------
    # ABA 1: CALENDÁRIO DE VIAGENS
    # -------------------------------------------------------------------------
    if selected_subtab == "📅 Calendário de Viagens":
        st.sidebar.markdown("## 🔍 Filtros do Calendário")

        # Filtro de Técnicos
        tecnicos_unicos = set()
        for quem in df_viagens["quem_foi"].dropna().unique():
            partes = [p.strip() for p in quem.replace("/", ",").split(",") if p.strip()]
            tecnicos_unicos.update(partes)

        filtro_tecnico = st.sidebar.selectbox("👤 Técnico / Membro", ["Todos"] + sorted(list(tecnicos_unicos)), key="f_viagem_tecnico_cal")

        df_cal = df_viagens.copy()
        if filtro_tecnico != "Todos":
            df_cal = df_cal[df_cal["quem_foi"].str.contains(filtro_tecnico, case=False, na=False)]

        events = []
        for idx, row in df_cal.iterrows():
            saida_iso = row.get("saida_iso", "")
            retorno_iso = row.get("retorno_iso", "")
            localidade = row.get("localidade", "")
            quem_foi = row.get("quem_foi", "")
            chamado = row.get("chamado", "")
            saida_br = row.get("saida_br", "")
            retorno_br = row.get("retorno_br", "")

            if not saida_iso:
                continue

            # FullCalendar: all-day end date is exclusive, so add 1 day to cover the full return day
            cal_end = saida_iso
            if retorno_iso:
                try:
                    dt_ret = datetime.strptime(retorno_iso, "%Y-%m-%d")
                    cal_end = (dt_ret + timedelta(days=1)).strftime("%Y-%m-%d")
                except Exception:
                    cal_end = retorno_iso

            title_text = f"✈️ {localidade}"
            if quem_foi:
                title_text += f" ({quem_foi})"

            events.append({
                "id": f"viagem_{row.get('id', idx)}",
                "title": title_text,
                "start": saida_iso,
                "end": cal_end,
                "backgroundColor": "#0891b2",
                "borderColor": "#0e7490",
                "textColor": "#ffffff",
                "extendedProps": {
                    "categoria_evento": "viagem",
                    "tipo": "✈️ Viagem da Bancada",
                    "localidade": localidade,
                    "quem_foi": quem_foi,
                    "chamado": chamado,
                    "saida_br": saida_br,
                    "retorno_br": retorno_br,
                    "raw_data_inicio": saida_br,
                    "raw_data_fim": retorno_br
                }
            })

        render_master_calendar(events)

    # -------------------------------------------------------------------------
    # ABA 2: TABELA DE VIAGENS
    # -------------------------------------------------------------------------
    elif selected_subtab == "📋 Tabela de Viagens":
        st.sidebar.markdown("## 🔍 Filtros da Tabela")

        # Filtro de Técnicos
        tecnicos_unicos = set()
        for quem in df_viagens["quem_foi"].dropna().unique():
            partes = [p.strip() for p in quem.replace("/", ",").split(",") if p.strip()]
            tecnicos_unicos.update(partes)

        filtro_tecnico_tab = st.sidebar.selectbox("👤 Técnico / Membro", ["Todos"] + sorted(list(tecnicos_unicos)), key="f_viagem_tecnico_tab")
        search_viagem = st.sidebar.text_input("🔎 Buscar (Localidade, Chamado, Técnico)", "", key="f_viagem_search").strip().lower()

        items_per_page = render_items_per_page_selector(
            key_prefix="viagens_tab",
            options=[10, 25, 50, 100, "Todos"],
            default_index=1,
            label="📄 Viagens por página:"
        )

        df_filtered = df_viagens.copy()
        if filtro_tecnico_tab != "Todos":
            df_filtered = df_filtered[df_filtered["quem_foi"].str.contains(filtro_tecnico_tab, case=False, na=False)]

        if search_viagem:
            mask = (
                df_filtered["localidade"].str.lower().str.contains(search_viagem, na=False) |
                df_filtered["quem_foi"].str.lower().str.contains(search_viagem, na=False) |
                df_filtered["chamado"].str.lower().str.contains(search_viagem, na=False) |
                df_filtered["saida_br"].str.lower().str.contains(search_viagem, na=False) |
                df_filtered["retorno_br"].str.lower().str.contains(search_viagem, na=False)
            )
            df_filtered = df_filtered[mask]

        # CARDS KPI
        k1, k2, k3 = st.columns(3)
        with k1:
            st.markdown(f"""
                <div class="metric-card" style="border-left-color: #06b6d4;">
                    <div class="metric-title">TOTAL DE VIAGENS</div>
                    <div class="metric-value" style="color: #06b6d4;">{len(df_filtered)}</div>
                </div>
            """, unsafe_allow_html=True)
        with k2:
            comarcas_unicas = len(df_filtered["localidade"].dropna().unique())
            st.markdown(f"""
                <div class="metric-card" style="border-left-color: #10b981;">
                    <div class="metric-title">COMARCAS / DESTINOS</div>
                    <div class="metric-value" style="color: #10b981;">{comarcas_unicas}</div>
                </div>
            """, unsafe_allow_html=True)
        with k3:
            chamados_atendidos = len(df_filtered[df_filtered["chamado"].str.strip() != ""])
            st.markdown(f"""
                <div class="metric-card" style="border-left-color: #3b82f6;">
                    <div class="metric-title">COM CHAMADO REGISTRADO</div>
                    <div class="metric-value" style="color: #3b82f6;">{chamados_atendidos}</div>
                </div>
            """, unsafe_allow_html=True)

        st.markdown("<br>", unsafe_allow_html=True)

        h_col1, h_col2 = st.columns([3, 1])
        with h_col1:
            st.subheader(f"📋 Registros de Viagens ({len(df_filtered)} encontrados)")
        with h_col2:
            buffer = io.BytesIO()
            with pd.ExcelWriter(buffer, engine='openpyxl') as writer:
                df_filtered.to_excel(writer, index=False, sheet_name='Viagens')
            buffer.seek(0)
            st.download_button(
                label="📥 Exportar Excel",
                data=buffer,
                file_name=f"viagens_bancada_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                width='stretch'
            )

        paginated_df, current_page, total_pages, total_items = paginate_items(
            df_filtered,
            items_per_page=items_per_page,
            page_key="viagens_table_page"
        )

        cols_display = ["localidade", "quem_foi", "chamado", "saida_br", "retorno_br"]
        cols_display = [c for c in cols_display if c in paginated_df.columns]

        st.dataframe(
            paginated_df[cols_display],
            column_config={
                "localidade": st.column_config.TextColumn("📍 Destino / Localidade"),
                "quem_foi": st.column_config.TextColumn("👤 Quem foi"),
                "chamado": st.column_config.TextColumn("🎫 Chamado(s)"),
                "saida_br": st.column_config.TextColumn("📅 Saída"),
                "retorno_br": st.column_config.TextColumn("🏁 Retorno"),
            },
            hide_index=True,
            width='stretch'
        )

        render_pagination_controls(
            page_key="viagens_table_page",
            current_page=current_page,
            total_pages=total_pages,
            total_items=total_items,
            items_per_page=items_per_page
        )
