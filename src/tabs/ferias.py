import os
import sys
import time
import subprocess
import pandas as pd
import streamlit as st
from datetime import datetime, timedelta
from pathlib import Path

root_dir = Path(__file__).parent.parent.parent
sys.path.insert(0, str(root_dir))

from src.config import _cfg, FERIAS_EXCEL_RELATIVE_PATH
from src.database import (
    get_ferias_df,
    get_ferias_membros,
    get_ferias_anos,
    sync_ferias_from_excel
)
from src.syncs.sync_ferias import (
    check_ferias_sync_running,
    read_ferias_last_log_lines
)
from src.components.status_banner import render_log_expander
from src.components.subtabs import render_subtabs
from src.components.calendar import render_master_calendar
from src.components.pagination import (
    render_items_per_page_selector,
    paginate_items,
    render_pagination_controls
)
from src.components.metric_cards import render_metric_card


PALETA_MEMBROS = {
    "Paulo Rezende": "#10b981",       # Esmeralda
    "Reginaldo Bandeira": "#3b82f6",  # Azul
    "Luiz Villalba": "#f59e0b",       # Âmbar
    "Alex": "#8b5cf6",                # Roxo
    "Marcos Larrea (Luppa)": "#ec4899",# Rosa
    "Matheus (Luppa)": "#14b8a6",     # Ciano
    "Rafael (Luppa)": "#f97316",      # Laranja
    "Rian (Luppa)": "#6366f1",        # Índigo
    "Estagiário": "#94a3b8"           # Cinza
}

def get_cor_membro(membro: str) -> str:
    """Retorna cor de destaque para cada membro da equipe."""
    for k, v in PALETA_MEMBROS.items():
        if k.lower() in membro.lower() or membro.lower() in k.lower():
            return v
    return "#3b82f6"


@st.dialog("⚙️ Configurar / Enviar Planilha de Férias", width="large")
def modal_config_ferias():
    """Modal (@st.dialog) para gerenciar link do SharePoint e upload manual da planilha de férias."""
    st.markdown("### 🏖️ Gestão da Planilha de Previsão de Férias")
    st.caption("Consulte a planilha oficial no SharePoint ou envie uma cópia (.xlsx) diretamente:")

    tab_online, tab_upload = st.tabs(["🌐 Link SharePoint / Atualizar Online", "📥 Envio Direto de Planilha"])

    with tab_online:
        from dotenv import load_dotenv
        load_dotenv(override=True)
        excel_url = (
            os.getenv("FERIAS_EXCEL_RELATIVE_PATH", "")
            or _cfg("FERIAS_EXCEL_RELATIVE_PATH")
            or FERIAS_EXCEL_RELATIVE_PATH
            or "https://ministeriopublicoms.sharepoint.com/:x:/r/sites/dit-manutencao/_layouts/15/Doc.aspx?sourcedoc=%7BE197F2AD-7143-4E92-A56B-B049D930E4C5%7D&file=Previs%C3%A3o%20de%20F%C3%A9rias-Manutencao.xlsx&action=default&mobileredirect=true&wdwpf=doclib-t"
        ).strip()
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
        else:
            st.info(f"Caminho configurado: `{excel_url}`")

        if st.button("🚀 Sincronizar pelo Link do SharePoint Agora", type="primary", width='stretch'):
            popen_kwargs = {"creationflags": subprocess.CREATE_NO_WINDOW} if sys.platform == "win32" else {}
            subprocess.Popen([sys.executable, "src/syncs/sync_ferias.py"], **popen_kwargs)
            st.toast("🚀 Sincronização disparada com sucesso!", icon="🤖")
            st.rerun()

    with tab_upload:
        st.write("Faça o upload manual do arquivo Excel (.xlsx) da previsão de férias:")
        uploaded_excel = st.file_uploader("Selecione o arquivo Excel (.xlsx)", type=["xlsx", "xls"], key="modal_up_ferias_excel")

        if st.button("⚡ Processar e Gravar no Banco", type="primary", width='stretch'):
            if not uploaded_excel:
                st.warning("Selecione um arquivo Excel primeiro.")
            else:
                with st.spinner("Processando planilha de férias..."):
                    res = sync_ferias_from_excel(uploaded_excel)
                    if res:
                        st.success("🎉 Planilha de férias importada com sucesso!")
                        time.sleep(1.2)
                        st.rerun()
                    else:
                        st.error("Não foi possível processar a planilha. Verifique o formato do arquivo.")


def detectar_sobreposicoes(df: pd.DataFrame) -> list[dict]:
    """Identifica períodos onde 2 ou mais membros da equipe estarão ausentes simultaneamente."""
    if df.empty or len(df) < 2:
        return []

    conflitos = []
    df_valid = df.dropna(subset=["data_inicio_iso", "data_fim_iso"]).copy()
    registros = df_valid.to_dict("records")

    for i in range(len(registros)):
        r1 = registros[i]
        try:
            ini1 = datetime.strptime(r1["data_inicio_iso"], "%Y-%m-%d").date()
            fim1 = datetime.strptime(r1["data_fim_iso"], "%Y-%m-%d").date()
        except Exception:
            continue

        for j in range(i + 1, len(registros)):
            r2 = registros[j]
            # Se for o mesmo membro, não é conflito de equipe
            if r1["membro"].strip() == r2["membro"].strip():
                continue

            try:
                ini2 = datetime.strptime(r2["data_inicio_iso"], "%Y-%m-%d").date()
                fim2 = datetime.strptime(r2["data_fim_iso"], "%Y-%m-%d").date()
            except Exception:
                continue

            # Verifica sobreposição de datas: max(ini1, ini2) <= min(fim1, fim2)
            sobrepos_ini = max(ini1, ini2)
            sobrepos_fim = min(fim1, fim2)

            if sobrepos_ini <= sobrepos_fim:
                dias_conflito = (sobrepos_fim - sobrepos_ini).days + 1
                conflitos.append({
                    "membro1": r1["membro"],
                    "tipo1": r1.get("tipo_escala", "Férias"),
                    "membro2": r2["membro"],
                    "tipo2": r2.get("tipo_escala", "Férias"),
                    "inicio_br": sobrepos_ini.strftime("%d/%m/%Y"),
                    "fim_br": sobrepos_fim.strftime("%d/%m/%Y"),
                    "dias": dias_conflito,
                    "ano": r1.get("ano", ini1.year)
                })

    return conflitos


def render_ferias_page():
    """Renderiza a página principal de Férias e Ausências da Bancada STI."""
    col_t, col_b = st.columns([3, 1])
    with col_t:
        st.title("🏖️ Férias da Bancada STI")
        st.write("Acompanhe o planejamento de férias, compensações e recesso forense da equipe de Manutenção STI.")

    ferias_ativo = check_ferias_sync_running()

    if "was_ferias_syncing" not in st.session_state:
        st.session_state["was_ferias_syncing"] = False

    if st.session_state["was_ferias_syncing"] and not ferias_ativo:
        st.session_state["was_ferias_syncing"] = False
        st.cache_data.clear()
        st.toast("🎉 Sincronização de férias concluída com sucesso!", icon="✅")
        st.rerun()

    if ferias_ativo:
        st.session_state["was_ferias_syncing"] = True

    with col_b:
        st.markdown("<div style='height: 15px;'></div>", unsafe_allow_html=True)
        if ferias_ativo:
            st.button("🤖 Sincronizando...", width='stretch', disabled=True)
        else:
            if st.button("🔄 Sincronizar Tudo", type="primary", width='stretch', help="Executa a sincronização completa da planilha de férias em segundo plano."):
                popen_kwargs = {"creationflags": subprocess.CREATE_NO_WINDOW} if sys.platform == "win32" else {}
                subprocess.Popen([sys.executable, "src/syncs/sync_ferias.py"], **popen_kwargs)
                time.sleep(0.5)
                st.session_state["was_ferias_syncing"] = True
                st.toast("🚀 Sincronização iniciada em segundo plano!", icon="🤖")
                st.rerun()

    render_log_expander(
        "🤖 Sincronização de Férias em Segundo Plano",
        ferias_ativo,
        read_ferias_last_log_lines,
        check_ferias_sync_running,
        "O robô está consultando a planilha oficial do SharePoint. O painel permanece livre para uso!"
    )

    st.markdown("---")

    # Mapeamento de Sub-abas
    FERIAS_SUBTAB_MAP = {
        "escala": "📊 Planilha & Escala",
        "calendario": "📅 Calendário de Férias"
    }

    selected_subtab = render_subtabs(FERIAS_SUBTAB_MAP, default_slug="escala", key="ferias_subtab_radio")
    st.markdown("<br>", unsafe_allow_html=True)

    # Coleta dados base do SQLite
    df_total = get_ferias_df()
    membros_disponiveis = get_ferias_membros()
    anos_disponiveis = get_ferias_anos()

    if df_total.empty:
        st.info("Nenhum registro de férias encontrado no banco relacional local.")
        if st.button("⚙️ Configurar ou Importar Planilha de Férias"):
            modal_config_ferias()
        return

    # -------------------------------------------------------------------------
    # ABA 1: 📊 PLANILHA & ESCALA
    # -------------------------------------------------------------------------
    if selected_subtab == "📊 Planilha & Escala":
        st.sidebar.markdown("## ⚙️ Ações")
        if st.sidebar.button("📥 Importar / Configurar Planilha", width='stretch', help="Fazer upload manual ou configurar link do SharePoint."):
            modal_config_ferias()

        st.sidebar.markdown("---")
        st.sidebar.markdown("## 🔍 Filtros da Tabela")

        filtro_ano = st.sidebar.selectbox("📅 Exercício / Ano:", ["Todos"] + anos_disponiveis, index=1 if len(anos_disponiveis) > 1 else 0, key="f_ano_escala")
        filtro_membro = st.sidebar.selectbox("👤 Membro da Equipe:", ["Todos"] + membros_disponiveis, key="f_membro_escala")
        filtro_tipo = st.sidebar.selectbox("📋 Modalidade:", ["Todas", "Férias", "Compensação/Licença", "Recesso Forense"], key="f_tipo_escala")
        
        items_per_page = render_items_per_page_selector(
            key_prefix="ferias_tabela",
            options=[10, 25, 50, 100, "Todos"],
            default_index=1,
            label="📄 Registros por página:"
        )

        df_filtered = df_total.copy()
        if filtro_ano != "Todos":
            df_filtered = df_filtered[df_filtered["ano"] == int(filtro_ano)]
        if filtro_membro != "Todos":
            df_filtered = df_filtered[df_filtered["membro"] == filtro_membro]
        if filtro_tipo != "Todas":
            tipo_map = {
                "Férias": "ferias",
                "Compensação/Licença": "licencas_compensacoes",
                "Recesso Forense": "recesso_forense"
            }
            target_slug = tipo_map.get(filtro_tipo)
            if target_slug:
                df_filtered = df_filtered[df_filtered["tipo_escala"] == target_slug]

        # KPIs Resumo
        st.subheader("📊 Indicadores da Escala de Ausências")
        kpi_c1, kpi_c2, kpi_c3, kpi_c4 = st.columns(4)

        total_dias = int(df_filtered["dias"].sum()) if not df_filtered.empty else 0
        total_periodos = len(df_filtered)
        membros_com_ferias = df_filtered["membro"].nunique() if not df_filtered.empty else 0
        conflitos_lista = detectar_sobreposicoes(df_filtered)

        with kpi_c1:
            render_metric_card(
                title="📅 Períodos Agendados",
                value=total_periodos,
                border_color="#3b82f6",
                title_color="#3b82f6",
                subtitle="escalas cadastradas",
                text_align="center"
            )
        with kpi_c2:
            render_metric_card(
                title="⏱️ Total de Dias de Ausência",
                value=total_dias,
                border_color="#10b981",
                title_color="#10b981",
                subtitle="dias somados",
                text_align="center"
            )
        with kpi_c3:
            render_metric_card(
                title="👥 Membros com Escala",
                value=membros_com_ferias,
                border_color="#f59e0b",
                title_color="#f59e0b",
                subtitle="servidores/prestadores",
                text_align="center"
            )
        with kpi_c4:
            cor_conflito = "#ef4444" if len(conflitos_lista) > 0 else "#64748b"
            render_metric_card(
                title="⚠️ Sobreposições Detectadas",
                value=len(conflitos_lista),
                border_color=cor_conflito,
                title_color=cor_conflito,
                subtitle="conflitos de ausência simultânea",
                text_align="center"
            )

        st.markdown("<br>", unsafe_allow_html=True)

        # Alerta se houver sobreposições
        if conflitos_lista:
            with st.expander(f"⚠️ Atenção: {len(conflitos_lista)} sobreposição(ões) de período detectada(s) na equipe!", expanded=True):
                for conf in conflitos_lista:
                    st.warning(
                        f"**{conf['membro1']}** e **{conf['membro2']}** estarão ausentes simultaneamente no período de "
                        f"**{conf['inicio_br']} até {conf['fim_br']}** ({conf['dias']} dias de sobreposição)."
                    )

        st.markdown("### 📋 Escala Detalhada de Férias e Licenças")

        if df_filtered.empty:
            st.info("Nenhum registro encontrado para os filtros selecionados.")
        else:
            # Formatação elegante da tabela
            df_disp = df_filtered.copy()
            df_disp["Modalidade"] = df_disp["tipo_escala"].map({
                "ferias": "🏖️ Férias",
                "licencas_compensacoes": "📜 Compensação / Licença",
                "recesso_forense": "⚖️ Recesso Forense"
            }).fillna("🏖️ Férias")

            df_disp_view = df_disp[[
                "ano", "membro", "Modalidade", "mes_nome",
                "periodo_bruto", "data_inicio_br", "data_fim_br", "dias", "status"
            ]].rename(columns={
                "ano": "Exercício",
                "membro": "Membro",
                "mes_nome": "Mês / Referência",
                "periodo_bruto": "Período Indicado",
                "data_inicio_br": "Data Início",
                "data_fim_br": "Data Fim",
                "dias": "Dias",
                "status": "Status"
            })

            # Paginação
            paginated_df, current_page, total_pages, total_items = paginate_items(
                df_disp_view,
                page_key="ferias_tbl",
                items_per_page=items_per_page
            )

            st.dataframe(
                paginated_df,
                use_container_width=True,
                hide_index=True
            )

            render_pagination_controls(
                page_key="ferias_tbl",
                current_page=current_page,
                total_pages=total_pages,
                total_items=total_items,
                items_per_page=items_per_page
            )

            # Exportação
            col_exp1, col_exp2 = st.columns([1, 4])
            with col_exp1:
                csv_bytes = df_disp_view.to_csv(index=False).encode('utf-8-sig')
                st.download_button(
                    label="📥 Exportar CSV",
                    data=csv_bytes,
                    file_name=f"escala_ferias_bancada_{filtro_ano}.csv",
                    mime="text/csv",
                    use_container_width=True
                )

    # -------------------------------------------------------------------------
    # ABA 2: 📅 CALENDÁRIO DE FÉRIAS
    # -------------------------------------------------------------------------
    elif selected_subtab == "📅 Calendário de Férias":
        st.sidebar.markdown("## 🔍 Filtros do Calendário")

        filtro_ano_cal = st.sidebar.selectbox("📅 Ano:", ["Todos"] + anos_disponiveis, index=1 if len(anos_disponiveis) > 1 else 0, key="f_ano_cal")
        filtro_membro_cal = st.sidebar.selectbox("👤 Membro:", ["Todos"] + membros_disponiveis, key="f_membro_cal")
        filtro_modalidade_cal = st.sidebar.selectbox("🏖️ Modalidade:", ["Todas", "Férias", "Compensação/Licença", "Recesso Forense"], key="f_mod_cal")

        df_cal = df_total.copy()
        if filtro_ano_cal != "Todos":
            df_cal = df_cal[df_cal["ano"] == int(filtro_ano_cal)]
        if filtro_membro_cal != "Todos":
            df_cal = df_cal[df_cal["membro"] == filtro_membro_cal]
        if filtro_modalidade_cal != "Todas":
            mod_map = {
                "Férias": "ferias",
                "Compensação/Licença": "licencas_compensacoes",
                "Recesso Forense": "recesso_forense"
            }
            t_slug = mod_map.get(filtro_modalidade_cal)
            if t_slug:
                df_cal = df_cal[df_cal["tipo_escala"] == t_slug]

        st.subheader("📅 Visão Geral da Escala no Calendário")
        st.caption("Eventos coloridos por integrante da equipe. Clique em qualquer evento para ver detalhes da concessão, duração e exercício.")

        events = []
        for _, row in df_cal.iterrows():
            dt_ini_iso = str(row.get("data_inicio_iso", "")).strip()
            dt_fim_iso = str(row.get("data_fim_iso", "")).strip()
            membro = str(row.get("membro", "")).strip()
            tipo_escala = str(row.get("tipo_escala", "ferias"))
            dias = row.get("dias", 0)
            ano = row.get("ano", "")
            status = row.get("status", "Confirmada")

            if not dt_ini_iso:
                continue

            cor = get_cor_membro(membro)

            # Prefixo do título
            if tipo_escala == "recesso_forense":
                titulo = f"⚖️ Recesso: {membro.split()[0]}"
            elif tipo_escala == "licencas_compensacoes":
                titulo = f"📜 Licença: {membro.split()[0]} ({dias}d)"
            else:
                titulo = f"🏖️ Férias: {membro.split()[0]} ({dias}d)"

            # FullCalendar all-day end é exclusivo; adiciona +1 dia ao término
            cal_end = dt_fim_iso if dt_fim_iso else dt_ini_iso
            try:
                dt_fim_obj = datetime.strptime(cal_end, "%Y-%m-%d") + timedelta(days=1)
                cal_end = dt_fim_obj.strftime("%Y-%m-%d")
            except Exception:
                pass

            tipo_label = {
                "ferias": "Férias Regulamentares",
                "licencas_compensacoes": "Compensação de Plantão / Licença",
                "recesso_forense": "Recesso Forense de Fim de Ano"
            }.get(tipo_escala, "Férias")

            events.append({
                "title": titulo,
                "start": dt_ini_iso,
                "end": cal_end,
                "backgroundColor": cor,
                "borderColor": cor,
                "allDay": True,
                "extendedProps": {
                    "categoria_evento": "ferias",
                    "membro": membro,
                    "tipo": titulo.split(":")[0],
                    "tipo_escala_label": tipo_label,
                    "data_inicio_br": row.get("data_inicio_br", dt_ini_iso),
                    "data_fim_br": row.get("data_fim_br", dt_fim_iso),
                    "dias": dias,
                    "ano": ano,
                    "status": status,
                    "cor_hex": cor,
                    "raw_data_inicio": row.get("data_inicio_br", dt_ini_iso),
                    "raw_data_fim": row.get("data_fim_br", dt_fim_iso)
                }
            })

        render_master_calendar(events, height_px=860, scrolling_enabled=True)
