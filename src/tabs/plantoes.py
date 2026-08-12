import re
import sys
import subprocess
import pandas as pd
import streamlit as st
import streamlit.components.v1 as components
from pathlib import Path
from datetime import datetime

root_dir = Path(__file__).parent.parent.parent
sys.path.insert(0, str(root_dir))

from src.database import get_plantoes_matutino, get_plantoes_semanal
from src.plantoes_scraper import check_plantoes_sync_running, read_plantoes_last_log_lines
from src.components.status_banner import render_log_expander
from src.components.subtabs import render_subtabs
from src.components.calendar import render_master_calendar
from src.components.pagination import (
    render_items_per_page_selector,
    paginate_items,
    render_pagination_controls
)



def format_phone_number(phone_raw: str) -> str:
    """
    Formata número de telefone para o padrão brasileiro com hífen antes dos 4 últimos dígitos.
    Exemplo: +55 67 991455446 -> +55 67 99145-5446
    """
    if not phone_raw or pd.isna(phone_raw):
        return ""
        
    s = str(phone_raw).strip()
    digits = re.sub(r'\D', '', s)
    
    if len(digits) == 13 and digits.startswith("55"):
        return f"+{digits[:2]} {digits[2:4]} {digits[4:9]}-{digits[9:]}"
    elif len(digits) == 11:
        return f"({digits[:2]}) {digits[2:7]}-{digits[7:]}"
    elif len(digits) == 9:
        return f"{digits[:5]}-{digits[5:]}"
    
    if not "-" in s and len(s) > 4:
        return f"{s[:-4]}-{s[-4:]}"
        
    return s


def is_bancada_member(nome_str: str) -> bool:
    """Verifica se o nome contém algum dos integrantes da bancada (Paulo, Reginaldo, Luiz, Murillo)."""
    if not nome_str:
        return False
    n = str(nome_str).lower()
    return "paulo henrique" in n or "reginaldo" in n or "luiz leonardo" in n or "villalba" in n or "murillo" in n or "yazbek" in n





def render_plantoes_page():
    """Renderiza a página principal dos Plantões da Bancada."""
    col_t, col_b = st.columns([3, 1])
    with col_t:
        st.title("📅 Escala de Plantões da Bancada")
        st.write("Acompanhe as escalas de Plantão Matutino (PGJ) e Plantão Semanal (SIMP) dos integrantes da equipe.")

    plantoes_ativo = check_plantoes_sync_running()

    if "was_plantoes_syncing" not in st.session_state:
        st.session_state["was_plantoes_syncing"] = False

    if st.session_state["was_plantoes_syncing"] and not plantoes_ativo:
        st.session_state["was_plantoes_syncing"] = False
        st.toast("🎉 Sincronização de plantões concluída com sucesso! Atualizando agenda...", icon="✅")
        st.cache_data.clear()
        st.rerun()

    if plantoes_ativo:
        st.session_state["was_plantoes_syncing"] = True

    with col_b:
        st.markdown("<div style='height: 15px;'></div>", unsafe_allow_html=True)
        if plantoes_ativo:
            st.button("🤖 Sincronizando...", use_container_width=True, disabled=True)
        else:
            if st.button("🔄 Sincronizar Tudo", type="primary", use_container_width=True, help="Executa sincronização completa do Matutino e SIMP em segundo plano."):
                import time
                subprocess.Popen([sys.executable, "src/plantoes_scraper.py"])
                time.sleep(0.8)
                st.session_state["was_plantoes_syncing"] = True
                st.toast("🚀 Robô de plantões iniciado em segundo plano!", icon="🤖")
                st.cache_data.clear()
                st.rerun()

    render_log_expander(
        "🤖 Robô de Plantões em Segundo Plano – Acompanhar Progresso",
        plantoes_ativo,
        read_plantoes_last_log_lines,
        check_plantoes_sync_running,
        "O robô está conectando aos portais e sincronizando as escalas neste momento. O uso da aplicação permanece livre!"
    )

    st.markdown("---")

    # --- FILTROS SIDEBAR ---
    st.sidebar.markdown("## 🔍 Filtros de Plantão")
    
    opcoes_servidores = [
        "🟢 Apenas Bancada (Paulo, Reginaldo, Luiz, Murillo)",
        "🌐 Todos os Servidores da STI"
    ]
    selected_servidor_mode = st.sidebar.radio("👥 Servidores Exibidos:", opcoes_servidores)
    bancada_only = "Apenas Bancada" in selected_servidor_mode
    
    anos_disponiveis = [2026, 2025, 2024]
    selected_ano = st.sidebar.selectbox("📅 Selecionar Ano:", anos_disponiveis)
    
    items_per_page = render_items_per_page_selector(
        key_prefix="plantoes",
        options=[10, 20, 50, 100, "Todos"],
        default_index=1,
        label="📄 Escalas por página:"
    )

    df_matutino = get_plantoes_matutino(selected_ano)
    df_semanal = get_plantoes_semanal(selected_ano)


    if not df_matutino.empty and 'telefone' in df_matutino.columns:
        df_matutino['telefone'] = df_matutino['telefone'].apply(format_phone_number)

    # --- PREPARAÇÃO DOS EVENTOS DO CALENDÁRIO ---
    events = []
    
    if not df_matutino.empty:
        for _, row in df_matutino.iterrows():
            servidor = str(row['servidor']).strip()
            if bancada_only and not is_bancada_member(servidor):
                continue
                
            dt_iso = str(row['data_iso']).strip()
            if not dt_iso:
                continue
                
            tel_formatted = format_phone_number(str(row.get('telefone', '')))
            
            dt_ini_str = f"{dt_iso}T08:00:00"
            dt_fim_str = f"{dt_iso}T15:00:00"
            
            try:
                d_obj = datetime.strptime(dt_iso, "%Y-%m-%d")
                raw_ini_br = d_obj.strftime("%d/%m/%Y 08:00:00")
                raw_fim_br = d_obj.strftime("%d/%m/%Y 15:00:00")
            except Exception:
                raw_ini_br = ""
                raw_fim_br = ""

            events.append({
                "title": f"☀️ Matutino: {servidor.split()[0]} ({servidor.split()[-1]})",
                "start": dt_ini_str,
                "end": dt_fim_str,
                "backgroundColor": "#e67e22",
                "borderColor": "#d35400",
                "extendedProps": {
                    "categoria_evento": "plantao",
                    "servidor": servidor,
                    "telefone": tel_formatted,
                    "tipo": "Plantão Matutino PGJ (08h-15h)",
                    "raw_data_inicio": raw_ini_br,
                    "raw_data_fim": raw_fim_br
                }
            })
            
    if not df_semanal.empty:
        for _, row in df_semanal.iterrows():
            manut = str(row.get('manutencao', '')).strip()
            sdesk = str(row.get('service_desk', '')).strip()
            infra = str(row.get('infraestrutura', '')).strip()
            dev = str(row.get('desenvolvimento', '')).strip()
            
            dt_ini_raw = str(row.get('data_inicio', '')).strip()
            dt_fim_raw = str(row.get('data_fim', '')).strip()
            
            dt_ini = dt_ini_raw.replace(" ", "T")
            dt_fim = dt_fim_raw.replace(" ", "T")
            
            raw_ini_br = ""
            raw_fim_br = ""
            try:
                if dt_ini_raw:
                    dt_o = datetime.strptime(dt_ini_raw, "%Y-%m-%d %H:%M:%S")
                    raw_ini_br = dt_o.strftime("%d/%m/%Y %H:%M:%S")
                if dt_fim_raw:
                    dt_f = datetime.strptime(dt_fim_raw, "%Y-%m-%d %H:%M:%S")
                    raw_fim_br = dt_f.strftime("%d/%m/%Y %H:%M:%S")
            except Exception:
                pass

            if bancada_only:
                bancada_na_escala = [s for s in [manut, sdesk, infra, dev] if is_bancada_member(s)]
                if not bancada_na_escala:
                    continue
                display_name = ", ".join([s.split()[0] for s in bancada_na_escala])
            else:
                display_name = manut.split()[0] if manut else "STI"
                
            details_html = (
                f"<b>Manutenção:</b> {manut if manut else 'Não informado'}<br>"
                f"<b>Service Desk:</b> {sdesk if sdesk else 'Não informado'}<br>"
                f"<b>Infraestrutura:</b> {infra if infra else 'Não informado'}<br>"
                f"<b>Desenvolvimento:</b> {dev if dev else 'Não informado'}"
            )
            
            if dt_ini:
                events.append({
                    "title": f"🌙 Plantão Semanal: {display_name}",
                    "start": dt_ini,
                    "end": dt_fim if dt_fim else dt_ini,
                    "backgroundColor": "#8e44ad",
                    "borderColor": "#9b59b6",
                    "extendedProps": {
                        "categoria_evento": "plantao",
                        "servidor": f"Manutenção: {manut} | Service Desk: {sdesk} | Infra: {infra} | Dev: {dev}",
                        "detailsHtml": details_html,
                        "tipo": "Plantão Semanal SIMP",
                        "raw_data_inicio": raw_ini_br,
                        "raw_data_fim": raw_fim_br
                    }
                })

    # ABAS INTERNAS DE VISUALIZAÇÃO COM QUERY PARAMETERS (?subtab=slug)
    PLANTOES_SUBTAB_MAP = {
        "agenda": "📅 Agenda / Calendário Interativo",
        "matutino": "☀️ Plantão Matutino PGJ (08h-15h)",
        "semanal": "🌃 Plantão Semanal SIMP"
    }

    selected_subtab = render_subtabs(PLANTOES_SUBTAB_MAP, default_slug="agenda", key="plantoes_subtab_radio")

    st.markdown("<br>", unsafe_allow_html=True)

    if selected_subtab == "📅 Agenda / Calendário Interativo":
        render_master_calendar(events, height_px=860, scrolling_enabled=True)

    elif selected_subtab == "☀️ Plantão Matutino PGJ (08h-15h)":
        st.markdown("### ☀️ Escala do Plantão Matutino (PGJ)")
        if df_matutino.empty:
            st.info("Nenhum registro de plantão matutino cadastrado no banco.")
        else:
            df_disp = df_matutino.copy()
            if bancada_only:
                df_disp = df_disp[df_disp['servidor'].apply(is_bancada_member)]
            
            df_disp_renamed = df_disp.rename(columns={
                "data_iso": "Data ISO",
                "dia_semana": "Dia da Semana",
                "servidor": "Servidor",
                "telefone": "Telefone / Contato",
                "ano": "Ano"
            })
            df_disp_renamed['Data ISO'] = pd.to_datetime(df_disp_renamed['Data ISO'], errors='coerce')

            df_page_mat, current_page_mat, total_pages_mat, total_items_mat = paginate_items(
                df_disp_renamed,
                page_key="plantoes_matutino",
                items_per_page=items_per_page
            )

            st.dataframe(
                df_page_mat,
                use_container_width=True,
                hide_index=True,
                column_config={
                    "Data ISO": st.column_config.DateColumn("Data", format="DD/MM/YYYY")
                }
            )

            render_pagination_controls(
                page_key="plantoes_matutino",
                current_page=current_page_mat,
                total_pages=total_pages_mat,
                total_items=total_items_mat,
                items_per_page=items_per_page
            )

    elif selected_subtab == "🌃 Plantão Semanal SIMP":
        st.markdown("### 🌃 Escala do Plantão Semanal STI (SIMP)")
        if df_semanal.empty:
            st.info("Nenhum registro de plantão semanal cadastrado no banco. Clique em 'Sincronizar Tudo' para buscar do SIMP.")
        else:
            df_disp_s = df_semanal.copy()
            if bancada_only:
                df_disp_s = df_disp_s[
                    df_disp_s['manutencao'].apply(is_bancada_member) |
                    df_disp_s['service_desk'].apply(is_bancada_member) |
                    df_disp_s['infraestrutura'].apply(is_bancada_member) |
                    df_disp_s['desenvolvimento'].apply(is_bancada_member)
                ]
            
            df_disp_s_renamed = df_disp_s.rename(columns={
                "mes": "Mês",
                "periodo_str": "Período do Plantão",
                "data_inicio": "Início",
                "data_fim": "Término",
                "service_desk": "Service Desk",
                "manutencao": "Manutenção",
                "infraestrutura": "Infraestrutura",
                "desenvolvimento": "Desenvolvimento",
                "ano": "Ano"
            })
            df_disp_s_renamed['Início'] = pd.to_datetime(df_disp_s_renamed['Início'], errors='coerce')
            df_disp_s_renamed['Término'] = pd.to_datetime(df_disp_s_renamed['Término'], errors='coerce')

            df_page_sem, current_page_sem, total_pages_sem, total_items_sem = paginate_items(
                df_disp_s_renamed,
                page_key="plantoes_semanal",
                items_per_page=items_per_page
            )

            st.dataframe(
                df_page_sem,
                use_container_width=True,
                hide_index=True,
                column_config={
                    "Início": st.column_config.DatetimeColumn("Início", format="DD/MM/YYYY HH:mm:ss"),
                    "Término": st.column_config.DatetimeColumn("Término", format="DD/MM/YYYY HH:mm:ss")
                }
            )

            render_pagination_controls(
                page_key="plantoes_semanal",
                current_page=current_page_sem,
                total_pages=total_pages_sem,
                total_items=total_items_sem,
                items_per_page=items_per_page
            )


