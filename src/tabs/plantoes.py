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


def render_fullcalendar(events: list[dict]):
    """
    Renderiza o componente interativo FullCalendar (v6) com ajuste dinâmico de altura (height: auto),
    tema Dark Glassmorphism, modal customizado e máscaras no padrão brasileiro.
    """
    import json
    events_json = json.dumps(events, ensure_ascii=False)

    calendar_html = f"""
    <!DOCTYPE html>
    <html lang="pt-br">
    <head>
      <meta charset="utf-8"/>
      <script src="https://cdn.jsdelivr.net/npm/fullcalendar@6.1.10/index.global.min.js"></script>
      <style>
        html, body {{
          margin: 0;
          padding: 0;
          height: 100%;
          overflow: hidden;
          font-family: 'Segoe UI', Roboto, Helvetica, Arial, sans-serif;
          transition: background-color 0.3s ease, color 0.3s ease;
        }}

        /* TEMA ESCURO (DARK MODE) */
        body, body.dark-mode {{
          background-color: #0e1117;
          color: #fafafa;
        }}
        body.dark-mode .fc {{
          --fc-border-color: #31333f;
          --fc-page-bg-color: #0e1117;
          --fc-neutral-bg-color: #1a1c24;
          --fc-list-event-hover-bg-color: #262730;
          --fc-today-bg-color: rgba(230, 126, 34, 0.15);
        }}
        body.dark-mode .fc-toolbar-title {{ color: #ffffff !important; }}
        body.dark-mode .fc-col-header-cell-cushion,
        body.dark-mode .fc-daygrid-day-number {{ color: #ffffff !important; }}
        body.dark-mode .fc-button-primary {{
          background-color: #262730 !important;
          border-color: #41444c !important;
          color: #ffffff !important;
        }}
        body.dark-mode .fc-button-primary:hover {{
          background-color: #31333f !important;
        }}
        body.dark-mode .fc-list-day-cushion {{
          background-color: #2a2b36 !important;
          color: #ffffff !important;
        }}
        body.dark-mode .modal-content {{
          background: #1e2129;
          border: 1px solid #363945;
          color: #ffffff;
        }}

        /* TEMA CLARO (LIGHT MODE) */
        body.light-mode {{
          background-color: #ffffff;
          color: #0f172a;
        }}
        body.light-mode .fc {{
          --fc-border-color: #e2e8f0;
          --fc-page-bg-color: #ffffff;
          --fc-neutral-bg-color: #f8fafc;
          --fc-list-event-hover-bg-color: #f1f5f9;
          --fc-today-bg-color: rgba(230, 126, 34, 0.15);
        }}
        body.light-mode .fc-toolbar-title {{ color: #0f172a !important; }}
        body.light-mode .fc-col-header-cell-cushion,
        body.light-mode .fc-daygrid-day-number {{ color: #0f172a !important; }}
        body.light-mode .fc-button-primary {{
          background-color: #f1f5f9 !important;
          border-color: #cbd5e1 !important;
          color: #0f172a !important;
        }}
        body.light-mode .fc-button-primary:hover {{
          background-color: #e2e8f0 !important;
          color: #0f172a !important;
        }}
        body.light-mode .fc-button-active {{
          background-color: #ff4b4b !important;
          border-color: #ff4b4b !important;
          color: #ffffff !important;
        }}
        body.light-mode .fc-list-day-cushion {{
          background-color: #f1f5f9 !important;
          color: #0f172a !important;
        }}
        body.light-mode .modal-content {{
          background: #ffffff;
          border: 1px solid #cbd5e1;
          color: #0f172a;
        }}
        body.light-mode .modal-body p {{ color: #334155; }}
        body.light-mode .modal-body strong {{ color: #0f172a; }}
        body.light-mode .close-btn {{ color: #64748b; }}

        #calendar {{
          max-width: 100%;
          height: 100%;
          box-sizing: border-box;
          padding: 10px;
        }}
        .fc-scroller {{
          overflow-y: auto !important;
        }}
        .fc-event {{
          cursor: pointer;
          border-radius: 4px;
          padding: 2px 4px;
          font-size: 0.85rem;
          box-shadow: 0 2px 4px rgba(0,0,0,0.2);
          transition: transform 0.1s ease;
        }}
        .fc-event:hover {{
          transform: scale(1.02);
        }}


        /* GLASSMORPHISM MODAL OVERLAY */
        .modal-overlay {{
          display: none;
          position: fixed;
          top: 0;
          left: 0;
          width: 100%;
          height: 100%;
          background: rgba(0, 0, 0, 0.75);
          backdrop-filter: blur(5px);
          z-index: 9999;
          justify-content: center;
          align-items: center;
        }}
        .modal-content {{
          background: #1e2129;
          border: 1px solid #363945;
          border-radius: 12px;
          width: 90%;
          max-width: 520px;
          padding: 24px;
          box-shadow: 0 10px 30px rgba(0,0,0,0.5);
          color: #ffffff;
          animation: fadeIn 0.2s ease-out;
        }}
        @keyframes fadeIn {{
          from {{ opacity: 0; transform: translateY(-10px); }}
          to {{ opacity: 1; transform: translateY(0); }}
        }}
        .modal-header {{
          display: flex;
          justify-content: space-between;
          align-items: center;
          border-bottom: 1px solid #363945;
          padding-bottom: 12px;
          margin-bottom: 16px;
        }}
        .modal-title {{
          margin: 0;
          font-size: 1.15rem;
          font-weight: 600;
          color: #ff4b4b;
        }}
        .close-btn {{
          background: transparent;
          border: none;
          color: #aaaaaa;
          font-size: 1.5rem;
          cursor: pointer;
          line-height: 1;
        }}
        .close-btn:hover {{
          color: #ffffff;
        }}
        .modal-body p {{
          margin: 10px 0;
          font-size: 0.95rem;
          line-height: 1.5;
          color: #d1d5db;
        }}
        .modal-body strong {{
          color: #ffffff;
        }}
        .modal-badge {{
          display: inline-block;
          padding: 4px 10px;
          border-radius: 12px;
          font-size: 0.75rem;
          font-weight: 600;
          margin-bottom: 12px;
        }}
        .badge-matutino {{
          background-color: rgba(230, 126, 34, 0.2);
          color: #e67e22;
          border: 1px solid #e67e22;
        }}
        .badge-semanal {{
          background-color: rgba(142, 68, 173, 0.2);
          color: #9b59b6;
          border: 1px solid #9b59b6;
        }}
        .modal-footer {{
          margin-top: 20px;
          text-align: right;
        }}
        .btn-dismiss {{
          background: #ff4b4b;
          color: #fff;
          border: none;
          padding: 8px 16px;
          border-radius: 6px;
          cursor: pointer;
          font-weight: 500;
        }}
        .btn-dismiss:hover {{
          background: #e03e3e;
        }}
      </style>
    </head>
    <body>
      <div id="calendar"></div>

      <!-- COMPONENTE MODAL CUSTOMIZADO -->
      <div id="plantaoModal" class="modal-overlay">
        <div class="modal-content">
          <div class="modal-header">
            <h3 id="mTitle" class="modal-title">Detalhes do Plantão</h3>
            <button class="close-btn" onclick="closeModal()">&times;</button>
          </div>
          <div class="modal-body">
            <span id="mBadge" class="modal-badge"></span>
            <p><strong>📅 Data Início:</strong> <span id="mStart"></span></p>
            <p id="pEnd"><strong>🏁 Data Fim:</strong> <span id="mEnd"></span></p>
            <div style="margin-top: 14px; padding-top: 10px; border-top: 1px solid #363945;">
              <p style="margin-bottom: 6px;"><strong>👤 Plantonista(s):</strong></p>
              <div id="mServidor" style="padding-left: 8px; color: #e2e8f0; font-size: 0.93rem; line-height: 1.6;"></div>
            </div>
            <p id="pTelefone" style="margin-top: 12px;"><strong>📞 Contato:</strong> <span id="mTelefone"></span></p>
          </div>
          <div class="modal-footer">
            <button class="btn-dismiss" onclick="closeModal()">Fechar</button>
          </div>
        </div>
      </div>

      <script>
        function updateThemeFromParent() {{
          var isLight = false;
          try {{
            var parentBody = window.parent.document.body;
            var parentApp = window.parent.document.querySelector('.stApp');
            var themeAttr = (parentBody && parentBody.getAttribute('data-theme')) || 
                            (parentApp && parentApp.getAttribute('data-theme'));
            
            if (themeAttr === 'light') {{
              isLight = true;
            }} else if (themeAttr === 'dark') {{
              isLight = false;
            }} else {{
              isLight = window.parent.matchMedia('(prefers-color-scheme: light)').matches;
            }}
          }} catch(e) {{
            isLight = window.matchMedia('(prefers-color-scheme: light)').matches;
          }}

          if (isLight) {{
            document.body.className = 'light-mode';
          }} else {{
            document.body.className = 'dark-mode';
          }}
        }}

        function closeModal() {{
          document.getElementById('plantaoModal').style.display = 'none';
        }}

        window.onclick = function(event) {{
          var modal = document.getElementById('plantaoModal');
          if (event.target == modal) {{
            modal.style.display = 'none';
          }}
        }};

        function formatBrDateTime(dateObj, rawStr) {{
          if (rawStr && typeof rawStr === 'string' && rawStr.includes('/')) {{
            return rawStr;
          }}
          if (!dateObj) return rawStr || '';
          
          var d = new Date(dateObj);
          if (isNaN(d.getTime())) return rawStr || '';
          
          var day = String(d.getDate()).padStart(2, '0');
          var month = String(d.getMonth() + 1).padStart(2, '0');
          var year = d.getFullYear();
          var hours = String(d.getHours()).padStart(2, '0');
          var minutes = String(d.getMinutes()).padStart(2, '0');
          var seconds = String(d.getSeconds()).padStart(2, '0');
          
          return day + '/' + month + '/' + year + ' ' + hours + ':' + minutes + ':' + seconds;
        }}

        document.addEventListener('DOMContentLoaded', function() {{
          updateThemeFromParent();
          setInterval(updateThemeFromParent, 1000);

          var calendarEl = document.getElementById('calendar');
          var calendar = new FullCalendar.Calendar(calendarEl, {{
            initialView: 'dayGridMonth',
            height: '100%',
            locale: 'pt-br',

            headerToolbar: {{
              left: 'prev,next today',
              center: 'title',
              right: 'dayGridMonth,timeGridWeek,timeGridDay,listMonth'
            }},
            buttonText: {{
              today:    'Hoje',
              month:    'Mês',
              week:     'Semana',
              day:      'Dia',
              list:     'Lista'
            }},
            events: {events_json},
            eventClick: function(info) {{
              var props = info.event.extendedProps;
              
              document.getElementById('mTitle').innerText = info.event.title;
              
              var startFormatted = formatBrDateTime(info.event.start, props.raw_data_inicio);
              var endFormatted = formatBrDateTime(info.event.end, props.raw_data_fim);
              
              document.getElementById('mStart').innerText = startFormatted;
              
              var mEnd = document.getElementById('mEnd');
              var pEnd = document.getElementById('pEnd');
              if (endFormatted) {{
                mEnd.innerText = endFormatted;
                pEnd.style.display = 'block';
              }} else {{
                pEnd.style.display = 'none';
              }}
              
              var mServ = document.getElementById('mServidor');
              if (props.detailsHtml) {{
                mServ.innerHTML = props.detailsHtml;
              }} else {{
                mServ.innerText = props.servidor || 'Não informado';
              }}
              
              var mTel = document.getElementById('mTelefone');
              var pTel = document.getElementById('pTelefone');
              if (props.telefone && props.telefone.trim() !== '') {{
                mTel.innerText = props.telefone;
                pTel.style.display = 'block';
              }} else {{
                pTel.style.display = 'none';
              }}
              
              var badge = document.getElementById('mBadge');
              badge.innerText = props.tipo || 'Plantão';
              if (props.tipo && props.tipo.includes('Matutino')) {{
                badge.className = 'modal-badge badge-matutino';
              }} else {{
                badge.className = 'modal-badge badge-semanal';
              }}
              
              document.getElementById('plantaoModal').style.display = 'flex';
            }}
          }});
          calendar.render();
        }});
      </script>
    </body>
    </html>
    """
    components.html(calendar_html, height=860)


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
                subprocess.Popen([sys.executable, "src/plantoes_scraper.py"])
                st.session_state["was_plantoes_syncing"] = True
                st.toast("🚀 Robô de plantões iniciado em segundo plano!", icon="🤖")
                st.cache_data.clear()
                st.rerun()

    if plantoes_ativo:
        with st.expander("🤖 Robô de Plantões em Segundo Plano – Acompanhar Progresso", expanded=False):
            st.info("O robô está conectando aos portais e sincronizando as escalas neste momento. O uso da aplicação permanece livre!")
            logs = read_plantoes_last_log_lines(15)
            st.code(logs, language="text")
            st.button("🔄 Atualizar Log de Plantões", help="Recarrega as últimas linhas de log")

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
                    "servidor": servidor,
                    "telefone": tel_formatted,
                    "tipo": "Plantão Matutino (08h às 15h)",
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
    PLANTOES_SUBTAB_REVERSE = {v: k for k, v in PLANTOES_SUBTAB_MAP.items()}

    url_subtab = st.query_params.get("subtab", "agenda")
    default_title = PLANTOES_SUBTAB_MAP.get(url_subtab, "📅 Agenda / Calendário Interativo")
    options = list(PLANTOES_SUBTAB_MAP.values())
    default_idx = options.index(default_title) if default_title in options else 0

    selected_subtab = st.radio(
        "Navegação do Plantão:",
        options=options,
        index=default_idx,
        horizontal=True,
        label_visibility="collapsed",
        key="plantoes_subtab_radio"
    )

    new_slug = PLANTOES_SUBTAB_REVERSE.get(selected_subtab, "agenda")
    if st.query_params.get("subtab") != new_slug:
        st.query_params["subtab"] = new_slug

    st.markdown("<br>", unsafe_allow_html=True)

    if selected_subtab == "📅 Agenda / Calendário Interativo":
        render_fullcalendar(events)

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

            df_page_mat, current_page_mat, total_pages_mat, total_items_mat = paginate_items(
                df_disp_renamed,
                page_key="plantoes_matutino",
                items_per_page=items_per_page
            )

            st.dataframe(
                df_page_mat,
                use_container_width=True,
                hide_index=True
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

            df_page_sem, current_page_sem, total_pages_sem, total_items_sem = paginate_items(
                df_disp_s_renamed,
                page_key="plantoes_semanal",
                items_per_page=items_per_page
            )

            st.dataframe(
                df_page_sem,
                use_container_width=True,
                hide_index=True
            )

            render_pagination_controls(
                page_key="plantoes_semanal",
                current_page=current_page_sem,
                total_pages=total_pages_sem,
                total_items=total_items_sem,
                items_per_page=items_per_page
            )


