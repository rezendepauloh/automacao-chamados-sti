import json
import re
import pandas as pd
import streamlit as st
import streamlit.components.v1 as components
from datetime import datetime, date, time

from src.database import (
    get_plantoes_matutino,
    get_plantoes_semanal,
    get_garantia_contratos_df,
    load_data,
    save_evento_manual,
    get_eventos_manuais
)
from src.tabs.plantoes import format_phone_number
from src.tabs.garantia import parse_date_to_iso_and_br
from src.tabs.portarias import fetch_portarias_bancada


def parse_ticket_date_iso_and_br(date_val):
    """Auxiliar para converter datas de chamados (ISO ou BR string) em (ISO_str, BR_str)."""
    if pd.isna(date_val) or not date_val or str(date_val).strip() == "":
        return None, None
    s = str(date_val).strip()
    try:
        dt = pd.to_datetime(s, dayfirst=True, errors='coerce')
        if pd.isna(dt):
            return None, None
        return dt.strftime('%Y-%m-%d'), dt.strftime('%d/%m/%Y %H:%M:%S')
    except Exception:
        return None, None


def extract_vacation_dates(text: str, fallback_date_str: str):
    """
    Procura por padrões de período de férias em formato de texto usando Regex.
    Exemplos: 'de 8 a 17.9.2026', '01/10 a 15/10', 'de 05 a 19/08/2026', '10/05/2026 a 25/05/2026'.
    Retorna uma tupla (start_iso, end_iso, start_br, end_br). Se falhar, faz fallback para a data de emissão.
    """
    if not text:
        return None, None, None, None

    # Tenta extrair ano de referência do texto ou fallback_date_str
    current_year = datetime.now().year
    if fallback_date_str:
        try:
            current_year = datetime.strptime(fallback_date_str, "%d/%m/%Y").year
        except Exception:
            pass

    # Padrão 1: "de DD a DD/MM/YYYY" ou "de DD.MM a DD.MM.YYYY" ou "DD/MM a DD/MM/YYYY"
    p1 = r'(?:no\s+período\s+)?(?:de\s+)?(\d{1,2})[\/\.]?(\d{1,2})?\s*(?:a|até|-)\s*(\d{1,2})[\/\.](\d{1,2})[\/\.](\d{4})'
    m1 = re.search(p1, text, re.IGNORECASE)
    if m1:
        d1, m1_m, d2, m2, y2 = m1.groups()
        m1_val = int(m1_m) if m1_m else int(m2)
        try:
            dt1 = datetime(int(y2), m1_val, int(d1))
            dt2 = datetime(int(y2), int(m2), int(d2))
            return dt1.strftime("%Y-%m-%d"), dt2.strftime("%Y-%m-%d"), dt1.strftime("%d/%m/%Y"), dt2.strftime("%d/%m/%Y")
        except Exception:
            pass

    # Padrão 2: "de DD/MM a DD/MM" (sem ano explícito)
    p2 = r'(?:no\s+período\s+)?(?:de\s+)?(\d{1,2})[\/\.](\d{1,2})\s*(?:a|até|-)\s*(\d{1,2})[\/\.](\d{1,2})'
    m2 = re.search(p2, text, re.IGNORECASE)
    if m2:
        d1, m1, d2, m2_m = m2.groups()
        try:
            dt1 = datetime(current_year, int(m1), int(d1))
            dt2 = datetime(current_year, int(m2_m), int(d2))
            return dt1.strftime("%Y-%m-%d"), dt2.strftime("%Y-%m-%d"), dt1.strftime("%d/%m/%Y"), dt2.strftime("%d/%m/%Y")
        except Exception:
            pass

    # Padrão 3: "de DD a DD.MM.YYYY" (ex: "de 8 a 17.9.2026")
    p3 = r'(?:de\s+)?(\d{1,2})\s*a\s*(\d{1,2})[\/\.](\d{1,2})[\/\.](\d{4})'
    m3 = re.search(p3, text, re.IGNORECASE)
    if m3:
        d1, d2, m, y = m3.groups()
        try:
            dt1 = datetime(int(y), int(m), int(d1))
            dt2 = datetime(int(y), int(m), int(d2))
            return dt1.strftime("%Y-%m-%d"), dt2.strftime("%Y-%m-%d"), dt1.strftime("%d/%m/%Y"), dt2.strftime("%d/%m/%Y")
        except Exception:
            pass

    return None, None, None, None


@st.dialog("➕ Novo Evento Manual")
def modal_novo_evento_manual():
    """Modal nativo do Streamlit (@st.dialog) para cadastro de eventos manuais."""
    st.write("Preencha as informações do evento para adicionar ao Calendário Geral:")

    titulo = st.text_input("📌 Título do Evento *", placeholder="Ex: Manutenção Preventiva no Servidor X")
    
    col_d1, col_t1 = st.columns(2)
    with col_d1:
        d_ini = st.date_input("📅 Data de Início *", value=date.today(), format="DD/MM/YYYY")
    with col_t1:
        t_ini = st.time_input("⏰ Hora de Início", value=time(8, 0))

    has_end_date = st.checkbox("Definir Data/Hora de Término")
    d_fim, t_fim = None, None
    if has_end_date:
        col_d2, col_t2 = st.columns(2)
        with col_d2:
            d_fim = st.date_input("📅 Data de Término", value=date.today(), format="DD/MM/YYYY")
        with col_t2:
            t_fim = st.time_input("⏰ Hora de Término", value=time(17, 0))

    autor = st.text_input("👤 Autor / Responsável", value="Bancada STI")
    descricao = st.text_area("📝 Descrição / Detalhes", placeholder="Descreva os detalhes e orientações do evento...")

    if st.button("💾 Salvar Evento", type="primary", use_container_width=True):
        if not titulo.strip():
            st.error("Por favor, preencha o título do evento.")
            return

        dt_inicio_iso = f"{d_ini.strftime('%Y-%m-%d')}T{t_ini.strftime('%H:%M:%S')}"
        dt_fim_iso = ""
        if has_end_date and d_fim:
            dt_fim_iso = f"{d_fim.strftime('%Y-%m-%d')}T{t_fim.strftime('%H:%M:%S')}"

        save_evento_manual(
            titulo=titulo.strip(),
            data_inicio=dt_inicio_iso,
            data_fim=dt_fim_iso,
            descricao=descricao.strip(),
            autor=autor.strip()
        )
        st.toast("✅ Evento manual salvo com sucesso!", icon="🎉")
        st.rerun()


def render_master_calendar(events: list[dict]):
    """
    Renderiza o componente interativo FullCalendar (v6) replicando o tema Dark/Light (Glassmorphism),
    responsividade e o Modal Dinâmico Inteligente com blocos independentes para Registros Manuais, Plantões, Garantia, Chamados e Portarias.
    """
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
          min-height: 100%;
          overflow-y: auto;
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
          border-radius: 12px;
          width: 90%;
          max-width: 540px;
          padding: 24px;
          box-shadow: 0 10px 30px rgba(0,0,0,0.5);
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
          border-bottom: 1px solid rgba(128,128,128,0.3);
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
          font-size: 1.5rem;
          cursor: pointer;
          line-height: 1;
        }}
        .modal-body p {{
          margin: 10px 0;
          font-size: 0.95rem;
          line-height: 1.5;
        }}
        .modal-badge {{
          display: inline-block;
          padding: 4px 10px;
          border-radius: 12px;
          font-size: 0.75rem;
          font-weight: 600;
          margin-bottom: 12px;
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

      <!-- MODAL DINÂMICO INTELIGENTE -->
      <div id="calendarioModal" class="modal-overlay">
        <div class="modal-content">
          <div class="modal-header">
            <h3 id="mTitle" class="modal-title">Detalhes do Evento</h3>
            <button class="close-btn" onclick="closeModal()">&times;</button>
          </div>
          <div class="modal-body">
            <span id="mBadge" class="modal-badge"></span>

            <!-- BLOCO MANUAL -->
            <div id="blocoManual" style="display:none;">
              <p><strong>📌 Título:</strong> <span id="mManualTitle"></span></p>
              <p><strong>📅 Data Início:</strong> <span id="mManualStart"></span></p>
              <p id="mManualEndContainer"><strong>🏁 Data Término:</strong> <span id="mManualEnd"></span></p>
              <p><strong>👤 Autor / Responsável:</strong> <span id="mManualAutor"></span></p>
              <div style="margin-top: 14px; padding-top: 10px; border-top: 1px solid rgba(128,128,128,0.2);">
                <p style="margin-bottom: 6px;"><strong>📝 Descrição / Observações:</strong></p>
                <div id="mManualDesc" style="padding: 10px; background: rgba(128,128,128,0.1); border-radius: 6px; font-size: 0.88rem; line-height: 1.5; max-height: 150px; overflow-y: auto;"></div>
              </div>
            </div>

            <!-- BLOCO PLANTÃO -->
            <div id="blocoPlantao" style="display:none;">
              <p><strong>📅 Data Início:</strong> <span id="pStart"></span></p>
              <p id="pEndContainer"><strong>🏁 Data Fim:</strong> <span id="pEnd"></span></p>
              <div style="margin-top: 14px; padding-top: 10px; border-top: 1px solid rgba(128,128,128,0.2);">
                <p style="margin-bottom: 6px;"><strong>👤 Plantonista(s):</strong></p>
                <div id="pServidor" style="padding-left: 8px; font-size: 0.93rem; line-height: 1.6;"></div>
              </div>
              <p id="pTelefoneContainer" style="margin-top: 12px;"><strong>📞 Contato:</strong> <span id="pTelefone"></span></p>
            </div>

            <!-- BLOCO GARANTIA -->
            <div id="blocoGarantia" style="display:none;">
              <p><strong>📜 Contrato:</strong> <span id="gContrato"></span></p>
              <p><strong>📂 PU SAJ:</strong> <span id="gPuSaj"></span></p>
              <p><strong>💻 Item / Equipamento:</strong> <span id="gItem"></span></p>
              <p><strong>🏢 Fornecedor:</strong> <span id="gFornecedor"></span></p>
              <p><strong>📅 Data do Evento:</strong> <span id="gData"></span></p>
              <p><strong>🛡️ Status Vigência:</strong> <span id="gStatus"></span></p>
              <p id="gNotaContainer"><strong>📄 Nota Fiscal:</strong> <span id="gNota"></span></p>
              <p id="gLinkContainer"><strong>🌐 Suporte:</strong> <a id="gLink" href="#" target="_blank" style="color: #3b82f6; text-decoration: underline;">Abrir Chamado / Suporte ↗</a></p>
            </div>

            <!-- BLOCO CHAMADO -->
            <div id="blocoChamado" style="display:none;">
              <p><strong>🎫 ID do Chamado:</strong> <span id="cId"></span></p>
              <p><strong>🌐 Sistema de Origem:</strong> <span id="cBase"></span></p>
              <p><strong>📋 Status:</strong> <span id="cStatus"></span></p>
              <p><strong>👤 Solicitante:</strong> <span id="cSolicitante"></span></p>
              <p><strong>📍 Localidade / Unidade:</strong> <span id="cLocalidade"></span></p>
              <p><strong>📅 Data Criação:</strong> <span id="cDataCriacao"></span></p>
              <div style="margin-top: 14px; padding-top: 10px; border-top: 1px solid rgba(128,128,128,0.2);">
                <p style="margin-bottom: 6px;"><strong>📝 Resumo da Descrição:</strong></p>
                <div id="cDescricao" style="padding: 10px; background: rgba(128,128,128,0.1); border-radius: 6px; font-size: 0.88rem; line-height: 1.5; max-height: 150px; overflow-y: auto;"></div>
              </div>
            </div>

            <!-- BLOCO PORTARIA -->
            <div id="blocoPortaria" style="display:none;">
              <p><strong>📜 Título:</strong> <span id="poTitulo"></span></p>
              <p><strong>👥 Membros Envolvidos:</strong> <span id="poMembros"></span></p>
              <p><strong>📅 Data de Publicação:</strong> <span id="poDataPub"></span></p>
              <div style="margin-top: 14px; padding-top: 10px; border-top: 1px solid rgba(128,128,128,0.2);">
                <p style="margin-bottom: 6px;"><strong>📝 Ementa / Resumo:</strong></p>
                <div id="poEmenta" style="padding: 10px; background: rgba(128,128,128,0.1); border-radius: 6px; font-size: 0.88rem; line-height: 1.5; max-height: 150px; overflow-y: auto;"></div>
              </div>
              <p id="poPdfContainer" style="margin-top: 14px;">
                <a id="poPdfLink" href="#" target="_blank" style="display: inline-block; background: #3b82f6; color: #ffffff; padding: 6px 12px; border-radius: 6px; text-decoration: none; font-size: 0.88rem; font-weight: 500;">📄 Visualizar Portaria em PDF ↗</a>
              </p>
            </div>

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
          document.getElementById('calendarioModal').style.display = 'none';
        }}

        window.onclick = function(event) {{
          var modal = document.getElementById('calendarioModal');
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
          
          return day + '/' + month + '/' + year + (hours !== '00' || minutes !== '00' ? ' ' + hours + ':' + minutes + ':' + seconds : '');
        }}

        document.addEventListener('DOMContentLoaded', function() {{
          updateThemeFromParent();
          setInterval(updateThemeFromParent, 1000);

          var calendarEl = document.getElementById('calendar');
          var calendar = new FullCalendar.Calendar(calendarEl, {{
            initialView: 'dayGridMonth',
            height: 'auto',
            contentHeight: 'auto',
            dayMaxEvents: 4,
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
              var props = info.event.extendedProps || {{}};
              var cat = props.categoria_evento;

              document.getElementById('mTitle').innerText = info.event.title;
              
              var badge = document.getElementById('mBadge');
              badge.innerText = props.tipo || 'Evento';

              // Oculta todos os blocos internos
              document.getElementById('blocoManual').style.display = 'none';
              document.getElementById('blocoPlantao').style.display = 'none';
              document.getElementById('blocoGarantia').style.display = 'none';
              document.getElementById('blocoChamado').style.display = 'none';
              document.getElementById('blocoPortaria').style.display = 'none';

              switch(cat) {{
                case 'manual':
                  document.getElementById('blocoManual').style.display = 'block';
                  badge.className = 'modal-badge';
                  badge.style.backgroundColor = 'rgba(16, 185, 129, 0.2)';
                  badge.style.color = '#10b981';
                  badge.style.border = '1px solid #10b981';

                  document.getElementById('mManualTitle').innerText = info.event.title;
                  document.getElementById('mManualStart').innerText = formatBrDateTime(info.event.start, props.raw_data_inicio);
                  
                  var mEnd = document.getElementById('mManualEnd');
                  var mEndContainer = document.getElementById('mManualEndContainer');
                  var endFormatted = formatBrDateTime(info.event.end, props.raw_data_fim);
                  if (endFormatted) {{
                    mEnd.innerText = endFormatted;
                    mEndContainer.style.display = 'block';
                  }} else {{
                    mEndContainer.style.display = 'none';
                  }}

                  document.getElementById('mManualAutor').innerText = props.autor || 'Bancada STI';
                  document.getElementById('mManualDesc').innerText = props.descricao || 'Sem descrição informada.';
                  break;

                case 'plantao':
                  document.getElementById('blocoPlantao').style.display = 'block';
                  
                  if (props.tipo && props.tipo.includes('Matutino')) {{
                    badge.className = 'modal-badge';
                    badge.style.backgroundColor = 'rgba(230, 126, 34, 0.2)';
                    badge.style.color = '#e67e22';
                    badge.style.border = '1px solid #e67e22';
                  }} else {{
                    badge.className = 'modal-badge';
                    badge.style.backgroundColor = 'rgba(142, 68, 173, 0.2)';
                    badge.style.color = '#9b59b6';
                    badge.style.border = '1px solid #9b59b6';
                  }}

                  var startFormatted = formatBrDateTime(info.event.start, props.raw_data_inicio);
                  var endFormatted = formatBrDateTime(info.event.end, props.raw_data_fim);
                  
                  document.getElementById('pStart').innerText = startFormatted;
                  
                  var pEnd = document.getElementById('pEnd');
                  var pEndContainer = document.getElementById('pEndContainer');
                  if (endFormatted) {{
                    pEnd.innerText = endFormatted;
                    pEndContainer.style.display = 'block';
                  }} else {{
                    pEndContainer.style.display = 'none';
                  }}
                  
                  var pServ = document.getElementById('pServidor');
                  if (props.detailsHtml) {{
                    pServ.innerHTML = props.detailsHtml;
                  }} else {{
                    pServ.innerText = props.servidor || 'Não informado';
                  }}
                  
                  var pTel = document.getElementById('pTelefone');
                  var pTelContainer = document.getElementById('pTelefoneContainer');
                  if (props.telefone && props.telefone.trim() !== '') {{
                    pTel.innerText = props.telefone;
                    pTelContainer.style.display = 'block';
                  }} else {{
                    pTelContainer.style.display = 'none';
                  }}
                  break;

                case 'garantia':
                  document.getElementById('blocoGarantia').style.display = 'block';
                  
                  if (props.tipo && props.tipo.includes('Início')) {{
                    badge.className = 'modal-badge';
                    badge.style.backgroundColor = 'rgba(16, 185, 129, 0.2)';
                    badge.style.color = '#10b981';
                    badge.style.border = '1px solid #10b981';
                  }} else {{
                    badge.className = 'modal-badge';
                    badge.style.backgroundColor = 'rgba(239, 68, 68, 0.2)';
                    badge.style.color = '#ef4444';
                    badge.style.border = '1px solid #ef4444';
                  }}

                  document.getElementById('gContrato').innerText = props.contrato || 'N/A';
                  document.getElementById('gPuSaj').innerText = props.pu_saj || 'N/A';
                  document.getElementById('gItem').innerText = props.item || 'N/A';
                  document.getElementById('gFornecedor').innerText = props.fornecedor || 'N/A';
                  document.getElementById('gData').innerText = props.data_formatada || 'N/A';
                  document.getElementById('gStatus').innerText = props.status_garantia || 'N/A';

                  var gNota = document.getElementById('gNota');
                  var gNotaContainer = document.getElementById('gNotaContainer');
                  if (props.nota_fiscal) {{
                    gNota.innerText = props.nota_fiscal;
                    gNotaContainer.style.display = 'block';
                  }} else {{
                    gNotaContainer.style.display = 'none';
                  }}

                  var gLink = document.getElementById('gLink');
                  var gLinkContainer = document.getElementById('gLinkContainer');
                  if (props.link_suporte) {{
                    gLink.href = props.link_suporte;
                    gLinkContainer.style.display = 'block';
                  }} else {{
                    gLinkContainer.style.display = 'none';
                  }}
                  break;

                case 'chamado':
                  document.getElementById('blocoChamado').style.display = 'block';
                  
                  badge.className = 'modal-badge';
                  if (props.base === 'OTRS') {{
                    badge.style.backgroundColor = 'rgba(14, 165, 233, 0.2)';
                    badge.style.color = '#0ea5e9';
                    badge.style.border = '1px solid #0ea5e9';
                  }} else {{
                    badge.style.backgroundColor = 'rgba(245, 158, 11, 0.2)';
                    badge.style.color = '#f59e0b';
                    badge.style.border = '1px solid #f59e0b';
                  }}

                  document.getElementById('cId').innerText = props.id || 'N/A';
                  document.getElementById('cBase').innerText = props.base || 'OTRS';
                  document.getElementById('cStatus').innerText = props.status || 'Aberto';
                  document.getElementById('cSolicitante').innerText = props.usuario || 'Não informado';
                  document.getElementById('cLocalidade').innerText = (props.localidade || '') + (props.unidade ? ' - ' + props.unidade : '');
                  document.getElementById('cDataCriacao').innerText = props.data_criacao || 'N/A';
                  document.getElementById('cDescricao').innerText = props.descricao || 'Sem descrição cadastrada.';
                  break;

                case 'portaria':
                  document.getElementById('blocoPortaria').style.display = 'block';

                  badge.className = 'modal-badge';
                  if (props.is_ferias) {{
                    badge.style.backgroundColor = 'rgba(13, 148, 136, 0.2)';
                    badge.style.color = '#0d9488';
                    badge.style.border = '1px solid #0d9488';
                  }} else if (props.is_fiscal) {{
                    badge.style.backgroundColor = 'rgba(139, 92, 246, 0.2)';
                    badge.style.color = '#8b5cf6';
                    badge.style.border = '1px solid #8b5cf6';
                  }} else {{
                    badge.style.backgroundColor = 'rgba(100, 116, 139, 0.2)';
                    badge.style.color = '#94a3b8';
                    badge.style.border = '1px solid #94a3b8';
                  }}

                  document.getElementById('poTitulo').innerText = props.titulo || info.event.title;
                  document.getElementById('poMembros').innerText = props.membros || 'Não informado';
                  document.getElementById('poDataPub').innerText = props.data_publicacao || props.data_emissao || 'N/A';
                  document.getElementById('poEmenta').innerText = props.ementa || 'Sem ementa disponível.';

                  var poPdfLink = document.getElementById('poPdfLink');
                  var poPdfContainer = document.getElementById('poPdfContainer');
                  if (props.pdf_url) {{
                    poPdfLink.href = props.pdf_url;
                    poPdfContainer.style.display = 'block';
                  }} else {{
                    poPdfContainer.style.display = 'none';
                  }}
                  break;

                default:
                  break;
              }}

              document.getElementById('calendarioModal').style.display = 'flex';
            }}
          }});
          calendar.render();
        }});
      </script>
    </body>
    </html>
    """
    components.html(calendar_html, height=1150, scrolling=True)


def render_calendario_geral_page():
    """Renderiza a página principal do Calendário Geral Unificado com Filtros Laterais e Botão de Novo Evento."""
    st.title("📅 Calendário Geral Unificado")
    st.caption("Visão centralizada de registros manuais, plantões da bancada, vigências de contratos de garantia, portarias e chamados técnicos.")

    # --- BOTÃO DE DESTAQUE NO TOPO DA SIDEBAR ---
    if st.sidebar.button("➕ Novo Evento Manual", type="primary", use_container_width=True):
        modal_novo_evento_manual()

    st.sidebar.markdown("---")

    # --- FILTROS SIDEBAR (Categorias Desagrupadas) ---
    st.sidebar.markdown("## 📅 Agendas Visíveis")
    
    chk_manuais = st.sidebar.checkbox("Registros Manuais", value=True)
    chk_matutino = st.sidebar.checkbox("Plantão Matutino", value=True)
    chk_semanal = st.sidebar.checkbox("Plantão Semanal", value=True)
    chk_garantias = st.sidebar.checkbox("Garantias", value=True)
    chk_otrs = st.sidebar.checkbox("Chamados OTRS", value=True)
    chk_citsmart = st.sidebar.checkbox("Chamados CitSmart", value=True)
    chk_portarias = st.sidebar.checkbox("Portarias (Geral)", value=True)
    chk_ferias = st.sidebar.checkbox("Portarias (Férias)", value=True)
    chk_portarias_fiscais = st.sidebar.checkbox("Portarias (Fiscais de Contrato)", value=True)

    st.sidebar.markdown("---")
    st.sidebar.markdown("## 🔍 Opções Adicionais")

    search_query = st.sidebar.text_input("🔎 Pesquisar Evento:", placeholder="Ex: Paulo, impressora, #4645...").strip().lower()

    anos_disponiveis = [2026, 2025, 2024]
    selected_ano = st.sidebar.selectbox("📅 Selecionar Ano (Plantões):", anos_disponiveis)

    # --- CONSTRUÇÃO DA LISTA UNIFICADA DE EVENTOS ---
    events = []

    # 1. EVENTOS MANUAIS
    if chk_manuais:
        df_manuais = get_eventos_manuais()
        if not df_manuais.empty:
            for idx, row in df_manuais.iterrows():
                titulo = str(row.get('titulo', '')).strip()
                dt_ini = str(row.get('data_inicio', '')).strip()
                dt_fim = str(row.get('data_fim', '')).strip()
                autor = str(row.get('autor', 'Bancada STI')).strip()
                descricao = str(row.get('descricao', '')).strip()

                if not dt_ini:
                    continue

                events.append({
                    "id": f"manual_{row.get('id', idx)}",
                    "title": f"📝 {titulo}",
                    "start": dt_ini,
                    "end": dt_fim if dt_fim else dt_ini,
                    "backgroundColor": "#10b981",
                    "borderColor": "#059669",
                    "extendedProps": {
                        "categoria_evento": "manual",
                        "tipo": "Registro Manual",
                        "autor": autor,
                        "descricao": descricao,
                        "raw_data_inicio": dt_ini,
                        "raw_data_fim": dt_fim
                    }
                })

    # 2. PLANTÃO MATUTINO
    if chk_matutino:
        df_matutino = get_plantoes_matutino(selected_ano)
        if not df_matutino.empty:
            for _, row in df_matutino.iterrows():
                servidor = str(row.get('servidor', '')).strip()
                dt_iso = str(row.get('data_iso', '')).strip()
                if not dt_iso:
                    continue
                tel_formatted = format_phone_number(str(row.get('telefone', '')))
                
                dt_ini_str = f"{dt_iso}T08:00:00"
                dt_fim_str = f"{dt_iso}T15:00:00"
                
                raw_ini_br, raw_fim_br = "", ""
                try:
                    d_obj = datetime.strptime(dt_iso, "%Y-%m-%d")
                    raw_ini_br = d_obj.strftime("%d/%m/%Y 08:00:00")
                    raw_fim_br = d_obj.strftime("%d/%m/%Y 15:00:00")
                except Exception:
                    pass

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

    # 3. PLANTÃO SEMANAL
    if chk_semanal:
        df_semanal = get_plantoes_semanal(selected_ano)
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
                
                raw_ini_br, raw_fim_br = "", ""
                try:
                    if dt_ini_raw:
                        dt_o = datetime.strptime(dt_ini_raw, "%Y-%m-%d %H:%M:%S")
                        raw_ini_br = dt_o.strftime("%d/%m/%Y %H:%M:%S")
                    if dt_fim_raw:
                        dt_f = datetime.strptime(dt_fim_raw, "%Y-%m-%d %H:%M:%S")
                        raw_fim_br = dt_f.strftime("%d/%m/%Y %H:%M:%S")
                except Exception:
                    pass

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

    # 4. CONTRATOS DE GARANTIA
    if chk_garantias:
        df_garantia = get_garantia_contratos_df()
        if not df_garantia.empty:
            for idx, row in df_garantia.iterrows():
                contrato = str(row.get('contrato', '')).strip()
                pu_saj = str(row.get('pu_saj', '')).strip()
                item = str(row.get('item', '')).strip()
                fornecedor = str(row.get('fornecedor', '')).strip()
                nota_fiscal = str(row.get('nota_fiscal', '')).strip()
                status_garantia = str(row.get('status_garantia', '')).strip()
                link_suporte = str(row.get('link_suporte', '')).strip()

                iso_ini, br_ini = parse_date_to_iso_and_br(row.get('garantia_inicio'))
                if iso_ini:
                    events.append({
                        "id": f"garantia_ini_{idx}",
                        "title": f"🟢 Início Garantia: {item} ({fornecedor})",
                        "start": iso_ini,
                        "backgroundColor": "#3b82f6",
                        "borderColor": "#1d4ed8",
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

                iso_fim, br_fim = parse_date_to_iso_and_br(row.get('garantia_fim'))
                if iso_fim:
                    bg_col = "#ef4444" if status_garantia == "Garantia Vencida" else ("#f59e0b" if "30" in status_garantia else "#3b82f6")
                    events.append({
                        "id": f"garantia_fim_{idx}",
                        "title": f"🔴 Fim Garantia: {item} ({fornecedor})",
                        "start": iso_fim,
                        "backgroundColor": bg_col,
                        "borderColor": bg_col,
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

    # 5. CHAMADOS (OTRS e CITSMART)
    if chk_otrs or chk_citsmart:
        df_chamados = load_data()
        if not df_chamados.empty:
            for _, row in df_chamados.iterrows():
                base = str(row.get('base', 'OTRS')).strip()
                
                # Filtra estritamente de acordo com a base e o checkbox selecionado
                if base == "OTRS" and not chk_otrs:
                    continue
                if base == "CitSmart" and not chk_citsmart:
                    continue
                if base not in ["OTRS", "CitSmart"] and not (chk_otrs and chk_citsmart):
                    continue

                cid = str(row.get('id', '')).strip()
                titulo = str(row.get('titulo', '')).strip()
                status = str(row.get('status', 'Aberto')).strip()
                usuario = str(row.get('usuario', '')).strip()
                localidade = str(row.get('localidade_fisica', '')).strip()
                unidade = str(row.get('unidade', '')).strip()
                descricao = str(row.get('descricao', '')).strip()
                dt_criacao_raw = row.get('data_criacao')

                iso_dt, br_dt = parse_ticket_date_iso_and_br(dt_criacao_raw)
                if not iso_dt:
                    continue

                bg_col = "#0ea5e9" if base == "OTRS" else "#f59e0b"
                border_col = "#0284c7" if base == "OTRS" else "#d97706"

                desc_resumo = (descricao[:250] + "...") if len(descricao) > 250 else descricao

                events.append({
                    "id": f"chamado_{cid}",
                    "title": f"📋 #{cid} ({base}): {titulo[:35]}",
                    "start": iso_dt,
                    "backgroundColor": bg_col,
                    "borderColor": border_col,
                    "extendedProps": {
                        "categoria_evento": "chamado",
                        "tipo": f"Chamado {base}",
                        "id": cid,
                        "base": base,
                        "titulo": titulo,
                        "status": status,
                        "usuario": usuario,
                        "localidade": localidade,
                        "unidade": unidade,
                        "data_criacao": br_dt if br_dt else str(dt_criacao_raw),
                        "descricao": desc_resumo if desc_resumo else "Sem descrição."
                    }
                })

    # 6. PORTARIAS DA BANCADA (3 NÍVEIS: FÉRIAS, FISCAIS E GERAL)
    if chk_portarias or chk_ferias or chk_portarias_fiscais:
        try:
            portarias_list = fetch_portarias_bancada()
            for p_item in portarias_list:
                p_id = p_item.get("id")
                titulo = p_item.get("titulo", "")
                texto = p_item.get("texto", "")
                dt_emissao = p_item.get("data_emissao", "")
                dt_pub = p_item.get("data_publicacao", "")
                membros = ", ".join(p_item.get("membros", []))
                pdf_url = p_item.get("pdf_url", "")

                full_text = f"{titulo} {texto}".lower()

                # CLASSIFICAÇÃO EM 3 NÍVEIS
                is_ferias = "férias" in full_text or "ferias" in full_text
                
                is_fiscal = False
                if not is_ferias:
                    keywords_fiscal = ["fiscalização", "fiscalizacao", "fiscal", "fiscais", "nota de empenho", "gestão", "gestao", "contrato"]
                    is_fiscal = any(kw in full_text for kw in keywords_fiscal)

                # FILTRAGEM POR CHECKBOX
                if is_ferias and not chk_ferias:
                    continue
                if is_fiscal and not chk_portarias_fiscais:
                    continue
                if not is_ferias and not is_fiscal and not chk_portarias:
                    continue

                # Converte data de emissão para ISO
                iso_date_emissao, br_date_emissao = None, None
                if dt_emissao:
                    try:
                        dt_obj = datetime.strptime(dt_emissao, "%d/%m/%Y")
                        iso_date_emissao = dt_obj.strftime("%Y-%m-%d")
                        br_date_emissao = dt_emissao
                    except Exception:
                        pass

                if not iso_date_emissao:
                    continue

                start_dt, end_dt = iso_date_emissao, iso_date_emissao
                start_br, end_br = br_date_emissao, br_date_emissao

                # Se for Férias, tenta extrair o período via Regex (com fallback)
                if is_ferias:
                    r_start_iso, r_end_iso, r_start_br, r_end_br = extract_vacation_dates(full_text, br_date_emissao)
                    if r_start_iso and r_end_iso:
                        start_dt, end_dt = r_start_iso, r_end_iso
                        start_br, end_br = r_start_br, r_end_br

                # CORES E LABELS POR NÍVEL
                if is_ferias:
                    bg_col, border_col = "#0d9488", "#0f766e"
                    event_type_label = "🏖️ Portaria de Férias"
                    icon = "🏖️"
                elif is_fiscal:
                    bg_col, border_col = "#8b5cf6", "#7c3aed"
                    event_type_label = "📜 Fiscalização de Contrato"
                    icon = "📜"
                else:
                    bg_col, border_col = "#475569", "#334155"
                    event_type_label = "📜 Portaria Geral"
                    icon = "📜"

                ementa_snippet = (texto[:300] + "...") if len(texto) > 300 else (texto if texto else titulo)

                events.append({
                    "id": f"portaria_{p_id}",
                    "title": f"{icon} Portaria #{p_item.get('numero', p_id)}: {membros.split(',')[0]}",
                    "start": start_dt,
                    "end": end_dt,
                    "backgroundColor": bg_col,
                    "borderColor": border_col,
                    "extendedProps": {
                        "categoria_evento": "portaria",
                        "tipo": event_type_label,
                        "is_ferias": is_ferias,
                        "is_fiscal": is_fiscal,
                        "titulo": titulo,
                        "membros": membros,
                        "data_emissao": start_br,
                        "data_publicacao": dt_pub if dt_pub else start_br,
                        "ementa": ementa_snippet,
                        "pdf_url": pdf_url,
                        "raw_data_inicio": start_br,
                        "raw_data_fim": end_br
                    }
                })
        except Exception as e:
            pass

    # --- APLICA PESQUISA GLOBAL (FILTRO DE EVENTOS) ---
    filtered_events = events
    if search_query:
        matching_events = []
        for ev in events:
            # Varre o título e todas as extendedProps em busca da palavra-chave
            ev_title = str(ev.get("title", "")).lower()
            props = ev.get("extendedProps", {})
            props_str = " ".join([str(v).lower() for v in props.values() if v])
            
            combined_search_text = f"{ev_title} {props_str}"
            if search_query in combined_search_text:
                matching_events.append(ev)
        filtered_events = matching_events

    render_master_calendar(filtered_events)

    # --- TABELA DE RESULTADOS DA PESQUISA (EXIBIDA SE HOUVER BUSCA) ---
    if search_query:
        st.markdown("<br>", unsafe_allow_html=True)
        st.subheader(f"🔎 Resultados da Pesquisa para '{search_query}' ({len(filtered_events)} eventos encontrados)")
        
        if not filtered_events:
            st.info("Nenhum evento encontrado para a palavra-chave informada.")
        else:
            table_rows = []
            for ev in filtered_events:
                props = ev.get("extendedProps", {})
                
                # Monta detalhes amigáveis por categoria
                cat = props.get("categoria_evento", "")
                detalhes = ""
                if cat == "chamado":
                    detalhes = f"Solicitante: {props.get('usuario', 'N/A')} | Local: {props.get('localidade', 'N/A')}"
                elif cat == "plantao":
                    detalhes = props.get("servidor", "")
                elif cat == "garantia":
                    detalhes = f"Contrato: {props.get('contrato', 'N/A')} | Fornecedor: {props.get('fornecedor', 'N/A')}"
                elif cat == "portaria":
                    detalhes = f"Membros: {props.get('membros', 'N/A')} | Ementa: {props.get('ementa', '')[:100]}..."
                elif cat == "manual":
                    detalhes = f"Autor: {props.get('autor', 'N/A')} | {props.get('descricao', '')}"

                dt_exibicao = props.get("raw_data_inicio") or str(ev.get("start", ""))

                table_rows.append({
                    "Título": ev.get("title", ""),
                    "Categoria / Tipo": props.get("tipo", "Evento"),
                    "Data / Período": dt_exibicao,
                    "Detalhes / Resumo": detalhes
                })

            df_search_res = pd.DataFrame(table_rows)
            st.dataframe(
                df_search_res,
                column_config={
                    "Título": st.column_config.TextColumn("Título do Evento"),
                    "Categoria / Tipo": st.column_config.TextColumn("Categoria / Tipo"),
                    "Data / Período": st.column_config.TextColumn("Data / Período"),
                    "Detalhes / Resumo": st.column_config.TextColumn("Detalhes / Resumo"),
                },
                hide_index=True,
                use_container_width=True
            )
