import json
import streamlit as st
import streamlit.components.v1 as components


def render_master_calendar(events: list[dict], height_px=860, scrolling_enabled=True):
    """
    Renderiza o componente interativo FullCalendar (v6) replicando o tema Dark/Light (Glassmorphism),
    responsividade e o Modal Dinâmico Inteligente com blocos independentes para Registros Manuais, Plantões, Garantia, Chamados e Portarias.
    """
    events_json = json.dumps(events, ensure_ascii=True).replace("<", "\\u003c").replace(">", "\\u003e")

    html_template = """
    <!DOCTYPE html>
    <html lang="pt-br">
    <head>
      <meta charset="utf-8"/>
      <script src="https://cdn.jsdelivr.net/npm/fullcalendar@6.1.10/index.global.min.js"></script>
      <style>
        html, body {
          margin: 0;
          padding: 0;
          width: 100%;
          max-width: 100%;
          overflow-x: hidden !important;
          overflow-y: auto;
          font-family: 'Segoe UI', Roboto, Helvetica, Arial, sans-serif;
          transition: background-color 0.3s ease, color 0.3s ease;
        }

        #calendar {
          width: 100% !important;
          max-width: 100% !important;
          box-sizing: border-box;
          padding: 5px;
        }
        .fc {
          max-width: 100% !important;
          overflow-x: hidden !important;
        }
        .fc-view-harness {
          max-width: 100% !important;
        }
        .fc-scroller {
          overflow-x: hidden !important;
          overflow-y: auto !important;
        }
        .fc-scrollgrid {
          width: 100% !important;
        }

        /* Scrollbars Customizados no Iframe do Calendário */
        ::-webkit-scrollbar {
          width: 8px !important;
          height: 8px !important;
        }
        ::-webkit-scrollbar-track {
          background: rgba(15, 23, 42, 0.4) !important;
          border-radius: 4px !important;
        }
        ::-webkit-scrollbar-thumb {
          background: #3b82f6 !important;
          border-radius: 4px !important;
        }
        ::-webkit-scrollbar-thumb:hover {
          background: #2563eb !important;
        }

        /* TEMA ESCURO (DARK MODE) */
        body.dark-mode {
          background-color: #0e1117;
          color: #fafafa;
        }
        body.dark-mode .fc-theme-standard td, 
        body.dark-mode .fc-theme-standard th, 
        body.dark-mode .fc-theme-standard .fc-scrollgrid {
          border-color: #262730 !important;
        }
        body.dark-mode .fc-col-header-cell {
          background-color: #161b22;
          color: #c9d1d9;
        }
        body.dark-mode .fc-daygrid-day-number, 
        body.dark-mode .fc-list-day-text, 
        body.dark-mode .fc-list-day-side-text {
          color: #e6edf3 !important;
          text-decoration: none !important;
        }
        body.dark-mode .fc-button-primary {
          background-color: #21262d !important;
          border-color: #30363d !important;
          color: #c9d1d9 !important;
        }
        body.dark-mode .fc-button-primary:hover {
          background-color: #30363d !important;
        }
        body.dark-mode .fc-button-primary:disabled {
          background-color: #161b22 !important;
          border-color: #21262d !important;
        }
        body.dark-mode .fc-button-active {
          background-color: #ff4b4b !important;
          border-color: #ff4b4b !important;
          color: #ffffff !important;
        }
        body.dark-mode .fc-day-today {
          background-color: rgba(255, 75, 75, 0.08) !important;
        }
        body.dark-mode .fc-list-day-cushion {
          background-color: #161b22 !important;
        }
        body.dark-mode .fc-list-event:hover td {
          background-color: #21262d !important;
        }
        body.dark-mode .fc-list-day.fc-day-today .fc-list-day-cushion {
          background-color: rgba(230, 126, 34, 0.25) !important;
          color: #e67e22 !important;
        }
        body.dark-mode .modal-content {
          background: #161b22;
          border: 1px solid #30363d;
          color: #e6edf3;
        }
        body.dark-mode .modal-body p { color: #8b949e; }
        body.dark-mode .modal-body strong { color: #f0f6fc; }
        body.dark-mode .close-btn { color: #8b949e; }

        /* POPOVER "+X MAIS" (FULLCALENDAR MORE-POPOVER) - DARK MODE */
        body.dark-mode .fc {
          --fc-page-bg-color: #161b22;
          --fc-neutral-bg-color: #21262d;
          --fc-border-color: #262730;
        }
        body.dark-mode .fc-popover {
          background-color: #161b22 !important;
          border: 1px solid #30363d !important;
          box-shadow: 0 10px 25px rgba(0, 0, 0, 0.8) !important;
          border-radius: 10px !important;
          color: #e6edf3 !important;
        }
        body.dark-mode .fc-popover-header {
          background-color: #21262d !important;
          color: #f0f6fc !important;
          border-bottom: 1px solid #30363d !important;
          border-top-left-radius: 9px !important;
          border-top-right-radius: 9px !important;
          padding: 8px 12px !important;
        }
        body.dark-mode .fc-popover-title {
          color: #f0f6fc !important;
          font-weight: 600 !important;
          font-size: 0.95rem !important;
        }
        body.dark-mode .fc-popover-close {
          color: #8b949e !important;
          opacity: 0.8 !important;
          cursor: pointer !important;
        }
        body.dark-mode .fc-popover-close:hover {
          color: #f0f6fc !important;
          opacity: 1 !important;
        }
        body.dark-mode .fc-popover-body {
          background-color: #161b22 !important;
          border-bottom-left-radius: 9px !important;
          border-bottom-right-radius: 9px !important;
          padding: 8px !important;
        }
        body.dark-mode .fc-popover .fc-daygrid-event {
          background-color: #21262d;
          border: 1px solid #30363d;
          color: #e6edf3 !important;
          margin-bottom: 4px !important;
          border-radius: 6px !important;
          padding: 4px 6px !important;
        }
        body.dark-mode .fc-popover .fc-daygrid-event:hover {
          filter: brightness(1.15);
        }
        body.dark-mode .fc-popover .fc-event-title,
        body.dark-mode .fc-popover .fc-event-time,
        body.dark-mode .fc-popover .fc-event-main,
        body.dark-mode .fc-popover .fc-daygrid-event-dot {
          color: #e6edf3 !important;
        }

        /* TEMA CLARO (LIGHT MODE) */
        body.light-mode {
          background-color: #ffffff;
          color: #0f172a;
        }
        body.light-mode .fc-theme-standard td, 
        body.light-mode .fc-theme-standard th, 
        body.light-mode .fc-theme-standard .fc-scrollgrid {
          border-color: #e2e8f0 !important;
        }
        body.light-mode .fc-col-header-cell {
          background-color: #f8fafc;
          color: #334155;
        }
        body.light-mode .fc-daygrid-day-number, 
        body.light-mode .fc-list-day-text, 
        body.light-mode .fc-list-day-side-text {
          color: #1e293b !important;
          text-decoration: none !important;
        }
        body.light-mode .fc-button-primary {
          background-color: #f1f5f9 !important;
          border-color: #cbd5e1 !important;
          color: #334155 !important;
        }
        body.light-mode .fc-button-primary:hover {
          background-color: #e2e8f0 !important;
        }
        body.light-mode .fc-button-primary:disabled {
          background-color: #f8fafc !important;
        }
        body.light-mode .fc-button-active {
          background-color: #ff4b4b !important;
          border-color: #ff4b4b !important;
          color: #ffffff !important;
        }
        body.light-mode .fc-day-today {
          background-color: rgba(255, 75, 75, 0.05) !important;
        }
        body.light-mode .fc-list-day-cushion {
          background-color: #f1f5f9 !important;
          color: #0f172a !important;
        }
        body.light-mode .fc-list-day.fc-day-today .fc-list-day-cushion {
          background-color: rgba(230, 126, 34, 0.20) !important;
          color: #d35400 !important;
        }
        body.light-mode .modal-content {
          background: #ffffff;
          border: 1px solid #cbd5e1;
          color: #0f172a;
        }
        body.light-mode .modal-body p { color: #334155; }
        body.light-mode .modal-body strong { color: #0f172a; }
        body.light-mode .close-btn { color: #64748b; }

        /* POPOVER "+X MAIS" (FULLCALENDAR MORE-POPOVER) - LIGHT MODE */
        body.light-mode .fc {
          --fc-page-bg-color: #ffffff;
          --fc-neutral-bg-color: #f8fafc;
          --fc-border-color: #e2e8f0;
        }
        body.light-mode .fc-popover {
          background-color: #ffffff !important;
          border: 1px solid #cbd5e1 !important;
          box-shadow: 0 10px 25px rgba(0, 0, 0, 0.15) !important;
          border-radius: 10px !important;
          color: #0f172a !important;
        }
        body.light-mode .fc-popover-header {
          background-color: #f8fafc !important;
          color: #0f172a !important;
          border-bottom: 1px solid #e2e8f0 !important;
          border-top-left-radius: 9px !important;
          border-top-right-radius: 9px !important;
          padding: 8px 12px !important;
        }
        body.light-mode .fc-popover-title {
          color: #0f172a !important;
          font-weight: 600 !important;
          font-size: 0.95rem !important;
        }
        body.light-mode .fc-popover-close {
          color: #64748b !important;
          opacity: 0.8 !important;
          cursor: pointer !important;
        }
        body.light-mode .fc-popover-close:hover {
          color: #0f172a !important;
          opacity: 1 !important;
        }
        body.light-mode .fc-popover-body {
          background-color: #ffffff !important;
          border-bottom-left-radius: 9px !important;
          border-bottom-right-radius: 9px !important;
          padding: 8px !important;
        }
        body.light-mode .fc-popover .fc-daygrid-event {
          background-color: #f8fafc;
          border: 1px solid #e2e8f0;
          color: #0f172a !important;
          margin-bottom: 4px !important;
          border-radius: 6px !important;
          padding: 4px 6px !important;
        }
        body.light-mode .fc-popover .fc-daygrid-event:hover {
          background-color: #f1f5f9;
        }
        body.light-mode .fc-popover .fc-event-title,
        body.light-mode .fc-popover .fc-event-time,
        body.light-mode .fc-popover .fc-event-main {
          color: inherit !important;
        }

        .fc-event {
          cursor: pointer;
          border-radius: 4px;
          padding: 2px 4px;
          font-size: 0.85rem;
          box-shadow: 0 2px 4px rgba(0,0,0,0.2);
          transition: transform 0.1s ease;
        }
        .fc-event:hover {
          transform: scale(1.02);
        }

        /* GLASSMORPHISM MODAL OVERLAY */
        .modal-overlay {
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
        }
        .modal-content {
          border-radius: 12px;
          width: 90%;
          max-width: 540px;
          padding: 24px;
          box-shadow: 0 10px 30px rgba(0,0,0,0.5);
          animation: fadeIn 0.2s ease-out;
        }
        @keyframes fadeIn {
          from { opacity: 0; transform: translateY(-10px); }
          to { opacity: 1; transform: translateY(0); }
        }
        .modal-header {
          display: flex;
          justify-content: space-between;
          align-items: center;
          border-bottom: 1px solid rgba(128,128,128,0.3);
          padding-bottom: 12px;
          margin-bottom: 16px;
        }
        .modal-title {
          margin: 0;
          font-size: 1.15rem;
          font-weight: 600;
          color: #ff4b4b;
        }
        .close-btn {
          background: transparent;
          border: none;
          font-size: 1.5rem;
          cursor: pointer;
          line-height: 1;
        }
        .modal-body p {
          margin: 10px 0;
          font-size: 0.95rem;
          line-height: 1.5;
        }
        .modal-badge {
          display: inline-block;
          padding: 4px 10px;
          border-radius: 12px;
          font-size: 0.75rem;
          font-weight: 600;
          margin-bottom: 12px;
        }
        .modal-footer {
          margin-top: 20px;
          text-align: right;
        }
        .btn-dismiss {
          background: #ff4b4b;
          color: #fff;
          border: none;
          padding: 8px 16px;
          border-radius: 6px;
          cursor: pointer;
          font-weight: 500;
        }
        .btn-dismiss:hover {
          background: #e03e3e;
        }
      </style>
    </head>
    <body>
      <script id="eventsData" type="application/json">__EVENTS_JSON__</script>
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
              <p><strong>📌 Título do Chamado:</strong> <span id="cTitulo"></span></p>
              <p><strong>🏷️ TAG:</strong> <span id="cTag" style="padding: 2px 8px; border-radius: 6px; font-weight: 600; font-size: 0.82rem;"></span></p>
              <p><strong>🎫 ID do Chamado:</strong> <span id="cId"></span></p>
              <p><strong>🌐 Sistema de Origem:</strong> <span id="cBase"></span></p>
              <p><strong>📋 Status:</strong> <span id="cStatus"></span></p>
              <p><strong>👤 Solicitante:</strong> <span id="cSolicitante"></span></p>
              <p><strong>📍 Localidade / Unidade:</strong> <span id="cLocalidade"></span></p>
              <p><strong>📅 Data Criação:</strong> <span id="cDataCriacao"></span></p>
              <div style="margin-top: 14px; padding-top: 10px; border-top: 1px solid rgba(128,128,128,0.2);">
                <p style="margin-bottom: 6px;"><strong>📝 Resumo da Descrição:</strong></p>
                <div id="cDescricao" style="padding: 10px; background: rgba(128,128,128,0.1); border-radius: 6px; font-size: 0.88rem; line-height: 1.5; max-height: 140px; overflow-y: auto;"></div>
              </div>
              <div id="cComentariosContainer" style="margin-top: 14px; padding-top: 10px; border-top: 1px solid rgba(128,128,128,0.2);">
                <p style="margin-bottom: 8px;"><strong>💬 Histórico de Notas / Acompanhamentos:</strong></p>
                <div id="cComentariosList" style="max-height: 200px; overflow-y: auto; display: flex; flex-direction: column; gap: 8px;"></div>
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
        function updateThemeFromParent() {
          var isLight = false;
          try {
            var parentBody = window.parent.document.body;
            var parentApp = window.parent.document.querySelector('.stApp');
            var themeAttr = (parentBody && parentBody.getAttribute('data-theme')) || 
                            (parentApp && parentApp.getAttribute('data-theme'));
            
            if (themeAttr === 'light') {
              isLight = true;
            } else if (themeAttr === 'dark') {
              isLight = false;
            } else {
              isLight = window.parent.matchMedia('(prefers-color-scheme: light)').matches;
            }
          } catch(e) {
            isLight = window.matchMedia('(prefers-color-scheme: light)').matches;
          }

          if (isLight) {
            document.body.className = 'light-mode';
          } else {
            document.body.className = 'dark-mode';
          }
        }

        function closeModal() {
          document.getElementById('calendarioModal').style.display = 'none';
        }

        window.onclick = function(event) {
          var modal = document.getElementById('calendarioModal');
          if (event.target == modal) {
            modal.style.display = 'none';
          }
        };

        function formatBrDateTime(dateObj, rawStr) {
          if (rawStr && typeof rawStr === 'string' && rawStr.includes('/')) {
            return rawStr;
          }
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
        }

        function initCalendar() {
          updateThemeFromParent();
          setInterval(updateThemeFromParent, 1000);

          var calendarEl = document.getElementById('calendar');
          if (!calendarEl || !window.FullCalendar) {
            setTimeout(initCalendar, 100);
            return;
          }

          var calendar = new FullCalendar.Calendar(calendarEl, {
            initialView: 'dayGridMonth',
            height: 'auto',
            contentHeight: 'auto',
            dayMaxEvents: 4,
            locale: 'pt-br',

            headerToolbar: {
              left: 'prev,next today',
              center: 'title',
              right: 'dayGridMonth,timeGridWeek,timeGridDay,listMonth'
            },
            buttonText: {
              today:    'Hoje',
              month:    'Mês',
              week:     'Semana',
              day:      'Dia',
              list:     'Lista'
            },
            events: JSON.parse(document.getElementById('eventsData').textContent),
            eventClick: function(info) {
              var props = info.event.extendedProps || {};
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

              switch(cat) {
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
                  if (endFormatted) {
                    mEnd.innerText = endFormatted;
                    mEndContainer.style.display = 'block';
                  } else {
                    mEndContainer.style.display = 'none';
                  }

                  document.getElementById('mManualAutor').innerText = props.autor || 'Bancada STI';
                  document.getElementById('mManualDesc').innerText = props.descricao || 'Sem descrição informada.';
                  break;

                case 'plantao':
                  document.getElementById('blocoPlantao').style.display = 'block';
                  
                  if (props.tipo && props.tipo.includes('Matutino')) {
                    badge.className = 'modal-badge';
                    badge.style.backgroundColor = 'rgba(230, 126, 34, 0.2)';
                    badge.style.color = '#e67e22';
                    badge.style.border = '1px solid #e67e22';
                  } else {
                    badge.className = 'modal-badge';
                    badge.style.backgroundColor = 'rgba(142, 68, 173, 0.2)';
                    badge.style.color = '#9b59b6';
                    badge.style.border = '1px solid #9b59b6';
                  }

                  var startFormatted = formatBrDateTime(info.event.start, props.raw_data_inicio);
                  var endFormatted = formatBrDateTime(info.event.end, props.raw_data_fim);
                  
                  document.getElementById('pStart').innerText = startFormatted;
                  
                  var pEnd = document.getElementById('pEnd');
                  var pEndContainer = document.getElementById('pEndContainer');
                  if (endFormatted) {
                    pEnd.innerText = endFormatted;
                    pEndContainer.style.display = 'block';
                  } else {
                    pEndContainer.style.display = 'none';
                  }
                  
                  var pServ = document.getElementById('pServidor');
                  if (props.detailsHtml) {
                    pServ.innerHTML = props.detailsHtml;
                  } else {
                    pServ.innerText = props.servidor || 'Não informado';
                  }
                  
                  var pTel = document.getElementById('pTelefone');
                  var pTelContainer = document.getElementById('pTelefoneContainer');
                  if (props.telefone && props.telefone.trim() !== '') {
                    pTel.innerText = props.telefone;
                    pTelContainer.style.display = 'block';
                  } else {
                    pTelContainer.style.display = 'none';
                  }
                  break;

                case 'garantia':
                  document.getElementById('blocoGarantia').style.display = 'block';
                  
                  if (props.tipo && props.tipo.includes('Início')) {
                    badge.className = 'modal-badge';
                    badge.style.backgroundColor = 'rgba(16, 185, 129, 0.2)';
                    badge.style.color = '#10b981';
                    badge.style.border = '1px solid #10b981';
                  } else {
                    badge.className = 'modal-badge';
                    badge.style.backgroundColor = 'rgba(239, 68, 68, 0.2)';
                    badge.style.color = '#ef4444';
                    badge.style.border = '1px solid #ef4444';
                  }

                  document.getElementById('gContrato').innerText = props.contrato || 'N/A';
                  document.getElementById('gPuSaj').innerText = props.pu_saj || 'N/A';
                  document.getElementById('gItem').innerText = props.item || 'N/A';
                  document.getElementById('gFornecedor').innerText = props.fornecedor || 'N/A';
                  document.getElementById('gData').innerText = props.data_formatada || 'N/A';
                  document.getElementById('gStatus').innerText = props.status_garantia || 'N/A';

                  var gNota = document.getElementById('gNota');
                  var gNotaContainer = document.getElementById('gNotaContainer');
                  if (props.nota_fiscal) {
                    gNota.innerText = props.nota_fiscal;
                    gNotaContainer.style.display = 'block';
                  } else {
                    gNotaContainer.style.display = 'none';
                  }

                  var gLink = document.getElementById('gLink');
                  var gLinkContainer = document.getElementById('gLinkContainer');
                  if (props.link_suporte) {
                    gLink.href = props.link_suporte;
                    gLinkContainer.style.display = 'block';
                  } else {
                    gLinkContainer.style.display = 'none';
                  }
                  break;

                case 'chamado':
                  document.getElementById('blocoChamado').style.display = 'block';
                  
                  badge.className = 'modal-badge';
                  if (props.base === 'OTRS') {
                    badge.style.backgroundColor = 'rgba(14, 165, 233, 0.2)';
                    badge.style.color = '#0ea5e9';
                    badge.style.border = '1px solid #0ea5e9';
                  } else {
                    badge.style.backgroundColor = 'rgba(245, 158, 11, 0.2)';
                    badge.style.color = '#f59e0b';
                    badge.style.border = '1px solid #f59e0b';
                  }

                  document.getElementById('mTitle').innerText = '📋 Chamado #' + (props.id || '') + ' (' + (props.base || 'OTRS') + ')';
                  document.getElementById('cTitulo').innerText = props.titulo_completo || info.event.title || 'Sem Título';

                  var tagEl = document.getElementById('cTag');
                  var tagVal = props.tag || 'Sem TAG';
                  tagEl.innerText = tagVal;
                  tagEl.style.background = 'rgba(128,128,128,0.2)';
                  tagEl.style.color = '#38bdf8';
                  tagEl.style.border = '1px solid rgba(56, 189, 248, 0.4)';

                  document.getElementById('cId').innerText = props.id || 'N/A';
                  document.getElementById('cBase').innerText = props.base || 'OTRS';
                  document.getElementById('cStatus').innerText = props.status || 'Aberto';
                  document.getElementById('cSolicitante').innerText = props.usuario || 'Não informado';
                  document.getElementById('cLocalidade').innerText = (props.localidade || '') + (props.unidade ? ' - ' + props.unidade : '');
                  document.getElementById('cDataCriacao').innerText = props.data_criacao || 'N/A';
                  document.getElementById('cDescricao').innerText = props.descricao || 'Sem descrição cadastrada.';

                  var commentsContainer = document.getElementById('cComentariosContainer');
                  var commentsListEl = document.getElementById('cComentariosList');
                  commentsListEl.innerHTML = '';

                  if (props.comentarios && props.comentarios.length > 0) {
                    commentsContainer.style.display = 'block';
                    props.comentarios.forEach(function(c, idx) {
                      var details = document.createElement('details');
                      details.style.background = 'rgba(128,128,128,0.1)';
                      details.style.borderRadius = '6px';
                      details.style.padding = '8px 12px';
                      details.style.border = '1px solid rgba(128,128,128,0.2)';
                      details.style.fontSize = '0.85rem';

                      var summary = document.createElement('summary');
                      summary.style.cursor = 'pointer';
                      summary.style.fontWeight = '600';
                      summary.style.outline = 'none';
                      summary.innerText = '🕒 #' + (idx + 2) + ' – ' + (c.data || '') + ' – por ' + (c.autor || 'Sistema');

                      var bodyDiv = document.createElement('div');
                      bodyDiv.style.marginTop = '8px';
                      bodyDiv.style.whiteSpace = 'pre-wrap';
                      bodyDiv.style.lineHeight = '1.4';
                      bodyDiv.style.color = 'inherit';
                      bodyDiv.innerText = c.texto || '';

                      details.appendChild(summary);
                      details.appendChild(bodyDiv);
                      commentsListEl.appendChild(details);
                    });
                  } else {
                    commentsContainer.style.display = 'none';
                  }
                  break;

                case 'portaria':
                  document.getElementById('blocoPortaria').style.display = 'block';

                  badge.className = 'modal-badge';
                  if (props.is_ferias) {
                    badge.style.backgroundColor = 'rgba(13, 148, 136, 0.2)';
                    badge.style.color = '#0d9488';
                    badge.style.border = '1px solid #0d9488';
                  } else if (props.is_fiscal) {
                    badge.style.backgroundColor = 'rgba(139, 92, 246, 0.2)';
                    badge.style.color = '#8b5cf6';
                    badge.style.border = '1px solid #8b5cf6';
                  } else {
                    badge.style.backgroundColor = 'rgba(100, 116, 139, 0.2)';
                    badge.style.color = '#94a3b8';
                    badge.style.border = '1px solid #94a3b8';
                  }

                  document.getElementById('poTitulo').innerText = props.titulo || info.event.title;
                  document.getElementById('poMembros').innerText = props.membros || 'Não informado';
                  document.getElementById('poDataPub').innerText = props.data_publicacao || props.data_emissao || 'N/A';
                  document.getElementById('poEmenta').innerText = props.ementa || 'Sem ementa disponível.';

                  var poPdfLink = document.getElementById('poPdfLink');
                  var poPdfContainer = document.getElementById('poPdfContainer');
                  if (props.pdf_url) {
                    poPdfLink.href = props.pdf_url;
                    poPdfContainer.style.display = 'block';
                  } else {
                    poPdfContainer.style.display = 'none';
                  }
                  break;

                default:
                  break;
              }

              document.getElementById('calendarioModal').style.display = 'flex';
            }
          });
          calendar.render();
        }

        if (document.readyState === 'complete' || document.readyState === 'interactive') {
          initCalendar();
        } else {
          document.addEventListener('DOMContentLoaded', initCalendar);
        }
      </script>
    </body>
    </html>
    """
    calendar_html = html_template.replace("__EVENTS_JSON__", events_json)
    st.components.v1.html(calendar_html, height=height_px, scrolling=scrolling_enabled)
