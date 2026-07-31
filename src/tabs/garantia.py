import io
import streamlit as st
import streamlit.components.v1 as components
import pandas as pd
from datetime import datetime
from src.database import get_garantia_contratos_df, get_garantia_chamados_df, sync_garantia_from_excel
from src.components.pagination import (
    render_items_per_page_selector,
    paginate_items,
    render_pagination_controls
)



def parse_date_to_iso_and_br(date_val):
    if pd.isna(date_val) or not date_val:
        return None, None
    try:
        dt = pd.to_datetime(date_val, dayfirst=True, errors='coerce')
        if pd.isna(dt):
            return None, None
        return dt.strftime('%Y-%m-%d'), dt.strftime('%d/%m/%Y')
    except Exception:
        return None, None


def render_garantia_calendar(events: list[dict]):
    """Renderiza o componente interativo FullCalendar (v6) para a aba de Garantia com suporte a Tema Claro e Escuro."""
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
          border-radius: 12px;
          width: 90%;
          max-width: 520px;
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

      <!-- COMPONENTE MODAL CUSTOMIZADO -->
      <div id="garantiaModal" class="modal-overlay">
        <div class="modal-content">
          <div class="modal-header">
            <h3 id="mTitle" class="modal-title">Detalhes da Garantia</h3>
            <button class="close-btn" onclick="closeModal()">&times;</button>
          </div>
          <div class="modal-body">
            <span id="mBadge" class="modal-badge"></span>
            <p><strong>📜 Contrato:</strong> <span id="mContrato"></span></p>
            <p><strong>📂 PU SAJ:</strong> <span id="mPuSaj"></span></p>
            <p><strong>💻 Item / Equipamento:</strong> <span id="mItem"></span></p>
            <p><strong>🏢 Fornecedor:</strong> <span id="mFornecedor"></span></p>
            <p><strong>📅 Data do Evento:</strong> <span id="mData"></span></p>
            <p><strong>🛡️ Status Vigência:</strong> <span id="mStatus"></span></p>
            <p id="pNota"><strong>📄 Nota Fiscal:</strong> <span id="mNota"></span></p>
            <p id="pLink"><strong>🌐 Suporte:</strong> <a id="mLink" href="#" target="_blank" style="color: #3b82f6; text-decoration: underline;">Abrir Chamado / Suporte ↗</a></p>
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
          document.getElementById('garantiaModal').style.display = 'none';
        }}

        window.onclick = function(event) {{
          var modal = document.getElementById('garantiaModal');
          if (event.target == modal) {{
            modal.display = 'none';
          }}
        }};

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
              document.getElementById('mContrato').innerText = props.contrato || 'N/A';
              document.getElementById('mPuSaj').innerText = props.pu_saj || 'N/A';
              document.getElementById('mItem').innerText = props.item || 'N/A';
              document.getElementById('mFornecedor').innerText = props.fornecedor || 'N/A';
              document.getElementById('mData').innerText = props.data_formatada || 'N/A';
              document.getElementById('mStatus').innerText = props.status_garantia || 'N/A';

              var mNota = document.getElementById('mNota');
              var pNota = document.getElementById('pNota');
              if (props.nota_fiscal) {{
                mNota.innerText = props.nota_fiscal;
                pNota.style.display = 'block';
              }} else {{
                pNota.style.display = 'none';
              }}

              var mLink = document.getElementById('mLink');
              var pLink = document.getElementById('pLink');
              if (props.link_suporte) {{
                mLink.href = props.link_suporte;
                pLink.style.display = 'block';
              }} else {{
                pLink.style.display = 'none';
              }}
              
              var badge = document.getElementById('mBadge');
              badge.innerText = props.tipo || 'Garantia';
              if (props.tipo && props.tipo.includes('Início')) {{
                badge.style.backgroundColor = 'rgba(16, 185, 129, 0.2)';
                badge.style.color = '#10b981';
                badge.style.border = '1px solid #10b981';
              }} else {{
                badge.style.backgroundColor = 'rgba(239, 68, 68, 0.2)';
                badge.style.color = '#ef4444';
                badge.style.border = '1px solid #ef4444';
              }}

              document.getElementById('garantiaModal').style.display = 'flex';
            }}
          }});
          calendar.render();
        }});
      </script>
    </body>
    </html>
    """
    components.html(calendar_html, height=750, scrolling=False)


def render_garantia_page():
    st.markdown("# 🛡️ Sistema de Controle de Garantia")
    st.caption("Acompanhe os contratos de garantia, vigências e chamados de manutenção abertos junto aos fornecedores.")
    
    col_syn1, col_syn2 = st.columns([3, 1])
    with col_syn2:
        if st.button("🔄 Sincronizar com Excel", type="primary", use_container_width=True):
            with st.spinner("Lendo dados da planilha Excel oficial..."):
                ok = sync_garantia_from_excel()
                if ok:
                    st.success("Dados de garantia sincronizados com sucesso!")
                    st.rerun()
                else:
                    st.error("Não foi possível encontrar a planilha no caminho configurado.")

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
    GARANTIA_SUBTAB_REVERSE = {v: k for k, v in GARANTIA_SUBTAB_MAP.items()}

    url_subtab = st.query_params.get("subtab", "contratos")
    default_title = GARANTIA_SUBTAB_MAP.get(url_subtab, "📜 Contratos de Garantia")
    options = list(GARANTIA_SUBTAB_MAP.values())
    default_idx = options.index(default_title) if default_title in options else 0

    selected_subtab = st.radio(
        "Navegação da Garantia:",
        options=options,
        index=default_idx,
        horizontal=True,
        label_visibility="collapsed",
        key="garantia_subtab_radio"
    )

    new_slug = GARANTIA_SUBTAB_REVERSE.get(selected_subtab, "contratos")
    if st.query_params.get("subtab") != new_slug:
        st.query_params["subtab"] = new_slug

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
                use_container_width=True
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
            use_container_width=True
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
                use_container_width=True
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
            use_container_width=True
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
        render_garantia_calendar(events)

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
                st.bar_chart(data=st_counts, x='Status do Chamado', y='Quantidade', use_container_width=True)
            else:
                st.info("Sem dados de chamados para exibir gráfico.")

        with g2:
            st.markdown("### 🏢 Contratos por Fornecedor")
            if not df_contratos.empty and 'fornecedor' in df_contratos.columns:
                forn_counts = df_contratos['fornecedor'].value_counts().reset_index()
                forn_counts.columns = ['Fornecedor', 'Quantidade']
                st.bar_chart(data=forn_counts, x='Fornecedor', y='Quantidade', use_container_width=True)
            else:
                st.info("Sem dados de contratos para exibir gráfico.")
