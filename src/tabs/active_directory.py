import os
import io
import re
import json
import pandas as pd
import streamlit as st
import streamlit.components.v1 as components
from datetime import datetime

from src.database.ad_db import (
    get_ad_ous,
    get_ad_users_df,
    get_ad_departments,
    get_ad_groups_df,
    get_group_members,
    get_user_groups,
    get_ad_computers_df,
    get_ad_operating_systems,
    get_ad_sync_meta,
    get_ad_ou_stats,
    get_ou_members,
    get_ou_computers,
    get_all_ou_entities_compact
)
from src.services.ad_ldap_service import test_ad_connection, sync_active_directory_cache
from src.components.subtabs import render_subtabs
from src.components.metric_cards import render_metric_cards
from src.components.pagination import (
    render_items_per_page_selector,
    paginate_items,
    render_pagination_controls
)
from src.config import DOMINIO


def _sanitize_df_for_excel(df: pd.DataFrame) -> pd.DataFrame:
    """Remove caracteres de controle invisíveis de todas as colunas de texto para compatibilidade com openpyxl."""
    clean_df = df.copy()
    control_char_re = re.compile(r'[\x00-\x08\x0b\x0c\x0e-\x1f\x7f-\x9f]')
    for col in clean_df.columns:
        if clean_df[col].dtype == 'object':
            clean_df[col] = clean_df[col].apply(
                lambda val: control_char_re.sub('', str(val)).strip() if pd.notna(val) else val
            )
    return clean_df


def _build_gojs_tree_data(df_ous: pd.DataFrame, ou_stats: dict = None) -> list:
    """
    Transforma o DataFrame de OUs no formato de nós para o GoJS TreeModel.
    Injeta um nó Raiz centralizador ("ROOT_DOMAIN") para conectar todas as OUs de topo,
    garantindo que o GoJS renderize e expanda a árvore com 100% de consistência.
    Enriquece cada nó com contadores de Usuários, Computadores e Fichas de Usuários Amostra.
    """
    nodes = []
    if df_ous.empty:
        return nodes

    ou_stats = ou_stats or {}
    root_key = "ROOT_DOMAIN"
    domain_name = DOMINIO or "in.mpe.ms.gov.br"
    
    # Total de usuários e computadores somados para o nó raiz
    total_domain_users = sum(s.get("user_count", 0) for s in ou_stats.values())
    total_domain_comps = sum(s.get("comp_count", 0) for s in ou_stats.values())

    # Adiciona nó raiz mestre do domínio (em GoJS TreeModel, nó raiz não deve ter a chave 'parent' presente)
    nodes.append({
        "key": root_key,
        "name": f"Domínio: {domain_name}",
        "description": "Floresta e Raiz Corporativa do Active Directory",
        "isRoot": True,
        "userCount": total_domain_users,
        "compCount": total_domain_comps,
        "sampleUsers": []
    })

    dn_set = set(df_ous["dn"].dropna().tolist())

    # Contagem de filhos diretos para cada nó
    parent_counts = {}
    for _, row in df_ous.iterrows():
        parent = str(row["parent_dn"]) if pd.notna(row["parent_dn"]) else ""
        parent_key = parent if (parent and parent in dn_set) else root_key
        parent_counts[parent_key] = parent_counts.get(parent_key, 0) + 1

    # Atualiza contagem de filhos da raiz
    nodes[0]["childCount"] = parent_counts.get(root_key, 0)
    nodes[0]["level"] = 0

    for _, row in df_ous.iterrows():
        dn = str(row["dn"])
        name = str(row["name"])
        parent = str(row["parent_dn"]) if pd.notna(row["parent_dn"]) else ""
        desc = str(row.get("description", "")) if pd.notna(row.get("description")) else ""

        # Se o parent_dn for outra OU, conecta a ela; caso contrário, conecta à raiz corporativa
        parent_key = parent if (parent and parent in dn_set) else root_key
        is_level1 = (parent_key == root_key)

        st_info = ou_stats.get(dn, {})
        u_count = st_info.get("user_count", 0)
        c_count = st_info.get("comp_count", 0)
        s_users = st_info.get("sample_users", [])

        nodes.append({
            "key": dn,
            "name": name,
            "parent": parent_key,
            "description": desc,
            "isRoot": False,
            "level": 1 if is_level1 else 2,
            "childCount": parent_counts.get(dn, 0),
            "userCount": u_count,
            "compCount": c_count,
            "sampleUsers": s_users
        })

    return nodes


def render_gojs_tree_component(nodes: list, ou_entities: dict = None, height: int = 750):
    """
    Renderiza um diagrama GoJS interativo no Streamlit com busca instantânea profunda (OUs, usuários e computadores),
    destaque visual luminoso, expansão automática de pais, navegação entre resultados (X de N),
    e modais integrados para listagem completa de membros/computadores e fichas detalhadas de usuários e computadores.
    """
    # Carregar GoJS local se existir para garantir funcionamento 100% offline / sem bloqueio de CDN corporativo
    local_gojs_path = os.path.join(os.path.dirname(__file__), "..", "..", "assets", "js", "go.js")
    inline_gojs_script = ""
    if os.path.exists(local_gojs_path):
        try:
            with open(local_gojs_path, "r", encoding="utf-8") as f:
                inline_gojs_script = f"<script>{f.read()}</script>"
        except Exception:
            pass

    ou_entities = ou_entities or {"users": {}, "comps": {}}
    json_data = json.dumps(nodes)
    entities_data = json.dumps(ou_entities)

    html_code = f"""
    <!DOCTYPE html>
    <html lang="pt-BR">
    <head>
      <meta charset="UTF-8">
      {inline_gojs_script}
      <script>
        if (typeof go === 'undefined') {{
          document.write('<script src="https://cdnjs.cloudflare.com/ajax/libs/gojs/2.3.17/go.js"><\\/script>');
        }}
      </script>
      <script>
        if (typeof go === 'undefined') {{
          document.write('<script src="https://cdn.jsdelivr.net/npm/gojs@2.3.17/release/go.js"><\\/script>');
        }}
      </script>
      <style>
        * {{
          box-sizing: border-box;
          font-family: -apple-system, BlinkMacSystemFont, "Segoe UI", Roboto, Helvetica, Arial, sans-serif;
        }}
        html, body {{
          margin: 0;
          padding: 0;
          width: 100%;
          height: {height}px;
          background-color: #0b0f19;
          color: #f1f5f9;
          overflow: hidden;
          position: relative;
        }}
        #toolbar {{
          display: flex;
          align-items: center;
          gap: 8px;
          padding: 8px 14px;
          background: #111827;
          border-bottom: 1px solid #1f2937;
          height: 54px;
          z-index: 10;
          position: relative;
        }}
        .search-container {{
          position: relative;
          display: flex;
          align-items: center;
          gap: 6px;
          flex: 1;
          max-width: 440px;
        }}
        #searchBox {{
          width: 100%;
          padding: 7px 12px 7px 32px;
          border-radius: 8px;
          border: 1px solid #374151;
          background: #1f2937;
          color: #ffffff;
          font-size: 13px;
          outline: none;
          transition: all 0.2s;
        }}
        .search-icon {{
          position: absolute;
          left: 10px;
          color: #94a3b8;
          font-size: 13px;
          pointer-events: none;
        }}
        #searchBox:focus {{
          border-color: #38bdf8;
          background: #111827;
          box-shadow: 0 0 0 2px rgba(56, 189, 248, 0.25);
        }}
        #searchStatus {{
          font-size: 11px;
          font-weight: 600;
          color: #94a3b8;
          white-space: nowrap;
          min-width: 70px;
        }}
        .btn {{
          display: inline-flex;
          align-items: center;
          gap: 5px;
          background: #1f2937;
          color: #e2e8f0;
          border: 1px solid #374151;
          padding: 6px 12px;
          border-radius: 8px;
          cursor: pointer;
          font-size: 12px;
          font-weight: 500;
          transition: all 0.15s ease-in-out;
          user-select: none;
          white-space: nowrap;
        }}
        .btn:hover {{
          background: #374151;
          border-color: #64748b;
          color: #ffffff;
          transform: translateY(-1px);
        }}
        .btn:active {{
          transform: translateY(0px);
        }}
        .btn-primary {{
          background: #0284c7;
          border-color: #38bdf8;
          color: #ffffff;
        }}
        .btn-primary:hover {{
          background: #0369a1;
          border-color: #7dd3fc;
        }}
        .btn-icon {{
          padding: 6px 9px;
        }}
        .btn-group {{
          display: flex;
          align-items: center;
          gap: 4px;
        }}
        #infoBadge {{
          margin-left: auto;
          display: flex;
          align-items: center;
          gap: 8px;
          font-size: 12px;
          color: #94a3b8;
          background: #1e293b;
          padding: 4px 12px;
          border-radius: 9999px;
          border: 1px solid #334155;
          white-space: nowrap;
        }}
        #infoBadge b {{
          color: #38bdf8;
        }}
        #myDiagramDiv {{
          width: 100%;
          height: {height - 54}px !important;
          min-height: {height - 54}px !important;
          background: radial-gradient(circle at center, #0f172a 0%, #080d1a 100%);
          position: relative;
        }}

        /* Modal Overlay & Card Refinado */
        .ad-modal-backdrop {{
          display: none;
          position: absolute;
          top: 0;
          left: 0;
          width: 100%;
          height: 100%;
          background: rgba(4, 7, 15, 0.78);
          backdrop-filter: blur(6px);
          z-index: 100;
          align-items: center;
          justify-content: center;
          padding: 20px;
        }}
        .ad-modal-backdrop.open {{
          display: flex;
        }}
        .ad-modal-card {{
          background: #0f172a;
          border: 1px solid #334155;
          border-radius: 14px;
          box-shadow: 0 25px 50px -12px rgba(0, 0, 0, 0.7), 0 0 0 1px rgba(255, 255, 255, 0.05);
          width: 100%;
          max-width: 820px;
          max-height: 90vh;
          display: flex;
          flex-direction: column;
          overflow: hidden;
          animation: modalFadeIn 0.18s ease-out;
        }}
        @keyframes modalFadeIn {{
          from {{ opacity: 0; transform: scale(0.96) translateY(8px); }}
          to {{ opacity: 1; transform: scale(1) translateY(0); }}
        }}
        .ad-modal-header {{
          padding: 16px 20px;
          background: #1e293b;
          border-bottom: 1px solid #334155;
          display: flex;
          align-items: center;
          justify-content: space-between;
        }}
        .ad-modal-title {{
          font-size: 16px;
          font-weight: 600;
          color: #f8fafc;
          display: flex;
          align-items: center;
          gap: 10px;
        }}
        .ad-modal-close {{
          background: transparent;
          border: none;
          color: #94a3b8;
          font-size: 20px;
          cursor: pointer;
          width: 32px;
          height: 32px;
          border-radius: 6px;
          display: flex;
          align-items: center;
          justify-content: center;
          transition: all 0.15s;
        }}
        .ad-modal-close:hover {{
          background: #334155;
          color: #ffffff;
        }}
        .ad-modal-body {{
          padding: 18px 20px;
          overflow-y: auto;
          flex: 1;
        }}
        .ad-tabs {{
          display: flex;
          gap: 8px;
          border-bottom: 1px solid #1e293b;
          margin-bottom: 14px;
          padding-bottom: 6px;
        }}
        .ad-tab-btn {{
          background: transparent;
          border: none;
          color: #94a3b8;
          font-size: 13px;
          font-weight: 600;
          padding: 6px 12px;
          border-radius: 6px;
          cursor: pointer;
          transition: all 0.15s;
        }}
        .ad-tab-btn.active {{
          background: #1e293b;
          color: #38bdf8;
        }}
        .modal-search-box {{
          position: relative;
          display: flex;
          align-items: center;
          min-width: 280px;
          flex: 1;
          max-width: 380px;
        }}
        .modal-search-box input {{
          width: 100%;
          padding: 6px 10px 6px 28px;
          border-radius: 6px;
          border: 1px solid #334155;
          background: #1e293b;
          color: #ffffff;
          font-size: 12px;
          outline: none;
          transition: all 0.2s;
        }}
        .modal-search-box input:focus {{
          border-color: #38bdf8;
          box-shadow: 0 0 0 2px rgba(56, 189, 248, 0.2);
        }}
        .modal-search-icon {{
          position: absolute;
          left: 8px;
          font-size: 11px;
          color: #94a3b8;
          pointer-events: none;
        }}
        .member-table {{
          width: 100%;
          border-collapse: collapse;
          font-size: 12px;
        }}
        .member-table th {{
          text-align: left;
          padding: 8px 10px;
          background: #1e293b;
          color: #94a3b8;
          font-weight: 600;
          border-bottom: 1px solid #334155;
          position: sticky;
          top: 0;
        }}
        .member-table td {{
          padding: 8px 10px;
          border-bottom: 1px solid #1e293b;
          color: #e2e8f0;
        }}
        .member-table tr:hover td {{
          background: rgba(56, 189, 248, 0.08);
          cursor: pointer;
        }}
        .badge-active {{
          display: inline-block;
          padding: 2px 7px;
          border-radius: 4px;
          font-size: 10px;
          font-weight: 700;
          background: rgba(16, 185, 129, 0.2);
          color: #34d399;
          border: 1px solid rgba(16, 185, 129, 0.4);
        }}
        .badge-inactive {{
          display: inline-block;
          padding: 2px 7px;
          border-radius: 4px;
          font-size: 10px;
          font-weight: 700;
          background: rgba(239, 68, 68, 0.2);
          color: #f87171;
          border: 1px solid rgba(239, 68, 68, 0.4);
        }}
        .grid-details {{
          display: grid;
          grid-template-columns: 1fr 1fr;
          gap: 14px;
        }}
        .detail-card {{
          background: #1e293b;
          border: 1px solid #334155;
          border-radius: 8px;
          padding: 12px 14px;
        }}
        .detail-card h4 {{
          margin: 0 0 8px 0;
          font-size: 13px;
          color: #38bdf8;
          border-bottom: 1px solid #334155;
          padding-bottom: 4px;
        }}
        .detail-field {{
          margin-bottom: 6px;
          font-size: 12px;
          display: flex;
          flex-direction: column;
        }}
        .detail-field label {{
          color: #94a3b8;
          font-size: 11px;
          font-weight: 600;
        }}
        .detail-field span {{
          color: #f1f5f9;
          word-break: break-all;
        }}
        .quick-actions {{
          display: grid;
          grid-template-columns: 1fr 1fr 1fr;
          gap: 10px;
          margin-top: 14px;
        }}
        .action-link {{
          text-decoration: none;
        }}
        .action-btn {{
          padding: 10px;
          border-radius: 8px;
          text-align: center;
          font-weight: 600;
          font-size: 12px;
          cursor: pointer;
          transition: all 0.15s;
          display: flex;
          flex-direction: column;
          gap: 4px;
          align-items: center;
        }}
        .action-btn:hover {{
          transform: translateY(-2px);
          filter: brightness(1.15);
        }}
        .action-btn.rdp {{
          background: #1e293b;
          border: 1px solid #3b82f6;
          color: #60a5fa;
        }}
        .action-btn.share {{
          background: #1e293b;
          border: 1px solid #10b981;
          color: #34d399;
        }}
        .action-btn.ping {{
          background: #1e293b;
          border: 1px solid #f59e0b;
          color: #fbbf24;
        }}
      </style>
    </head>
    <body>
      <div id="toolbar">
        <div class="search-container">
          <span class="search-icon">🔍</span>
          <input type="text" id="searchBox" placeholder="Buscar OU, Usuário, Computador ou Cargo..." />
          <span id="searchStatus"></span>
          <button class="btn btn-icon" id="btnPrevMatch" title="Resultado Anterior" style="display:none;">◀</button>
          <button class="btn btn-icon" id="btnNextMatch" title="Próximo Resultado" style="display:none;">▶</button>
        </div>
        <button class="btn" id="btnExpandAll">➕ Expandir Tudo</button>
        <button class="btn" id="btnCollapseAll">➖ Recolher Tudo</button>
        <button class="btn btn-primary" id="btnFocusRoot">🎯 Focar na Raiz</button>
        <div class="btn-group">
          <button class="btn" id="btnZoomIn" title="Aumentar Zoom">🔍 +</button>
          <button class="btn" id="btnZoomOut" title="Diminuir Zoom">🔎 -</button>
        </div>
        <div id="infoBadge">Total: <b id="nodeCount">{len(nodes) - 1}</b> OUs</div>
      </div>
      <div id="myDiagramDiv"></div>

      <!-- MODAL LISTAGEM DE MEMBROS E COMPUTADORES DA OU -->
      <div class="ad-modal-backdrop" id="ouModalBackdrop">
        <div class="ad-modal-card">
          <div class="ad-modal-header">
            <div class="ad-modal-title">
              <span>📁</span>
              <span id="ouModalTitle">Membros da OU</span>
            </div>
            <button class="ad-modal-close" onclick="closeOuModal()">✕</button>
          </div>
          <div class="ad-modal-body">
            <div style="display: flex; align-items: center; justify-content: space-between; gap: 12px; margin-bottom: 12px; flex-wrap: wrap;">
              <div class="ad-tabs" style="margin-bottom: 0; border-bottom: none;">
                <button class="ad-tab-btn active" id="tabBtnUsers" onclick="switchOuTab('users')">👥 Usuários (<span id="ouUsersCount">0</span>)</button>
                <button class="ad-tab-btn" id="tabBtnComps" onclick="switchOuTab('comps')">💻 Computadores (<span id="ouCompsCount">0</span>)</button>
              </div>
              <div class="modal-search-box">
                <span class="modal-search-icon">🔍</span>
                <input type="text" id="modalSearchInput" placeholder="Filtrar nesta OU (Nome, Login, E-mail, Máquina...)" oninput="filterOuModalTable()" />
              </div>
            </div>
            <div id="ouUsersView">
              <table class="member-table">
                <thead>
                  <tr>
                    <th>Status</th>
                    <th>Nome / Display Name</th>
                    <th>Login (sAM)</th>
                    <th>Cargo / Função</th>
                    <th>E-mail</th>
                    <th>Ramal</th>
                    <th>Cartão RFID</th>
                  </tr>
                </thead>
                <tbody id="ouUsersTbody"></tbody>
              </table>
            </div>
            <div id="ouCompsView" style="display: none;">
              <table class="member-table">
                <thead>
                  <tr>
                    <th>Status</th>
                    <th>Nome da Máquina</th>
                    <th>FQDN / Hostname</th>
                    <th>Sistema Operacional</th>
                    <th>Versão</th>
                    <th>Gerenciado Por</th>
                  </tr>
                </thead>
                <tbody id="ouCompsTbody"></tbody>
              </table>
            </div>
          </div>
        </div>
      </div>

      <!-- MODAL DETALHE DE USUÁRIO ESPECÍFICO -->
      <div class="ad-modal-backdrop" id="userDetailModalBackdrop">
        <div class="ad-modal-card" style="max-width: 680px;">
          <div class="ad-modal-header">
            <div class="ad-modal-title">
              <span>👤</span>
              <span id="userDetailTitle">Ficha do Usuário</span>
            </div>
            <button class="ad-modal-close" onclick="closeUserDetailModal()">✕</button>
          </div>
          <div class="ad-modal-body" id="userDetailContent"></div>
        </div>
      </div>

      <!-- MODAL DETALHE DE COMPUTADOR ESPECÍFICO COM AÇÕES RÁPIDAS -->
      <div class="ad-modal-backdrop" id="compDetailModalBackdrop">
        <div class="ad-modal-card" style="max-width: 680px;">
          <div class="ad-modal-header">
            <div class="ad-modal-title">
              <span>💻</span>
              <span id="compDetailTitle">Ficha do Computador</span>
            </div>
            <button class="ad-modal-close" onclick="closeCompDetailModal()">✕</button>
          </div>
          <div class="ad-modal-body" id="compDetailContent"></div>
        </div>
      </div>

      <script>
        const ouEntities = {entities_data};
        const nodeDataArray = {json_data};

        let currentMatches = [];
        let currentMatchIndex = -1;
        let activeOuDn = null;

        if (typeof go === 'undefined') {{
          document.getElementById('myDiagramDiv').innerHTML = '<div style="padding: 40px; color: #f87171; text-align: center;"><h3>⚠️ Não foi possível carregar a biblioteca GoJS</h3><p>Verifique o acesso à internet corporativa para carregar o script.</p></div>';
        }} else {{
          initDiagram();
        }}

        function initDiagram() {{
          const $ = go.GraphObject.make;

          // Paleta de cores corporativa refinada (estilo OrgChart Editor & FamilyTree)
          const colors = {{
            rootHeader: "#0284c7",
            rootFill: "#0c4a6e",
            rootBorder: "#38bdf8",
            
            l1Header: "#6366f1",
            l1Fill: "#1e1b4b",
            l1Border: "#818cf8",
            
            subHeader: "#334155",
            subFill: "#0f172a",
            subBorder: "#475569",

            textPrimary: "#f8fafc",
            textMuted: "#94a3b8",
            highlightBorder: "#f59e0b",
            highlightFill: "#451a03"
          }};

          window.myDiagram = $(go.Diagram, "myDiagramDiv", {{
            "undoManager.isEnabled": false,
            initialContentAlignment: go.Spot.TopCenter,
            initialScale: 0.95,
            minScale: 0.15,
            maxScale: 2.0,
            allowCopy: false,
            allowDelete: false,
            layout: $(go.TreeLayout, {{
              angle: 90,
              layerSpacing: 45,
              nodeSpacing: 22,
              alignment: go.TreeLayout.AlignmentCenterChildren,
              compaction: go.TreeLayout.CompactionBlock
            }}),
            "animationManager.isEnabled": false
          }});

          // Template para cada usuário listado dentro do Card (OrgChart style com e-mail e ramal)
          const userItemTemplate = $(go.Panel, "Vertical",
            {{
              margin: new go.Margin(2, 0, 2, 0),
              defaultStretch: go.GraphObject.Horizontal,
              cursor: "pointer",
              click: (e, obj) => {{
                e.handled = true;
                const u = obj.data;
                openUserDetailByName(u.name);
              }}
            }},
            $(go.Panel, "Auto",
              $(go.Shape, "RoundedRectangle", {{
                fill: "#1e293b",
                stroke: "#334155",
                strokeWidth: 1,
                parameter1: 4
              }}),
              $(go.Panel, "Vertical", {{ margin: new go.Margin(4, 6, 4, 6) }},
                // Linha 1: Ícone + Nome do Usuário
                $(go.Panel, "Horizontal", {{ alignment: go.Spot.Left }},
                  $(go.TextBlock, "👤 ", {{ font: "10px sans-serif" }}),
                  $(go.TextBlock, {{
                    font: "bold 11px -apple-system, BlinkMacSystemFont, sans-serif",
                    stroke: "#38bdf8",
                    maxSize: new go.Size(200, NaN),
                    wrap: go.TextBlock.WrapFit
                  }},
                  new go.Binding("text", "name"))
                ),
                // Linha 2: Cargo (se houver)
                $(go.TextBlock, {{
                  font: "italic 10px -apple-system, BlinkMacSystemFont, sans-serif",
                  stroke: "#94a3b8",
                  margin: new go.Margin(1, 0, 0, 14),
                  maxSize: new go.Size(200, NaN),
                  wrap: go.TextBlock.WrapFit
                }},
                new go.Binding("text", "title"),
                new go.Binding("visible", "title", t => Boolean(t))),

                // Linha 3: E-mail e/ou Ramal
                $(go.Panel, "Horizontal", {{ margin: new go.Margin(2, 0, 0, 14), alignment: go.Spot.Left }},
                  new go.Binding("visible", "", u => Boolean(u.mail || u.phone)),
                  $(go.TextBlock, {{
                    font: "10px -apple-system, BlinkMacSystemFont, sans-serif",
                    stroke: "#cbd5e1"
                  }},
                  new go.Binding("text", "", u => {{
                    const parts = [];
                    if (u.mail) parts.push("✉️ " + u.mail);
                    if (u.phone) parts.push("📞 " + u.phone);
                    return parts.join("  |  ");
                  }}))
                )
              )
            )
          );

          // Template do Nó OrgChart / FamilyTree Card
          myDiagram.nodeTemplate = $(go.Node, "Auto",
            {{
              cursor: "pointer",
              selectionAdorned: false
            }},
            // Moldura externa e sombra visual
            $(go.Shape, "RoundedRectangle", {{
              fill: colors.subFill,
              stroke: colors.subBorder,
              strokeWidth: 1.5,
              parameter1: 8
            }},
            new go.Binding("fill", "level", lvl => {{
              if (lvl === 0) return colors.rootFill;
              if (lvl === 1) return colors.l1Fill;
              return colors.subFill;
            }}),
            new go.Binding("stroke", "level", lvl => {{
              if (lvl === 0) return colors.rootBorder;
              if (lvl === 1) return colors.l1Border;
              return colors.subBorder;
            }}),
            new go.Binding("stroke", "isHighlighted", h => h ? colors.highlightBorder : null).ofObject(),
            new go.Binding("strokeWidth", "isHighlighted", h => h ? 3.5 : 1.5).ofObject(),
            new go.Binding("fill", "isHighlighted", h => h ? colors.highlightFill : null).ofObject()
            ),

            // Conteúdo interno do Card
            $(go.Panel, "Vertical", {{ margin: 0, defaultStretch: go.GraphObject.Horizontal }},
              // Faixa de cabeçalho colorida (OrgChart style)
              $(go.Panel, "Auto", {{ stretch: go.GraphObject.Horizontal }},
                $(go.Shape, "RoundedRectangle", {{
                  parameter1: 7,
                  strokeWidth: 0
                }},
                new go.Binding("fill", "level", lvl => {{
                  if (lvl === 0) return colors.rootHeader;
                  if (lvl === 1) return colors.l1Header;
                  return colors.subHeader;
                }})
                ),
                $(go.Panel, "Horizontal", {{ margin: new go.Margin(4, 8, 4, 8) }},
                  $(go.TextBlock, {{
                    font: "12px sans-serif",
                    margin: new go.Margin(0, 5, 0, 0)
                  }},
                  new go.Binding("text", "level", lvl => {{
                    if (lvl === 0) return "🏛️";
                    if (lvl === 1) return "🏢";
                    return "📁";
                  }})),
                  $(go.TextBlock, {{
                    font: "bold 10px -apple-system, BlinkMacSystemFont, sans-serif",
                    stroke: "#ffffff",
                    isMultiline: false
                  }},
                  new go.Binding("text", "level", lvl => {{
                    if (lvl === 0) return "DOMÍNIO CORPORATIVO";
                    if (lvl === 1) return "UNIDADE PRINCIPAL";
                    return "DEPARTAMENTO / OU";
                  }}))
                )
              ),

              // Corpo do card com Nome
              $(go.Panel, "Vertical", {{ margin: new go.Margin(8, 12, 4, 12) }},
                $(go.TextBlock, {{
                  font: "600 13px -apple-system, BlinkMacSystemFont, 'Segoe UI', Roboto, sans-serif",
                  stroke: colors.textPrimary,
                  maxSize: new go.Size(240, NaN),
                  wrap: go.TextBlock.WrapFit,
                  textAlign: "center",
                  alignment: go.Spot.Center
                }},
                new go.Binding("text", "name")),

                // Badges Interativos de Usuários, Computadores e Sub-unidades
                $(go.Panel, "Horizontal", {{
                  margin: new go.Margin(6, 0, 2, 0),
                  alignment: go.Spot.Center
                }},
                  // Badge Usuários (Click abre modal com todos os usuários da OU)
                  $(go.Panel, "Auto", {{
                    margin: new go.Margin(0, 4, 0, 0),
                    cursor: "pointer",
                    click: (e, obj) => {{
                      e.handled = true;
                      const nData = obj.part.data;
                      if (!nData.isRoot) openOuModal(nData.key, nData.name, 'users');
                    }}
                  }},
                  new go.Binding("visible", "userCount", u => (u > 0)),
                  $(go.Shape, "RoundedRectangle", {{ fill: "#1e1b4b", stroke: "#6366f1", strokeWidth: 1, parameter1: 4 }}),
                  $(go.TextBlock, {{
                    font: "bold 10px -apple-system, BlinkMacSystemFont, sans-serif",
                    stroke: "#a5b4fc",
                    margin: new go.Margin(2, 6, 2, 6)
                  }},
                  new go.Binding("text", "userCount", u => "👥 " + u + (u === 1 ? " user" : " users")))
                  ),

                  // Badge Computadores (Click abre modal com computadores da OU)
                  $(go.Panel, "Auto", {{
                    margin: new go.Margin(0, 4, 0, 0),
                    cursor: "pointer",
                    click: (e, obj) => {{
                      e.handled = true;
                      const nData = obj.part.data;
                      if (!nData.isRoot) openOuModal(nData.key, nData.name, 'comps');
                    }}
                  }},
                  new go.Binding("visible", "compCount", c => (c > 0)),
                  $(go.Shape, "RoundedRectangle", {{ fill: "#083344", stroke: "#06b6d4", strokeWidth: 1, parameter1: 4 }}),
                  $(go.TextBlock, {{
                    font: "bold 10px -apple-system, BlinkMacSystemFont, sans-serif",
                    stroke: "#67e8f9",
                    margin: new go.Margin(2, 6, 2, 6)
                  }},
                  new go.Binding("text", "compCount", c => "💻 " + c + (c === 1 ? " comp" : " comps")))
                  ),

                  // Badge Sub-unidades
                  $(go.Panel, "Auto",
                  new go.Binding("visible", "childCount", c => (c > 0)),
                  $(go.Shape, "RoundedRectangle", {{ fill: "#1e293b", stroke: "#475569", strokeWidth: 1, parameter1: 4 }}),
                  $(go.TextBlock, {{
                    font: "10px -apple-system, BlinkMacSystemFont, sans-serif",
                    stroke: colors.textMuted,
                    margin: new go.Margin(2, 5, 2, 5)
                  }},
                  new go.Binding("text", "childCount", c => "📁 " + c))
                  )
                )
              ),

              // Lista de Amostra de Usuários Alocados na OU
              $(go.Panel, "Vertical", {{
                margin: new go.Margin(2, 8, 4, 8),
                defaultStretch: go.GraphObject.Horizontal,
                itemTemplate: userItemTemplate
              }},
              new go.Binding("itemArray", "sampleUsers"),
              new go.Binding("visible", "sampleUsers", s => Boolean(s && s.length > 0))
              ),

              // Botão "+ Ver todos" se houver mais de 3 usuários
              $(go.Panel, "Auto", {{
                margin: new go.Margin(2, 8, 4, 8),
                cursor: "pointer",
                click: (e, obj) => {{
                  e.handled = true;
                  const nData = obj.part.data;
                  openOuModal(nData.key, nData.name, 'users');
                }}
              }},
              new go.Binding("visible", "userCount", u => (u > 3)),
              $(go.Shape, "RoundedRectangle", {{ fill: "#0f172a", stroke: "#334155", strokeWidth: 1, parameter1: 4 }}),
              $(go.TextBlock, {{
                font: "600 10px -apple-system, BlinkMacSystemFont, sans-serif",
                stroke: "#38bdf8",
                alignment: go.Spot.Center,
                margin: new go.Margin(3, 6, 3, 6)
              }},
              new go.Binding("text", "userCount", u => "Ver todos os " + u + " usuários →"))
              ),

              // Botão de expandir/recolher subárvore
              $("TreeExpanderButton", {{
                alignment: go.Spot.Bottom,
                alignmentFocus: go.Spot.Top,
                margin: new go.Margin(2, 0, 6, 0)
              }})
            )
          );

          // Links estilizados (OrgChart Orthogonal limpo)
          myDiagram.linkTemplate = $(go.Link, {{
            routing: go.Link.Orthogonal,
            corner: 8,
            selectable: false
          }},
            $(go.Shape, {{ strokeWidth: 1.75, stroke: "#475569" }})
          );

          myDiagram.model = new go.TreeModel(nodeDataArray);

          // Estado inicial elegante: Raiz expandida e 1º nível recolhido
          myDiagram.startTransaction("initDisplay");
          myDiagram.nodes.each(n => {{
            if (n.data.isRoot) {{
              n.isTreeExpanded = true;
            }} else {{
              n.isTreeExpanded = false;
            }}
          }});
          myDiagram.commitTransaction("initDisplay");

          setTimeout(() => {{
            focusRootNode();
          }}, 100);

          function focusRootNode() {{
            const root = myDiagram.findNodeForKey("ROOT_DOMAIN");
            if (root) {{
              myDiagram.scale = 0.95;
              myDiagram.centerRect(root.actualBounds);
            }} else {{
              myDiagram.commandHandler.zoomToFit();
            }}
          }}

          // ------------------------------------------------------------------
          // Event Listeners dos Botões de Navegação e Zoom
          // ------------------------------------------------------------------
          document.getElementById("btnExpandAll").addEventListener("click", () => {{
            myDiagram.startTransaction("expandAll");
            myDiagram.nodes.each(n => n.isTreeExpanded = true);
            myDiagram.commitTransaction("expandAll");
            myDiagram.commandHandler.zoomToFit();
          }});

          document.getElementById("btnCollapseAll").addEventListener("click", () => {{
            myDiagram.startTransaction("collapseAll");
            myDiagram.nodes.each(n => {{
              if (n.data.isRoot) {{
                n.isTreeExpanded = true;
              }} else {{
                n.collapseTree();
                n.isTreeExpanded = false;
              }}
            }});
            myDiagram.commitTransaction("collapseAll");
            focusRootNode();
          }});

          document.getElementById("btnFocusRoot").addEventListener("click", () => {{
            focusRootNode();
          }});

          document.getElementById("btnZoomIn").addEventListener("click", () => {{
            myDiagram.commandHandler.increaseZoom(1.2);
          }});

          document.getElementById("btnZoomOut").addEventListener("click", () => {{
            myDiagram.commandHandler.decreaseZoom(0.8);
          }});

          // ------------------------------------------------------------------
          // BUSCA PROFUNDA (OUs, Usuários, Computadores, Cargos)
          // ------------------------------------------------------------------
          const searchBox = document.getElementById("searchBox");
          const searchStatus = document.getElementById("searchStatus");
          const btnPrev = document.getElementById("btnPrevMatch");
          const btnNext = document.getElementById("btnNextMatch");

          searchBox.addEventListener("input", () => {{
            executeSearch();
          }});

          searchBox.addEventListener("keydown", (e) => {{
            if (e.key === "Enter") {{
              e.preventDefault();
              if (e.shiftKey) {{
                navigateMatch(-1);
              }} else {{
                navigateMatch(1);
              }}
            }}
          }});

          btnPrev.addEventListener("click", () => navigateMatch(-1));
          btnNext.addEventListener("click", () => navigateMatch(1));

          function executeSearch() {{
            const input = searchBox.value.trim().toLowerCase();
            myDiagram.startTransaction("searchDeep");

            if (!input) {{
              myDiagram.clearHighlighteds();
              myDiagram.commitTransaction("searchDeep");
              currentMatches = [];
              currentMatchIndex = -1;
              searchStatus.textContent = "";
              btnPrev.style.display = "none";
              btnNext.style.display = "none";
              return;
            }}

            myDiagram.clearHighlighteds();
            currentMatches = [];

            myDiagram.nodes.each(n => {{
              const d = n.data;
              if (d.isRoot) return;

              let matched = false;
              // 1. Nome da OU e Descrição
              if ((d.name && d.name.toLowerCase().includes(input)) || (d.description && d.description.toLowerCase().includes(input))) {{
                matched = true;
              }}

              // 2. Usuários alocados na OU
              if (!matched && ouEntities.users && ouEntities.users[d.key]) {{
                const uList = ouEntities.users[d.key];
                for (let i = 0; i < uList.length; i++) {{
                  const u = uList[i];
                  if ((u.name && u.name.toLowerCase().includes(input)) ||
                      (u.sam && u.sam.toLowerCase().includes(input)) ||
                      (u.mail && u.mail.toLowerCase().includes(input)) ||
                      (u.title && u.title.toLowerCase().includes(input)) ||
                      (u.pager && u.pager.toLowerCase().includes(input))) {{
                    matched = true;
                    break;
                  }}
                }}
              }}

              // 3. Computadores alocados na OU
              if (!matched && ouEntities.comps && ouEntities.comps[d.key]) {{
                const cList = ouEntities.comps[d.key];
                for (let j = 0; j < cList.length; j++) {{
                  const c = cList[j];
                  if ((c.name && c.name.toLowerCase().includes(input)) ||
                      (c.dns && c.dns.toLowerCase().includes(input)) ||
                      (c.os && c.os.toLowerCase().includes(input)) ||
                      (c.managed_by && c.managed_by.toLowerCase().includes(input))) {{
                    matched = true;
                    break;
                  }}
                }}
              }}

              if (matched) {{
                n.isHighlighted = true;
                currentMatches.push(n);

                // Expande todos os nós pais até a raiz
                let curr = n.findTreeParentNode();
                while (curr) {{
                  curr.isTreeExpanded = true;
                  curr = curr.findTreeParentNode();
                }}
              }}
            }});

            myDiagram.commitTransaction("searchDeep");

            if (currentMatches.length > 0) {{
              currentMatchIndex = 0;
              focusCurrentMatch();
              searchStatus.textContent = `1 de ${{currentMatches.length}}`;
              btnPrev.style.display = "inline-flex";
              btnNext.style.display = "inline-flex";
            }} else {{
              currentMatchIndex = -1;
              searchStatus.textContent = "Nenhum";
              btnPrev.style.display = "none";
              btnNext.style.display = "none";
            }}
          }}

          function navigateMatch(direction) {{
            if (currentMatches.length === 0) return;
            currentMatchIndex += direction;
            if (currentMatchIndex >= currentMatches.length) currentMatchIndex = 0;
            if (currentMatchIndex < 0) currentMatchIndex = currentMatches.length - 1;
            focusCurrentMatch();
            searchStatus.textContent = `${{currentMatchIndex + 1}} de ${{currentMatches.length}}`;
          }}

          function focusCurrentMatch() {{
            const targetNode = currentMatches[currentMatchIndex];
            if (targetNode) {{
              myDiagram.scale = 1.05;
              myDiagram.centerRect(targetNode.actualBounds);
            }}
          }}
        }}

        // ------------------------------------------------------------------
        // FUNÇÕES DOS MODAIS INTERATIVOS
        // ------------------------------------------------------------------
        function formatDateBr(val) {{
          if (!val || val === '-' || val === 'None' || val === 'null') return '-';
          try {{
            // Suporta formatos ISO 'YYYY-MM-DD HH:MM:SS' ou Date strings
            const str = String(val).trim();
            const m = str.match(/^(\d{4})-(\d{2})-(\d{2})[T\s](\d{2}):(\d{2}):(\d{2})/);
            if (m) {{
              return `${{m[3]}}/${{m[2]}}/${{m[1]}} ${{m[4]}}:${{m[5]}}:${{m[6]}}`;
            }}
            const d = new Date(str);
            if (!isNaN(d.getTime())) {{
              const pad = n => String(n).padStart(2, '0');
              return `${{pad(d.getDate())}}/${{pad(d.getMonth() + 1)}}/${{d.getFullYear()}} ${{pad(d.getHours())}}:${{pad(d.getMinutes())}}:${{pad(d.getSeconds())}}`;
            }}
          }} catch(e) {{}}
          return String(val);
        }}

        function openOuModal(ouDn, ouName, initialTab = 'users') {{
          activeOuDn = ouDn;
          document.getElementById('ouModalTitle').textContent = ouName || "Unidade Organizacional";
          document.getElementById('modalSearchInput').value = '';

          renderOuModalTables();
          switchOuTab(initialTab);
          document.getElementById('ouModalBackdrop').classList.add('open');
        }}

        function filterOuModalTable() {{
          renderOuModalTables();
        }}

        function renderOuModalTables() {{
          if (!activeOuDn) return;
          const query = (document.getElementById('modalSearchInput').value || '').trim().toLowerCase();
          const allUsers = (ouEntities.users && ouEntities.users[activeOuDn]) || [];
          const allComps = (ouEntities.comps && ouEntities.comps[activeOuDn]) || [];

          // Filtra usuários
          const filteredUsers = query ? allUsers.filter(u =>
            (u.name && u.name.toLowerCase().includes(query)) ||
            (u.sam && u.sam.toLowerCase().includes(query)) ||
            (u.mail && u.mail.toLowerCase().includes(query)) ||
            (u.title && u.title.toLowerCase().includes(query)) ||
            (u.pager && u.pager.toLowerCase().includes(query)) ||
            (u.phone && u.phone.toLowerCase().includes(query))
          ) : allUsers;

          // Filtra computadores
          const filteredComps = query ? allComps.filter(c =>
            (c.name && c.name.toLowerCase().includes(query)) ||
            (c.dns && c.dns.toLowerCase().includes(query)) ||
            (c.os && c.os.toLowerCase().includes(query)) ||
            (c.os_ver && c.os_ver.toLowerCase().includes(query)) ||
            (c.managed_by && c.managed_by.toLowerCase().includes(query))
          ) : allComps;

          document.getElementById('ouUsersCount').textContent = filteredUsers.length;
          document.getElementById('ouCompsCount').textContent = filteredComps.length;

          // Renderiza Tbody Usuários
          const tbodyUsers = document.getElementById('ouUsersTbody');
          if (filteredUsers.length === 0) {{
            tbodyUsers.innerHTML = `<tr><td colspan="7" style="text-align: center; color: #94a3b8; padding: 20px;">${{allUsers.length === 0 ? 'Nenhum usuário alocado diretamente nesta OU.' : 'Nenhum usuário encontrado para a busca.'}}</td></tr>`;
          }} else {{
            tbodyUsers.innerHTML = filteredUsers.map(u => `
              <tr onclick='openUserDetailDirect(${{JSON.stringify(JSON.stringify(u))}})'>
                <td><span class="${{u.active ? 'badge-active' : 'badge-inactive'}}">${{u.active ? 'ATIVO' : 'DESATIVADO'}}</span></td>
                <td><b style="color: #38bdf8;">${{u.name}}</b></td>
                <td><code>${{u.sam}}</code></td>
                <td>${{u.title || '-'}}</td>
                <td>${{u.mail || '-'}}</td>
                <td>${{u.phone || '-'}}</td>
                <td>${{u.pager ? '💳 <code>' + u.pager + '</code>' : '-'}}</td>
              </tr>
            `).join('');
          }}

          // Renderiza Tbody Computadores
          const tbodyComps = document.getElementById('ouCompsTbody');
          if (filteredComps.length === 0) {{
            tbodyComps.innerHTML = `<tr><td colspan="6" style="text-align: center; color: #94a3b8; padding: 20px;">${{allComps.length === 0 ? 'Nenhum computador alocado diretamente nesta OU.' : 'Nenhum computador encontrado para a busca.'}}</td></tr>`;
          }} else {{
            tbodyComps.innerHTML = filteredComps.map(c => `
              <tr onclick='openCompDetailDirect(${{JSON.stringify(JSON.stringify(c))}})'>
                <td><span class="${{c.active ? 'badge-active' : 'badge-inactive'}}">${{c.active ? 'ATIVO' : 'DESATIVADO'}}</span></td>
                <td><b style="color: #67e8f9;">💻 ${{c.name}}</b></td>
                <td><code>${{c.dns || '-'}}</code></td>
                <td>${{c.os || '-'}}</td>
                <td>${{c.os_ver || '-'}}</td>
                <td>${{c.managed_by || '-'}}</td>
              </tr>
            `).join('');
          }}
        }}

        function switchOuTab(tab) {{
          const isUsers = (tab === 'users');
          document.getElementById('tabBtnUsers').classList.toggle('active', isUsers);
          document.getElementById('tabBtnComps').classList.toggle('active', !isUsers);
          document.getElementById('ouUsersView').style.display = isUsers ? 'block' : 'none';
          document.getElementById('ouCompsView').style.display = !isUsers ? 'block' : 'none';
        }}

        function closeOuModal() {{
          document.getElementById('ouModalBackdrop').classList.remove('open');
        }}

        function openUserDetailByName(name) {{
          if (!ouEntities.users) return;
          for (let ou in ouEntities.users) {{
            const list = ouEntities.users[ou];
            const found = list.find(u => u.name === name || u.sam === name);
            if (found) {{
              openUserDetailDirect(JSON.stringify(found));
              return;
            }}
          }}
        }}

        function openUserDetailDirect(userJsonStr) {{
          const u = typeof userJsonStr === 'string' ? JSON.parse(userJsonStr) : userJsonStr;
          document.getElementById('userDetailTitle').textContent = u.name;

          const statusBadge = u.active
            ? '<span class="badge-active">🟢 CONTA ATIVA</span>'
            : '<span class="badge-inactive">🔴 CONTA DESATIVADA</span>';

          const rfidBox = u.pager
            ? `<div style="background: rgba(14, 165, 233, 0.15); border: 1px solid #0284c7; border-radius: 8px; padding: 8px 12px; margin-bottom: 12px; color: #38bdf8;">💳 <b>ID do Cartão / Crachá RFID (PaperCut):</b> <code>${{u.pager}}</code></div>`
            : `<div style="background: rgba(234, 179, 8, 0.1); border: 1px solid #ca8a04; border-radius: 8px; padding: 8px 12px; margin-bottom: 12px; color: #fbbf24;">⚠️ <b>Crachá RFID (Pager):</b> <i>Não cadastrado no Active Directory</i></div>`;

          const content = `
            ${{rfidBox}}
            <div style="margin-bottom: 12px;"><b>Status:</b> ${{statusBadge}} &nbsp;|&nbsp; <b>Login (sAMAccountName):</b> <code>${{u.sam}}</code></div>
            <div class="grid-details">
              <div class="detail-card">
                <h4>🏢 Organização & Lotação</h4>
                <div class="detail-field"><label>Cargo / Função:</label><span>${{u.title || '-'}}</span></div>
                <div class="detail-field"><label>Departamento:</label><span>${{u.dept || '-'}}</span></div>
                <div class="detail-field"><label>Escritório / Sala:</label><span>${{u.office || '-'}}</span></div>
                <div class="detail-field"><label>Gestor Imediato:</label><span>${{u.manager || '-'}}</span></div>
              </div>
              <div class="detail-card">
                <h4>📞 Contato & Segurança</h4>
                <div class="detail-field"><label>E-mail:</label><span>${{u.mail ? '<a href="mailto:' + u.mail + '" style="color:#38bdf8;">' + u.mail + '</a>' : '-'}}</span></div>
                <div class="detail-field"><label>Telefone / Ramal:</label><span>${{u.phone || '-'}}</span></div>
                <div class="detail-field"><label>UAC Flags:</label><span><code>${{u.uac || '512'}}</code></span></div>
                <div class="detail-field"><label>Último Logon:</label><span>${{formatDateBr(u.logon)}}</span></div>
              </div>
            </div>
          `;
          document.getElementById('userDetailContent').innerHTML = content;
          document.getElementById('userDetailModalBackdrop').classList.add('open');
        }}

        function closeUserDetailModal() {{
          document.getElementById('userDetailModalBackdrop').classList.remove('open');
        }}

        function openCompDetailDirect(compJsonStr) {{
          const c = typeof compJsonStr === 'string' ? JSON.parse(compJsonStr) : compJsonStr;
          document.getElementById('compDetailTitle').textContent = `💻 ${{c.name}}`;

          const targetHost = c.dns || c.name;
          const statusBadge = c.active
            ? '<span class="badge-active">🟢 MÁQUINA ATIVA</span>'
            : '<span class="badge-inactive">🔴 DESATIVADA</span>';

          const content = `
            <div style="margin-bottom: 12px;"><b>Status:</b> ${{statusBadge}} &nbsp;|&nbsp; <b>FQDN / Hostname:</b> <code>${{targetHost}}</code></div>
            <div class="grid-details">
              <div class="detail-card">
                <h4>🖥️ Hardware & Sistema Operacional</h4>
                <div class="detail-field"><label>Sistema Operacional:</label><span>${{c.os || '-'}}</span></div>
                <div class="detail-field"><label>Versão / Build:</label><span>${{c.os_ver || '-'}}</span></div>
                <div class="detail-field"><label>Descrição:</label><span>${{c.desc || '-'}}</span></div>
                <div class="detail-field"><label>Gerenciado por:</label><span>${{c.managed_by || '-'}}</span></div>
              </div>
              <div class="detail-card">
                <h4>🛡️ Domínio & Auditoria</h4>
                <div class="detail-field"><label>Ingressado em:</label><span>${{formatDateBr(c.created)}}</span></div>
                <div class="detail-field"><label>Último Logon:</label><span>${{formatDateBr(c.logon)}}</span></div>
                <div class="detail-field"><label>Controle de Conta (UAC):</label><span><code>${{c.uac || '4096'}}</code></span></div>
              </div>
            </div>

            <div style="margin-top: 14px; font-weight: 600; font-size: 13px; color: #f59e0b;">⚡ Ações Rápidas de Suporte (Protocolo Bancada)</div>
            <div class="quick-actions">
              <a class="action-link" href="bancada://run?tool=rdp&host=${{encodeURIComponent(targetHost)}}">
                <div class="action-btn rdp">
                  <span>🖥️ Conexão Remota</span>
                  <span style="font-size: 10px; opacity: 0.8;">Abrir MSTSC (RDP)</span>
                </div>
              </a>
              <a class="action-link" href="bancada://run?tool=explorer&host=${{encodeURIComponent(targetHost)}}&path=c$">
                <div class="action-btn share">
                  <span>📂 Compartilhamento</span>
                  <span style="font-size: 10px; opacity: 0.8;">\\\\\\\\${{targetHost}}\\c$</span>
                </div>
              </a>
              <a class="action-link" href="bancada://run?tool=ping&host=${{encodeURIComponent(targetHost)}}">
                <div class="action-btn ping">
                  <span>⚡ Testar Ping</span>
                  <span style="font-size: 10px; opacity: 0.8;">ICMP Contínuo</span>
                </div>
              </a>
            </div>
          `;
          document.getElementById('compDetailContent').innerHTML = content;
          document.getElementById('compDetailModalBackdrop').classList.add('open');
        }}

        function closeCompDetailModal() {{
          document.getElementById('compDetailModalBackdrop').classList.remove('open');
        }}

        // Fechar modais ao clicar no backdrop
        document.querySelectorAll('.ad-modal-backdrop').forEach(backdrop => {{
          backdrop.addEventListener('click', (e) => {{
            if (e.target === backdrop) {{
              backdrop.classList.remove('open');
            }}
          }});
        }});
      </script>
    </body>
    </html>
    """
    components.html(html_code, height=height, scrolling=False)



def format_ad_datetime(dt_val) -> str:
    """Formata timestamps do Active Directory no padrão DD/MM/AAAA HH:MM:SS."""
    if not dt_val or pd.isna(dt_val):
        return "-"
    s = str(dt_val).strip()
    if not s or s.lower() in ["none", "nan", "null", ""]:
        return "-"
    try:
        dt = pd.to_datetime(s, errors="coerce")
        if pd.notna(dt):
            return dt.strftime("%d/%m/%Y %H:%M:%S")
    except Exception:
        pass
    return s


def _explain_uac(uac_val) -> str:
    """Retorna uma interpretação amigável das flags de UserAccountControl."""
    if not uac_val or pd.isna(uac_val):
        return "Padrão (512)"
    try:
        val = int(uac_val)
        flags = []
        if val & 0x0002:
            flags.append("Conta Desativada")
        else:
            flags.append("Conta Ativa")
        if val & 0x0010:
            flags.append("Bloqueada por Tentativas")
        if val & 0x0020:
            flags.append("Não requer Senha")
        if val & 0x0040:
            flags.append("Não pode alterar senha")
        if val & 0x10000:
            flags.append("Senha Nunca Expira")
        if val & 0x800000:
            flags.append("Senha Expirada")
        desc = " | ".join(flags)
        return f"{val} ({desc})"
    except Exception:
        return str(uac_val)


@st.dialog("👤 Ficha Completa do Usuário (Active Directory)", width="large")
def show_user_details_dialog(user_row):
    """Modal nativo (@st.dialog) detalhando todos os atributos da conta de usuário no AD."""
    display_name = str(user_row.get("display_name") or user_row.get("sam_account_name") or "Usuário").strip()
    sam = str(user_row.get("sam_account_name") or "").strip()
    is_active = bool(user_row.get("is_active", True))
    status_badge = "🟢 **Conta Ativa**" if is_active else "🔴 **Conta Inativa / Desativada**"

    st.markdown(f"### 👤 {display_name}")
    st.markdown(f"**Status da Conta:** {status_badge} &nbsp;|&nbsp; **Login de Rede (sAMAccountName):** `{sam}`")

    # Destaque Pager / Crachá RFID PaperCut
    pager_val = str(user_row.get("pager") or "").strip()
    if pager_val and pager_val.lower() not in ["none", "nan", "null", ""]:
        st.info(f"💳 **ID do Cartão / Crachá RFID (PaperCut / Pager):** `{pager_val}`")
    else:
        st.warning("⚠️ **ID do Cartão / Crachá RFID (Pager):** *Não cadastrado no Active Directory*")

    st.markdown("---")

    col1, col2 = st.columns(2)

    with col1:
        st.markdown("#### 🏢 Organização & Lotação")
        st.markdown(f"**Cargo / Função:** {user_row.get('title') or '-'}")
        st.markdown(f"**Departamento:** {user_row.get('department') or '-'}")
        st.markdown(f"**Empresa:** {user_row.get('company') or 'MPMS'}")
        st.markdown(f"**Escritório / Sala:** {user_row.get('office') or '-'}")
        st.markdown(f"**Superior / Gestor Imediato:** {user_row.get('manager') or '-'}")
        if user_row.get('description'):
            st.markdown(f"**Descrição / Observações:** {user_row.get('description')}")

        st.markdown("#### 📞 Telefones & Contato")
        mail_val = str(user_row.get("mail") or "").strip()
        if mail_val and mail_val.lower() not in ["none", "nan", ""]:
            st.markdown(f"**E-mail:** [{mail_val}](mailto:{mail_val})")
        else:
            st.markdown("**E-mail:** `-`")

        st.markdown(f"**Telefone / Ramal:** `{user_row.get('telephone_number') or '-'}`")
        st.markdown(f"**Celular:** `{user_row.get('mobile') or '-'}`")

    with col2:
        st.markdown("#### 🛡️ Auditoria & Segurança da Conta")
        st.markdown(f"**📅 Data de Criação (whenCreated):** {format_ad_datetime(user_row.get('when_created'))}")
        st.markdown(f"**🕒 Último Logon (lastLogonTimestamp):** {format_ad_datetime(user_row.get('last_logon'))}")

        uac_val = user_row.get("user_account_control")
        st.markdown(f"**Controle de Conta (UAC):** `{_explain_uac(uac_val)}`")

        parent_ou = str(user_row.get("parent_ou_dn") or "-").strip()
        st.markdown(f"**Unidade Organizacional (OU):**")
        st.code(parent_ou, language="text")

        dn_val = str(user_row.get("dn") or "-").strip()
        st.markdown(f"**Distinguished Name (DN):**")
        st.code(dn_val, language="text")

    st.markdown("---")

    # Grupos dos quais o usuário é membro
    user_dn = str(user_row.get("dn") or "").strip()
    user_groups = get_user_groups(user_dn)

    with st.expander(f"🛡️ Grupos de Segurança aos quais pertence ({len(user_groups)} grupos)", expanded=True):
        if user_groups:
            df_ug = pd.DataFrame(user_groups)
            df_ug_display = df_ug[["sam_account_name", "description"]].rename(columns={
                "sam_account_name": "Nome do Grupo",
                "description": "Descrição do Grupo"
            })
            st.dataframe(df_ug_display, use_container_width=True, hide_index=True)
        else:
            st.caption("Nenhum grupo adicional associado ou sincronizado.")

    st.markdown("<br>", unsafe_allow_html=True)
    if st.button("Fechar Ficha do Usuário", key="close_user_dialog_btn", use_container_width=True):
        st.rerun()


@st.dialog("🛡️ Membros do Grupo de Segurança", width="large")
def show_group_members_dialog(group_row):
    """Modal para listar detalhadamente todos os membros de um grupo do AD."""
    grp_name = group_row.get("sam_account_name", "Grupo")
    grp_dn = group_row.get("dn", "")
    grp_desc = group_row.get("description", "")
    member_count = group_row.get("member_count", 0)

    st.markdown(f"### 🛡️ Grupo: `{grp_name}`")
    if grp_desc:
        st.info(f"**Descrição:** {grp_desc}")
    st.caption(f"**Distinguished Name (DN):** `{grp_dn}`")
    st.write(f"**Total de Membros Associados:** `{member_count}`")

    members = get_group_members(grp_dn)
    if not members:
        st.warning("Nenhum membro listado ou grupo sem membros diretos.")
        return

    df_mem = pd.DataFrame(members)
    df_mem["Status"] = df_mem["is_active"].apply(lambda a: "🟢 Ativo" if a else "🔴 Desativado")

    display_df = df_mem[["Status", "sam_account_name", "display_name", "mail", "department", "pager", "telephone_number"]].rename(columns={
        "sam_account_name": "Login de Rede",
        "display_name": "Nome Completo",
        "mail": "E-mail",
        "department": "Departamento",
        "pager": "Cartão/RFID (Pager)",
        "telephone_number": "Ramal/Telefone"
    })

    st.dataframe(display_df, use_container_width=True, hide_index=True)


@st.dialog("💻 Ficha Completa do Computador / Servidor (Active Directory)", width="large")
def show_computer_details_dialog(comp_row):
    """Modal nativo (@st.dialog) detalhando os atributos de conta de máquina no AD."""
    name = str(comp_row.get("name") or "Computador").strip()
    dns_host = str(comp_row.get("dns_hostname") or "-").strip()
    is_active = bool(comp_row.get("is_active", True))
    status_badge = "🟢 **Máquina Ativa no Domínio**" if is_active else "🔴 **Conta de Máquina Desativada**"

    st.markdown(f"### 💻 `{name}`")
    st.markdown(f"**Status:** {status_badge} &nbsp;|&nbsp; **FQDN / DNS Hostname:** `{dns_host}`")
    st.markdown("---")

    c1, c2 = st.columns(2)

    with c1:
        st.markdown("#### 🖥️ Sistema Operacional & Hardware")
        st.markdown(f"**Sistema Operacional:** `{comp_row.get('operating_system') or '-'}`")
        st.markdown(f"**Versão / Build do SO:** `{comp_row.get('os_version') or '-'}`")
        st.markdown(f"**Descrição / Função:** {comp_row.get('description') or '-'}")
        st.markdown(f"**Proprietário / Gerenciado por (managedBy):** {comp_row.get('managed_by') or '-'}")

    with c2:
        st.markdown("#### 🛡️ Auditoria & Domínio")
        st.markdown(f"**📅 Ingressado no Domínio (whenCreated):** {format_ad_datetime(comp_row.get('when_created'))}")
        st.markdown(f"**🕒 Última Atividade / Logon:** {format_ad_datetime(comp_row.get('last_logon'))}")

        uac_val = comp_row.get("user_account_control")
        st.markdown(f"**Controle de Conta (UAC):** `{_explain_uac(uac_val)}`")

        parent_ou = str(comp_row.get("parent_ou_dn") or "-").strip()
        st.markdown(f"**Unidade Organizacional (OU):**")
        st.code(parent_ou, language="text")

        dn_val = str(comp_row.get("dn") or "-").strip()
        st.markdown(f"**Distinguished Name (DN):**")
        st.code(dn_val, language="text")

    st.markdown("---")

    # Ações Rápidas de Suporte / Infraestrutura via Protocolo Bancada
    st.markdown("#### ⚡ Ações Rápidas na Estação / Servidor")
    b_col1, b_col2, b_col3 = st.columns(3)
    target_addr = dns_host if dns_host and dns_host != "-" else name

    with b_col1:
        rdp_uri = f"bancada://run?tool=rdp&host={target_addr}"
        st.markdown(
            f"""<a href="{rdp_uri}" style="text-decoration: none;">
                <div style="background: #1e293b; border: 1px solid #3b82f6; border-radius: 8px; padding: 10px; text-align: center; color: #60a5fa; font-weight: 600; font-size: 0.85rem; cursor: pointer;">
                    🖥️ Conexão Remota (RDP)
                </div>
            </a>""",
            unsafe_allow_html=True
        )
        st.caption("Abre MSTSC via disparador local")

    with b_col2:
        c_share_uri = f"bancada://run?tool=explorer&host={target_addr}&path=c$"
        st.markdown(
            f"""<a href="{c_share_uri}" style="text-decoration: none;">
                <div style="background: #1e293b; border: 1px solid #10b981; border-radius: 8px; padding: 10px; text-align: center; color: #34d399; font-weight: 600; font-size: 0.85rem; cursor: pointer;">
                    📂 Compartilhamento (C$)
                </div>
            </a>""",
            unsafe_allow_html=True
        )
        st.markdown(f'<div style="color: #94a3b8; font-size: 0.8rem; margin-top: 4px;">Abre <code>\\\\{target_addr}\\c$</code></div>', unsafe_allow_html=True)

    with b_col3:
        ping_uri = f"bancada://run?tool=ping&host={target_addr}"
        st.markdown(
            f"""<a href="{ping_uri}" style="text-decoration: none;">
                <div style="background: #1e293b; border: 1px solid #f59e0b; border-radius: 8px; padding: 10px; text-align: center; color: #fbbf24; font-weight: 600; font-size: 0.85rem; cursor: pointer;">
                    ⚡ Testar Ping (ICMP)
                </div>
            </a>""",
            unsafe_allow_html=True
        )
        st.caption("Ping contínuo no console")

    st.markdown("<br>", unsafe_allow_html=True)
    if st.button("Fechar Ficha da Máquina", key="close_comp_dialog_btn", use_container_width=True):
        st.rerun()


def render_ad_page():
    """
    Página Principal de Árvore e Consulta do Active Directory (AD / LDAP).
    Utiliza render_subtabs com persistência de URL (?tab=active-directory&subtab=slug).
    """
    st.markdown("""
        <div style="margin-bottom: 1rem;">
            <h1 style="margin: 0; font-size: 2.1rem; display: flex; align-items: center; gap: 10px;">
                🌳 Active Directory (AD / LDAP)
            </h1>
            <p style="color: #8b949e; margin: 4px 0 0 0; font-size: 0.95rem;">
                Visualização hierárquica das Unidades Organizacionais (OUs), consulta de contas de usuários e grupos de segurança corporativos via LDAPv3.
            </p>
        </div>
    """, unsafe_allow_html=True)

    # Sub-Navegação Sincronizada por URL (?subtab=arvore|usuarios|computadores|grupos|sync)
    TAB_MAP = {
        "arvore": "🌳 Árvore Hierárquica (GoJS)",
        "usuarios": "👥 Usuários & Contas",
        "computadores": "💻 Computadores & Servidores",
        "grupos": "🛡️ Grupos de Segurança",
        "sync": "⚙️ Sincronização & Diagnóstico"
    }

    selected_subtab_title = render_subtabs(TAB_MAP, default_slug="arvore", key="ad_subtabs_radio")
    current_slug = [k for k, v in TAB_MAP.items() if v == selected_subtab_title][0]

    # -------------------------------------------------------------------------
    # SUBTAB 1: ÁRVORE HIERÁRQUICA INTERATIVA (GOJS)
    # -------------------------------------------------------------------------
    if current_slug == "arvore":
        df_ous = get_ad_ous()
        if df_ous.empty:
            st.warning("⚠️ Nenhuma Unidade Organizacional (OU) encontrada no cache local.")
            st.info("💡 Acesse a aba **'⚙️ Sincronização & Diagnóstico'** e execute a primeira sincronização com o Domain Controller corporativo.")
        else:
            st.markdown(
                "Navegue interativamente pela estrutura de OUs, usuários e computadores. Utilize a barra superior para buscar nós, expandir/recolher níveis ou reposicionar a visualização."
            )
            ou_stats = get_ad_ou_stats()
            ou_entities = get_all_ou_entities_compact()
            nodes = _build_gojs_tree_data(df_ous, ou_stats=ou_stats)
            render_gojs_tree_component(nodes, ou_entities=ou_entities, height=750)

    # -------------------------------------------------------------------------
    # SUBTAB 2: USUÁRIOS & CONTAS
    # -------------------------------------------------------------------------
    elif current_slug == "usuarios":
        df_all_users = get_ad_users_df(status_filter="Todos")
        total_users = len(df_all_users)
        active_users = len(df_all_users[df_all_users["is_active"] == 1]) if not df_all_users.empty else 0
        disabled_users = total_users - active_users

        render_metric_cards([
            {
                "title": "👥 TOTAL DE CONTAS",
                "value": f"{total_users:,}".replace(",", "."),
                "border_color": "#3b82f6",
            },
            {
                "title": "✅ USUÁRIOS ATIVOS",
                "value": f"{active_users:,}".replace(",", "."),
                "border_color": "#10b981",
            },
            {
                "title": "🔒 CONTAS DESATIVADAS",
                "value": f"{disabled_users:,}".replace(",", "."),
                "border_color": "#ef4444",
            },
        ])

        with st.sidebar:
            st.markdown("### 🔍 Filtros de Usuários")
            ad_search = st.text_input(
                "Buscar por Nome, Login, E-mail ou Descrição:",
                placeholder="Ex: João, Silva, jsilva, dti...",
                key="ad_user_search"
            )
            ad_status = st.selectbox(
                "Status da Conta:",
                ["Todos", "Ativos", "Desativados"],
                index=0,
                key="ad_user_status"
            )
            ad_dept_options = ["Todos"] + get_ad_departments()
            ad_dept_selected = st.selectbox(
                "Filtrar por Departamento:",
                ad_dept_options,
                index=0,
                key="ad_user_dept"
            )

        df_users = get_ad_users_df(
            status_filter=ad_status,
            department=ad_dept_selected,
            search=ad_search
        )

        if df_users.empty:
            st.warning("Nenhum usuário encontrado no cache local para os filtros selecionados.")
            st.caption("Verifique os filtros na barra lateral ou realize uma sincronização na aba '⚙️ Sincronização & Diagnóstico'.")
        else:
            with st.sidebar:
                st.markdown("---")
                st.markdown("### 📥 Exportar Dados")
                d_col1, d_col2 = st.columns(2)
                with d_col1:
                    csv_data = df_users.to_csv(index=False).encode('utf-8-sig')
                    st.download_button(
                        "📄 CSV",
                        data=csv_data,
                        file_name="ad_usuarios.csv",
                        mime="text/csv",
                        use_container_width=True
                    )
                with d_col2:
                    try:
                        import io
                        output = io.BytesIO()
                        with pd.ExcelWriter(output, engine='openpyxl') as writer:
                            df_users.to_excel(writer, index=False, sheet_name='Usuarios_AD')
                        st.download_button(
                            "📊 Excel",
                            data=output.getvalue(),
                            file_name="ad_usuarios.xlsx",
                            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                            use_container_width=True
                        )
                    except Exception:
                        pass

            st.caption(f"Mostrando **{len(df_users)}** usuário(s) encontrado(s).")

            items_per_page = render_items_per_page_selector(
                key_prefix="ad_users",
                options=[10, 25, 50, 100, "Todos"],
                default_index=1,
                label="Usuários por página:"
            )

            slice_df, curr_page, total_p, total_i = paginate_items(
                df_users,
                page_key="ad_users_pag",
                items_per_page=items_per_page
            )

            if not slice_df.empty:
                display_slice = slice_df.copy()
                display_slice["Status"] = display_slice["is_active"].apply(
                    lambda x: "🟢 Ativo" if x == 1 else "🔴 Desativado"
                )
                display_slice["last_logon_fmt"] = display_slice["last_logon"].apply(format_ad_datetime)

                # Colunas visíveis na tabela principal
                cols_to_show = [
                    "Status",
                    "sam_account_name",
                    "display_name",
                    "title",
                    "department",
                    "mail",
                    "telephone_number",
                    "last_logon_fmt"
                ]
                display_slice_renamed = display_slice[cols_to_show].rename(columns={
                    "sam_account_name": "Login",
                    "display_name": "Nome Completo",
                    "title": "Cargo",
                    "department": "Departamento",
                    "mail": "E-mail",
                    "telephone_number": "Ramal / Tel",
                    "last_logon_fmt": "Último Logon"
                })

                if "last_selected_ad_user" not in st.session_state:
                    st.session_state["last_selected_ad_user"] = None

                user_selection = st.dataframe(
                    display_slice_renamed,
                    use_container_width=True,
                    hide_index=True,
                    on_select="rerun",
                    selection_mode="single-row",
                    key="ad_users_datagrid"
                )

                selected_rows = user_selection.selection.rows if hasattr(user_selection, "selection") else []
                if selected_rows:
                    selected_idx = selected_rows[0]
                    if st.session_state["last_selected_ad_user"] != selected_idx:
                        st.session_state["last_selected_ad_user"] = selected_idx
                        chosen_user_row = slice_df.iloc[selected_idx]
                        show_user_details_dialog(chosen_user_row)
                else:
                    st.session_state["last_selected_ad_user"] = None

                render_pagination_controls(
                    page_key="ad_users_pag",
                    current_page=curr_page,
                    total_pages=total_p,
                    total_items=total_i,
                    items_per_page=items_per_page
                )
            else:
                st.warning("Nenhum usuário corresponde aos filtros aplicados.")

    # -------------------------------------------------------------------------
    # SUBTAB 3: COMPUTADORES & SERVIDORES
    # -------------------------------------------------------------------------
    elif current_slug == "computadores":
        df_all_comps = get_ad_computers_df(status_filter="Todos")
        total_comps = len(df_all_comps)
        active_comps = len(df_all_comps[df_all_comps["is_active"] == 1]) if not df_all_comps.empty else 0
        disabled_comps = total_comps - active_comps
        
        # Filtro de servidores (por nome do sistema operacional)
        server_comps = 0
        if not df_all_comps.empty and "operating_system" in df_all_comps.columns:
            server_comps = len(df_all_comps[df_all_comps["operating_system"].str.contains("Server", case=False, na=False)])

        render_metric_cards([
            {
                "title": "💻 TOTAL DE MÁQUINAS",
                "value": f"{total_comps:,}".replace(",", "."),
                "border_color": "#3b82f6",
            },
            {
                "title": "✅ MÁQUINAS ATIVAS",
                "value": f"{active_comps:,}".replace(",", "."),
                "border_color": "#10b981",
            },
            {
                "title": "🔒 MÁQUINAS DESATIVADAS",
                "value": f"{disabled_comps:,}".replace(",", "."),
                "border_color": "#ef4444",
            },
            {
                "title": "🖥️ SERVIDORES (OS)",
                "value": f"{server_comps:,}".replace(",", "."),
                "border_color": "#8b5cf6",
            },
        ])

        with st.sidebar:
            st.markdown("### 🔍 Filtros de Computadores")
            comp_search = st.text_input(
                "Buscar por Nome, DNS, Descrição ou Responsável:",
                placeholder="Ex: SRV-1165, BANCADA, Windows...",
                key="ad_comp_search"
            )
            comp_type = st.selectbox(
                "Categoria de Máquina:",
                ["Todos", "Estações de Trabalho", "Servidores"],
                index=0,
                key="ad_comp_type"
            )
            comp_status = st.selectbox(
                "Status da Máquina:",
                ["Todos", "Ativos", "Desativados"],
                index=0,
                key="ad_comp_status"
            )
            comp_stale = st.selectbox(
                "Auditoria de Inatividade (Último Logon):",
                ["Todos", "> 30 dias", "> 90 dias", "> 180 dias", "Sem Logon Registrado"],
                index=0,
                key="ad_comp_stale"
            )
            os_options = get_ad_operating_systems()
            comp_os_selected = st.multiselect(
                "Sistemas Operacionais:",
                options=os_options,
                default=[],
                placeholder="Selecione um ou mais SOs...",
                key="ad_comp_os"
            )

        df_comps = get_ad_computers_df(
            status_filter=comp_status,
            os_filter=comp_os_selected,
            search=comp_search,
            machine_type=comp_type,
            stale_days=comp_stale
        )

        # Painel retrátil de distribuição de Sistemas Operacionais
        if not df_all_comps.empty and "operating_system" in df_all_comps.columns:
            with st.expander("📊 Distribuição de Sistemas Operacionais no Parque (Active Directory)", expanded=False):
                os_counts = df_all_comps["operating_system"].replace("", "Não Especificado").value_counts().reset_index()
                os_counts.columns = ["Sistema Operacional", "Quantidade"]
                total_m = len(df_all_comps)
                os_counts["Percentual"] = os_counts["Quantidade"].apply(lambda q: f"{(q / total_m) * 100:.1f}%")
                st.dataframe(os_counts, use_container_width=True, hide_index=True)

        if df_comps.empty:
            st.warning("Nenhum computador ou servidor encontrado no cache local para os filtros informados.")
            st.caption("Ajuste os filtros na barra lateral ou execute uma nova sincronização na aba '⚙️ Sincronização & Diagnóstico'.")
        else:
            with st.sidebar:
                st.markdown("---")
                st.markdown("### 📥 Exportar Computadores")
                dc1, dc2 = st.columns(2)
                with dc1:
                    csv_data = df_comps.to_csv(index=False).encode('utf-8-sig')
                    st.download_button(
                        "📄 CSV",
                        data=csv_data,
                        file_name="ad_computadores.csv",
                        mime="text/csv",
                        use_container_width=True
                    )
                with dc2:
                    try:
                        import io
                        output = io.BytesIO()
                        with pd.ExcelWriter(output, engine='openpyxl') as writer:
                            df_comps.to_excel(writer, index=False, sheet_name='Computadores_AD')
                        st.download_button(
                            "📊 Excel",
                            data=output.getvalue(),
                            file_name="ad_computadores.xlsx",
                            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                            use_container_width=True
                        )
                    except Exception:
                        pass

            st.caption(f"Mostrando **{len(df_comps)}** máquina(s) localizada(s). Marque a caixa ao lado de uma linha para abrir a ficha completa.")

            c_items_per_page = render_items_per_page_selector(
                key_prefix="ad_comps",
                options=[10, 25, 50, 100, "Todos"],
                default_index=1,
                label="Máquinas por página:"
            )

            c_slice, c_curr, c_total_p, c_total_i = paginate_items(
                df_comps,
                page_key="ad_comps_pag",
                items_per_page=c_items_per_page
            )

            if not c_slice.empty:
                display_c_slice = c_slice.copy()
                display_c_slice["Status"] = display_c_slice["is_active"].apply(
                    lambda x: "🟢 Ativo" if x == 1 else "🔴 Desativado"
                )
                display_c_slice["last_logon_fmt"] = display_c_slice["last_logon"].apply(format_ad_datetime)

                # Colunas visíveis na tabela principal
                cols_to_show = [
                    "Status",
                    "name",
                    "dns_hostname",
                    "operating_system",
                    "os_version",
                    "description",
                    "managed_by",
                    "last_logon_fmt"
                ]
                display_c_slice_renamed = display_c_slice[cols_to_show].rename(columns={
                    "name": "Nome da Máquina",
                    "dns_hostname": "FQDN / DNS Hostname",
                    "operating_system": "Sistema Operacional",
                    "os_version": "Versão / Build",
                    "description": "Descrição",
                    "managed_by": "Responsável",
                    "last_logon_fmt": "Última Atividade"
                })

                if "last_selected_ad_comp" not in st.session_state:
                    st.session_state["last_selected_ad_comp"] = None

                comp_selection = st.dataframe(
                    display_c_slice_renamed,
                    use_container_width=True,
                    hide_index=True,
                    on_select="rerun",
                    selection_mode="single-row",
                    key="ad_comps_datagrid"
                )

                selected_rows = comp_selection.selection.rows if hasattr(comp_selection, "selection") else []
                if selected_rows:
                    selected_idx = selected_rows[0]
                    if st.session_state["last_selected_ad_comp"] != selected_idx:
                        st.session_state["last_selected_ad_comp"] = selected_idx
                        chosen_comp_row = c_slice.iloc[selected_idx]
                        show_computer_details_dialog(chosen_comp_row)
                else:
                    st.session_state["last_selected_ad_comp"] = None

                render_pagination_controls(
                    page_key="ad_comps_pag",
                    current_page=c_curr,
                    total_pages=c_total_p,
                    total_items=c_total_i,
                    items_per_page=c_items_per_page
                )
            else:
                st.warning("Nenhuma máquina corresponde aos filtros aplicados.")

    # -------------------------------------------------------------------------
    # SUBTAB 4: GRUPOS DE SEGURANÇA
    # -------------------------------------------------------------------------
    elif current_slug == "grupos":
        with st.sidebar:
            st.markdown("### 🔍 Filtros de Grupos")
            g_search = st.text_input(
                "Buscar por Nome ou Descrição:",
                placeholder="Ex: STI, Gestores, DTI...",
                key="ad_group_search"
            )

        df_groups = get_ad_groups_df(search=g_search)

        if df_groups.empty:
            st.warning("Nenhum grupo de segurança encontrado para a busca informada.")
        else:
            st.caption(f"Total de grupos localizados: **{len(df_groups)}**")

            g_per_page = render_items_per_page_selector(
                key_prefix="ad_groups",
                options=[10, 25, 50, "Todos"],
                default_index=0,
                label="Grupos por página:"
            )

            g_slice, g_curr, g_total_p, g_total_i = paginate_items(
                df_groups,
                page_key="ad_groups_pag",
                items_per_page=g_per_page
            )

            for _, g_row in g_slice.iterrows():
                with st.container():
                    c1, c2, c3 = st.columns([3, 1, 1.2])
                    with c1:
                        st.markdown(f"🛡️ **{g_row['sam_account_name']}**")
                        if g_row.get('description'):
                            st.caption(f"{g_row['description']}")
                    with c2:
                        st.markdown(f"👥 **{g_row['member_count']}** membros")
                    with c3:
                        if st.button("Ver Membros 🔍", key=f"btn_grp_{g_row['sam_account_name']}", use_container_width=True):
                            show_group_members_dialog(g_row)
                    st.divider()

            render_pagination_controls(
                page_key="ad_groups_pag",
                current_page=g_curr,
                total_pages=g_total_p,
                total_items=g_total_i,
                items_per_page=g_per_page
            )

    # -------------------------------------------------------------------------
    # SUBTAB 5: SINCRONIZAÇÃO & DIAGNÓSTICO
    # -------------------------------------------------------------------------
    elif current_slug == "sync":
        st.markdown("### ⚙️ Status da Conexão e Banco de Dados")

        meta = get_ad_sync_meta()
        last_sync_fmt = format_ad_datetime(meta.get("last_sync_at")) if meta.get("last_sync_at") else "Nunca executado"

        render_metric_cards([
            {
                "title": "🕒 ÚLTIMA SINCRONIZAÇÃO",
                "value": last_sync_fmt,
                "border_color": "#3b82f6",
            },
            {
                "title": "📁 OUS EM CACHE",
                "value": f"{meta.get('total_ous', 0):,}".replace(",", "."),
                "border_color": "#10b981",
            },
            {
                "title": "👥 USUÁRIOS EM CACHE",
                "value": f"{meta.get('total_users', 0):,}".replace(",", "."),
                "border_color": "#8b5cf6",
            },
            {
                "title": "💻 COMPUTADORES EM CACHE",
                "value": f"{meta.get('total_computers', 0):,}".replace(",", "."),
                "border_color": "#06b6d4",
            },
            {
                "title": "🛡️ GRUPOS EM CACHE",
                "value": f"{meta.get('total_groups', 0):,}".replace(",", "."),
                "border_color": "#f59e0b",
            },
        ])

        if meta.get("status") == "error":
            st.error(f"⚠️ Erro registrado na última sincronização: {meta.get('error_message')}")

        st.markdown("---")

        d_col1, d_col2 = st.columns(2)

        with d_col1:
            st.markdown("#### 🔌 Diagnóstico do Domain Controller")
            if st.button("Testar Conexão LDAP ⚡", use_container_width=True):
                with st.spinner("Testando conexão e autenticação com o DC..."):
                    diag = test_ad_connection()
                    if diag.get("success"):
                        st.success(f"✅ Conexão estabelecida com sucesso em **{diag['latency_seconds']}s**!")
                        st.write(f"- **Servidor Host:** `{diag['server_host']}:{diag['server_port']}`")
                        st.write(f"- **Domínio:** `{diag['domain']}`")
                        st.write(f"- **Base DN:** `{diag['base_dn']}`")
                        st.write(f"- **Usuário Autenticado:** `{diag['user']}`")
                    else:
                        st.error(f"❌ Falha de Conexão: {diag.get('error')}")

        with d_col2:
            st.markdown("#### 🔄 Atualizar Cache Local")
            st.caption("Efetua a leitura completa da hierarquia de OUs, contas de usuários, computadores e grupos e atualiza o banco relacional.")
            if st.button("Sincronizar Active Directory Agora 🚀", type="primary", use_container_width=True):
                prog_bar = st.progress(0, text="Iniciando sincronização...")

                def _update_progress(pct: int, msg: str):
                    prog_bar.progress(pct, text=msg)

                res = sync_active_directory_cache(page_size=500, progress_callback=_update_progress)

                if res.get("success"):
                    st.success(f"🎉 Sincronização concluída com sucesso! ({res.get('total_ous', 0)} OUs, {res.get('total_users', 0)} Usuários, {res.get('total_computers', 0)} Computadores, {res.get('total_groups', 0)} Grupos)")
                    st.rerun()
                else:
                    st.error(f"❌ Erro na sincronização: {res.get('error')}")
