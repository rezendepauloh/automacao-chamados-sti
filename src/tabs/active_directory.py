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
    get_ad_sync_meta
)
from src.services.ad_ldap_service import test_ad_connection, sync_active_directory_cache
from src.components.subtabs import render_subtabs
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


def _build_gojs_tree_data(df_ous: pd.DataFrame) -> list:
    """
    Transforma o DataFrame de OUs no formato de nós para o GoJS TreeModel.
    Injeta um nó Raiz centralizador ("ROOT_DOMAIN") para conectar todas as OUs de topo,
    garantindo que o GoJS renderize e expanda a árvore com 100% de consistência.
    """
    nodes = []
    if df_ous.empty:
        return nodes

    root_key = "ROOT_DOMAIN"
    domain_name = DOMINIO or "in.mpe.ms.gov.br"
    
    # Adiciona nó raiz mestre do domínio
    nodes.append({
        "key": root_key,
        "name": f"🏛️ Domínio: {domain_name}",
        "parent": None,
        "description": "Floresta e Raiz Corporativa do Active Directory",
        "isRoot": True
    })

    dn_set = set(df_ous["dn"].dropna().tolist())

    for _, row in df_ous.iterrows():
        dn = str(row["dn"])
        name = str(row["name"])
        parent = str(row["parent_dn"]) if pd.notna(row["parent_dn"]) else ""
        desc = str(row.get("description", "")) if pd.notna(row.get("description")) else ""

        # Se o parent_dn for outra OU, conecta a ela; caso contrário, conecta à raiz corporativa
        parent_key = parent if (parent and parent in dn_set) else root_key

        nodes.append({
            "key": dn,
            "name": name,
            "parent": parent_key,
            "description": desc,
            "isRoot": False
        })

    return nodes


def render_gojs_tree_component(nodes: list, height: int = 700):
    """
    Renderiza um diagrama GoJS interativo no Streamlit com busca, zoom e expansão/recolhimento de nós.
    """
    json_data = json.dumps(nodes)

    html_code = f"""
    <!DOCTYPE html>
    <html lang="pt-BR">
    <head>
      <meta charset="UTF-8">
      <script src="https://unpkg.com/gojs@2.3.17/release/go.js"></script>
      <style>
        * {{
          box-sizing: border-box;
          font-family: -apple-system, BlinkMacSystemFont, "Segoe UI", Roboto, Helvetica, Arial, sans-serif;
        }}
        html, body {{
          margin: 0;
          padding: 0;
          width: 100%;
          height: 100%;
          background-color: #0e1117;
          color: #fafafa;
          overflow: hidden;
        }}
        #toolbar {{
          display: flex;
          align-items: center;
          gap: 8px;
          padding: 10px 14px;
          background: #161b22;
          border-bottom: 1px solid #30363d;
          height: 52px;
        }}
        #searchBox {{
          flex: 1;
          max-width: 320px;
          padding: 7px 12px;
          border-radius: 6px;
          border: 1px solid #30363d;
          background: #0d1117;
          color: #ffffff;
          font-size: 13px;
          outline: none;
        }}
        #searchBox:focus {{
          border-color: #58a6ff;
        }}
        .btn {{
          background: #21262d;
          color: #c9d1d9;
          border: 1px solid #30363d;
          padding: 6px 12px;
          border-radius: 6px;
          cursor: pointer;
          font-size: 12px;
          transition: all 0.2s;
        }}
        .btn:hover {{
          background: #30363d;
          border-color: #8b949e;
          color: #ffffff;
        }}
        #infoBadge {{
          margin-left: auto;
          font-size: 12px;
          color: #8b949e;
        }}
        #myDiagramDiv {{
          width: 100%;
          height: calc(100% - 52px);
          background-color: #0e1117;
        }}
      </style>
    </head>
    <body>
      <div id="toolbar">
        <input type="text" id="searchBox" placeholder="🔍 Buscar Unidade / Departamento..." onkeyup="searchNode()" />
        <button class="btn" onclick="expandAll()">➕ Expandir Tudo</button>
        <button class="btn" onclick="collapseAll()">➖ Recolher</button>
        <button class="btn" onclick="zoomFit()">🎯 Centralizar</button>
        <div id="infoBadge">Total: <b id="nodeCount">{len(nodes) - 1}</b> OUs</div>
      </div>
      <div id="myDiagramDiv"></div>

      <script>
        const $ = go.GraphObject.make;
        const myDiagram = $(go.Diagram, "myDiagramDiv", {{
          "undoManager.isEnabled": false,
          layout: $(go.TreeLayout, {{
            angle: 90,
            layerSpacing: 40,
            nodeSpacing: 18,
            alignment: go.TreeLayout.AlignmentCenterChildren
          }}),
          initialAutoScale: go.Diagram.Uniform,
          "animationManager.isEnabled": true
        }});

        myDiagram.nodeTemplate = $(go.Node, "Auto",
          {{
            isTreeExpanded: false,
            cursor: "pointer",
            selectionAdorned: true
          }},
          $(go.Shape, "RoundedRectangle", {{
            fill: "#161b22",
            stroke: "#30363d",
            strokeWidth: 1.5,
            parameter1: 6
          }},
          new go.Binding("fill", "isRoot", isR => isR ? "#1f3a5f" : "#161b22"),
          new go.Binding("stroke", "isRoot", isR => isR ? "#58a6ff" : "#30363d"),
          new go.Binding("stroke", "isHighlighted", h => h ? "#f59e0b" : "#30363d").ofObject(),
          new go.Binding("strokeWidth", "isHighlighted", h => h ? 3 : 1.5).ofObject(),
          new go.Binding("fill", "isHighlighted", h => h ? "#3d2b0f" : "#161b22").ofObject()
          ),
          $(go.Panel, "Vertical", {{ margin: 8 }},
            $(go.Panel, "Horizontal",
              $(go.TextBlock, {{ font: "14px sans-serif", margin: new go.Margin(0, 6, 0, 0) }},
                new go.Binding("text", "isRoot", isR => isR ? "🏛️" : "📁")),
              $(go.TextBlock, {{
                font: "bold 13px -apple-system, BlinkMacSystemFont, 'Segoe UI', Roboto, sans-serif",
                stroke: "#f0f6fc",
                maxSize: new go.Size(220, NaN),
                wrap: go.TextBlock.WrapFit
              }},
              new go.Binding("text", "name"))
            ),
            $("TreeExpanderButton", {{
              alignment: go.Spot.Bottom,
              alignmentFocus: go.Spot.Top,
              margin: new go.Margin(6, 0, 0, 0)
            }})
          )
        );

        myDiagram.linkTemplate = $(go.Link, {{
          routing: go.Link.Orthogonal,
          corner: 5,
          selectable: false
        }},
          $(go.Shape, {{ strokeWidth: 1.5, stroke: "#3b4252" }})
        );

        const nodeDataArray = {json_data};
        myDiagram.model = new go.TreeModel(nodeDataArray);

        // Expande o nó raiz e o primeiro nível ao carregar
        myDiagram.findTreeRoots().each(r => {{
          r.isTreeExpanded = true;
          r.findTreeChildrenNodes().each(child => child.isTreeExpanded = false);
        }});

        setTimeout(() => {{
          myDiagram.requestUpdate();
          myDiagram.commandHandler.zoomToFit();
        }}, 200);

        function expandAll() {{
          myDiagram.startTransaction("expandAll");
          myDiagram.nodes.each(n => n.isTreeExpanded = true);
          myDiagram.commitTransaction("expandAll");
          myDiagram.commandHandler.zoomToFit();
        }}

        function collapseAll() {{
          myDiagram.startTransaction("collapseAll");
          myDiagram.findTreeRoots().each(r => {{
            r.collapseTree();
            r.isTreeExpanded = true;
          }});
          myDiagram.commitTransaction("collapseAll");
          myDiagram.commandHandler.zoomToFit();
        }}

        function zoomFit() {{
          myDiagram.commandHandler.zoomToFit();
        }}

        function searchNode() {{
          const input = document.getElementById("searchBox").value.trim().toLowerCase();
          myDiagram.startTransaction("highlight");
          if (!input) {{
            myDiagram.clearHighlighteds();
            myDiagram.commitTransaction("highlight");
            return;
          }}
          myDiagram.clearHighlighteds();
          let firstMatch = null;
          myDiagram.nodes.each(n => {{
            const name = (n.data.name || "").toLowerCase();
            if (name.includes(input)) {{
              n.isHighlighted = true;
              let curr = n.findTreeParentNode();
              while (curr) {{
                curr.isTreeExpanded = true;
                curr = curr.findTreeParentNode();
              }}
              if (!firstMatch) firstMatch = n;
            }}
          }});
          if (firstMatch) {{
            myDiagram.centerRect(firstMatch.actualBounds);
          }}
          myDiagram.commitTransaction("highlight");
        }}
      </script>
    </body>
    </html>
    """
    components.html(html_code, height=height, scrolling=False)


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

    display_df = df_mem[["Status", "sam_account_name", "display_name", "mail", "department"]].rename(columns={
        "sam_account_name": "Login de Rede",
        "display_name": "Nome Completo",
        "mail": "E-mail",
        "department": "Departamento"
    })

    st.dataframe(display_df, use_container_width=True, hide_index=True)


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

    # Sub-Navegação Sincronizada por URL (?subtab=arvore|usuarios|grupos|sync)
    TAB_MAP = {
        "arvore": "🌳 Árvore Hierárquica (GoJS)",
        "usuarios": "👥 Usuários & Contas",
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
                "Navegue interativamente pela estrutura de OUs e departamentos. Utilize a barra superior para buscar nós, expandir/recolher níveis ou reposicionar a visualização."
            )
            nodes = _build_gojs_tree_data(df_ous)
            render_gojs_tree_component(nodes, height=720)

    # -------------------------------------------------------------------------
    # SUBTAB 2: USUÁRIOS & CONTAS
    # -------------------------------------------------------------------------
    elif current_slug == "usuarios":
        df_all_users = get_ad_users_df(status_filter="Todos")
        total_users = len(df_all_users)

        if total_users == 0:
            st.warning("⚠️ Nenhuma conta de usuário sincronizada no banco de dados.")
            st.info("Execute a sincronização na aba **'⚙️ Sincronização & Diagnóstico'**.")
        else:
            total_active = int(df_all_users["is_active"].sum()) if not df_all_users.empty else 0
            total_inactive = total_users - total_active

            m_col1, m_col2, m_col3 = st.columns(3)
            m_col1.metric("👥 Total de Usuários", f"{total_users:,}".replace(",", "."))
            m_col2.metric("🟢 Usuários Ativos", f"{total_active:,}".replace(",", "."))
            m_col3.metric("🔴 Bloqueados / Desativados", f"{total_inactive:,}".replace(",", "."))

            st.markdown("---")

            # Filtros
            f_col1, f_col2, f_col3 = st.columns([2, 1.2, 1.2])
            with f_col1:
                search_query = st.text_input(
                    "🔍 Buscar por Nome, Login ou E-mail:",
                    placeholder="Digite para filtrar...",
                    key="ad_users_search"
                )
            with f_col2:
                status_opt = st.selectbox(
                    "Status da Conta:",
                    ["Todos", "Ativos", "Desativados"],
                    key="ad_users_status"
                )
            with f_col3:
                dept_list = ["Todos"] + get_ad_departments()
                dept_opt = st.selectbox(
                    "Departamento:",
                    dept_list,
                    key="ad_users_dept"
                )

            # Consulta filtrada
            df_filtered = get_ad_users_df(
                status_filter=status_opt,
                department=dept_opt,
                search=search_query
            )

            # Ações de Exportação Segura (Sanitizada contra IllegalCharacterError)
            action_col1, action_col2 = st.columns([3, 1])
            with action_col2:
                if not df_filtered.empty:
                    try:
                        excel_df = _sanitize_df_for_excel(df_filtered)
                        excel_buffer = io.BytesIO()
                        with pd.ExcelWriter(excel_buffer, engine="openpyxl") as writer:
                            excel_df.to_excel(writer, index=False, sheet_name="Usuarios_AD")
                        st.download_button(
                            label="📥 Exportar Excel",
                            data=excel_buffer.getvalue(),
                            file_name=f"usuarios_ad_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx",
                            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                            use_container_width=True
                        )
                    except Exception as e:
                        # Fallback transparente para CSV caso o openpyxl falhe
                        csv_data = df_filtered.to_csv(index=False, sep=";").encode("utf-8-sig")
                        st.download_button(
                            label="📥 Exportar CSV",
                            data=csv_data,
                            file_name=f"usuarios_ad_{datetime.now().strftime('%Y%m%d_%H%M%S')}.csv",
                            mime="text/csv",
                            use_container_width=True
                        )

            # Paginação de Usuários
            items_per_page = render_items_per_page_selector(
                key_prefix="ad_users",
                options=[15, 30, 50, 100, "Todos"],
                default_index=0,
                label="Exibir por página:"
            )

            slice_df, curr_page, total_p, total_i = paginate_items(
                df_filtered,
                page_key="ad_users_pag",
                items_per_page=items_per_page
            )

            if not slice_df.empty:
                display_slice = slice_df.copy()
                display_slice["Status"] = display_slice["is_active"].apply(lambda a: "🟢 Ativo" if a else "🔴 Inativo")
                display_slice = display_slice[[
                    "Status", "sam_account_name", "display_name", "mail", "department", "title", "last_logon"
                ]].rename(columns={
                    "sam_account_name": "Login (sAMAccountName)",
                    "display_name": "Nome Completo",
                    "mail": "E-mail",
                    "department": "Departamento",
                    "title": "Cargo",
                    "last_logon": "Último Logon"
                })

                st.dataframe(display_slice, use_container_width=True, hide_index=True)
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
    # SUBTAB 3: GRUPOS DE SEGURANÇA
    # -------------------------------------------------------------------------
    elif current_slug == "grupos":
        st.markdown("Consulta e auditoria de Grupos de Segurança do Active Directory e seus respectivos membros vinculados.")

        g_search = st.text_input("🔍 Buscar Grupo por Nome ou Descrição:", placeholder="Ex: STI, Gestores, DTI...", key="ad_group_search")
        df_groups = get_ad_groups_df(search=g_search)

        if df_groups.empty:
            st.warning("Nenhum grupo de segurança encontrado.")
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
    # SUBTAB 4: SINCRONIZAÇÃO & DIAGNÓSTICO
    # -------------------------------------------------------------------------
    elif current_slug == "sync":
        st.markdown("### ⚙️ Status da Conexão e Banco de Dados")

        meta = get_ad_sync_meta()

        col_st1, col_st2, col_st3, col_st4 = st.columns(4)
        col_st1.metric("Última Sincronização", meta.get("last_sync_at") or "Nunca executado")
        col_st2.metric("OUs em Cache", f"{meta.get('total_ous', 0):,}".replace(",", "."))
        col_st3.metric("Usuários em Cache", f"{meta.get('total_users', 0):,}".replace(",", "."))
        col_st4.metric("Grupos em Cache", f"{meta.get('total_groups', 0):,}".replace(",", "."))

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
            st.caption("Efetua a leitura completa da hierarquia de OUs, contas de usuários e grupos e atualiza o banco relacional.")
            if st.button("Sincronizar Active Directory Agora 🚀", type="primary", use_container_width=True):
                prog_bar = st.progress(0, text="Iniciando sincronização...")

                def _update_progress(pct: int, msg: str):
                    prog_bar.progress(pct, text=msg)

                res = sync_active_directory_cache(page_size=500, progress_callback=_update_progress)

                if res.get("success"):
                    st.success(f"🎉 Sincronização concluída com sucesso! ({res['total_ous']} OUs, {res['total_users']} Usuários, {res['total_groups']} Grupos)")
                    st.rerun()
                else:
                    st.error(f"❌ Erro na sincronização: {res.get('error')}")
