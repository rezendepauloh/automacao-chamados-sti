import sys
import io
import pandas as pd
import streamlit as st
from datetime import datetime
from pathlib import Path

root_dir = Path(__file__).parent.parent.parent
sys.path.insert(0, str(root_dir))

from src.database import get_central_telefonica_df
from src.scrapers.oxe_scraper import check_oxe_sync_running, read_oxe_last_log_lines
from src.components.status_banner import render_log_expander
from src.components.pagination import (
    render_items_per_page_selector,
    paginate_items,
    render_pagination_controls
)


def get_val(row, *keys, default="-"):
    """Busca insensível a maiúsculas/minúsculas e formato snake_case em dicionário/Series."""
    row_dict = {str(k).lower().replace(" ", "_").replace("/", "_").replace(".", "").replace("-", "_"): v for k, v in row.items()}
    for k in keys:
        norm_k = str(k).lower().replace(" ", "_").replace("/", "_").replace(".", "").replace("-", "_")
        if norm_k in row_dict:
            val = str(row_dict[norm_k]).strip()
            if val and val.lower() not in ["none", "nan", "null", "255"]:
                return val
    return default


def render_central_telefonica_page():
    """
    Renderiza a página principal de gestão e consulta da Central Telefônica (OXE).
    """
    st.markdown("""
        <style>
            .badge-ip {

                background-color: #065f46;
                color: #34d399;
                padding: 2px 8px;
                border-radius: 4px;
                font-weight: bold;
                font-size: 0.8rem;
            }
            .badge-analog {
                background-color: #7c2d12;
                color: #fb923c;
                padding: 2px 8px;
                border-radius: 4px;
                font-weight: bold;
                font-size: 0.8rem;
            }
        </style>
    """, unsafe_allow_html=True)

    st.title("📞 Central Telefônica (OXE)")
    st.caption("Consulta unificada de ramais, utilizadores, endereços IP e MAC Addresses de hardware.")

    oxe_ativo = check_oxe_sync_running()

    render_log_expander(
        "🤖 Robô do OXE Rodando em Segundo Plano – Acompanhar Progresso",
        oxe_ativo,
        read_oxe_last_log_lines,
        check_oxe_sync_running,
        "O robô está conectando à central Alcatel e pré-processando os dados neste momento. O painel permanece livre para uso!"
    )

    # Carrega dados do banco de dados SQLite / Tratados
    df = get_central_telefonica_df()

    # -----------------------------------------------------------------------------
    # FILTROS LATERAIS (SIDEBAR) & AÇÕES DE COLETA
    # -----------------------------------------------------------------------------
    with st.sidebar:
        st.markdown("## ⚙️ Ações e Coleta")
        if oxe_ativo:
            st.button("🤖 Sincronizando OXE...", width='stretch', disabled=True)
        else:
            if st.button("🔄 Sincronizar Ramais (OXE)", type="primary", width='stretch', help="Executa o scraper e o pré-processamento em segundo plano."):
                import subprocess, time
                popen_kwargs = {"creationflags": subprocess.CREATE_NO_WINDOW} if sys.platform == "win32" else {}
                subprocess.Popen([sys.executable, "src/scrapers/oxe_scraper.py"], **popen_kwargs)
                time.sleep(1.0)
                st.toast("🚀 Scraper do OXE iniciado em segundo plano!", icon="🤖")
                st.rerun()

        st.markdown("---")
        st.header("🔍 Filtros de Busca")

    total_ramais = len(df)
    
    def is_valid_series(series):
        return ~series.astype(str).str.strip().str.lower().isin(["", "-", "none", "nan", "null", "255"])

    col_mac = 'mac_address' if 'mac_address' in df.columns else None
    col_ip = 'ip_address' if 'ip_address' in df.columns else ('endereco_ip' if 'endereco_ip' in df.columns else None)
    col_tipo = 'tipo_estacao' if 'tipo_estacao' in df.columns else ('tipo_de_estacao' if 'tipo_de_estacao' in df.columns else None)
    col_grupo = 'grupo_captura' if 'grupo_captura' in df.columns else ('pickup_group_name' if 'pickup_group_name' in df.columns else None)
    col_cat_pub = 'cat_rede_publica' if 'cat_rede_publica' in df.columns else ('public_network_category_id' if 'public_network_category_id' in df.columns else None)
    col_role = 'funcao_role' if 'funcao_role' in df.columns else ('set_role' if 'set_role' in df.columns else None)

    ramais_com_mac = len(df[is_valid_series(df[col_mac])]) if col_mac and col_mac in df.columns else 0
    ramais_com_ip = len(df[is_valid_series(df[col_ip])]) if col_ip and col_ip in df.columns else 0
    ramais_analogicos = len(df[df[col_tipo].astype(str).str.upper().str.contains('ANALOG', na=False)]) if col_tipo and col_tipo in df.columns else 0

    c1, c2, c3, c4 = st.columns(4)

    with c1:
        st.markdown(f"""
            <div class="metric-card-oxe">
                <div class="metric-title-oxe">TOTAL DE RAMAIS</div>
                <div class="metric-value-oxe">{total_ramais}</div>
            </div>
        """, unsafe_allow_html=True)

    with c2:
        st.markdown(f"""
            <div class="metric-card-oxe" style="border-left-color: #10b981;">
                <div class="metric-title-oxe">TELEFONES IP COM MAC</div>
                <div class="metric-value-oxe">{ramais_com_mac}</div>
            </div>
        """, unsafe_allow_html=True)

    with c3:
        st.markdown(f"""
            <div class="metric-card-oxe" style="border-left-color: #3b82f6;">
                <div class="metric-title-oxe">RAMAIS COM ENDEREÇO IP</div>
                <div class="metric-value-oxe">{ramais_com_ip}</div>
            </div>
        """, unsafe_allow_html=True)

    with c4:
        st.markdown(f"""
            <div class="metric-card-oxe" style="border-left-color: #f59e0b;">
                <div class="metric-title-oxe">RAMAIS ANALÓGICOS / OUTROS</div>
                <div class="metric-value-oxe">{ramais_analogicos}</div>
            </div>
        """, unsafe_allow_html=True)

    st.markdown("---")

    # -----------------------------------------------------------------------------
    # FILTROS LATERAIS (SIDEBAR)
    # -----------------------------------------------------------------------------
    with st.sidebar:
        search_query = st.text_input(
            "Buscar por Ramal, Nome, IP, MAC ou Login:",
            value="",
            placeholder="Ex: 2153, Jean, 10.111..., 48:7A..."
        ).strip().lower()


        col_cat = 'categoria_dispositivo' if 'categoria_dispositivo' in df.columns else None
        categorias_disponiveis = sorted(df[col_cat].dropna().unique().tolist()) if col_cat else []
        
        sel_categorias = st.multiselect(
            "Categoria do Dispositivo:",
            options=categorias_disponiveis,
            default=[]
        )

        tipos_disponiveis = sorted(df[col_tipo].dropna().unique().tolist()) if col_tipo else []
        
        sel_tipos = st.multiselect(
            "Tipo de Estação (OXE):",
            options=tipos_disponiveis,
            default=[]
        )

        grupos_disponiveis = sorted(df[col_grupo].dropna().astype(str).unique().tolist()) if col_grupo else []
        sel_grupos = st.multiselect(
            "👥 Grupo de Captura:",
            options=[g for g in grupos_disponiveis if g.strip() and g.strip() not in ["-", "None", "nan"]],
            default=[]
        )

        cat_pub_disponiveis = sorted(df[col_cat_pub].dropna().astype(str).unique().tolist()) if col_cat_pub else []
        sel_cat_pub = st.multiselect(
            "🌐 Categoria Rede Pública:",
            options=[c for c in cat_pub_disponiveis if c.strip() and c.strip() not in ["-", "None", "nan"]],
            default=[]
        )

        roles_disponiveis = sorted(df[col_role].dropna().astype(str).unique().tolist()) if col_role else []
        sel_roles = st.multiselect(
            "🎭 Função / Role:",
            options=[r for r in roles_disponiveis if r.strip() and r.strip() not in ["-", "None", "nan"]],
            default=[]
        )

        only_with_mac = st.checkbox("Exibir apenas ramais com MAC Address", value=False)
        only_with_ip = st.checkbox("Exibir apenas ramais com Endereço IP", value=False)
        only_without_mac = st.checkbox("Exibir apenas ramais sem MAC Address", value=False)
        only_without_ip = st.checkbox("Exibir apenas ramais sem Endereço IP", value=False)

        items_per_page = render_items_per_page_selector(
            key_prefix="central_oxe",
            options=[10, 25, 50, 100, 200, 500, "Todos"],
            default_index=2,
            label="📄 Ramais por página:"
        )


    # -----------------------------------------------------------------------------
    # APLICAÇÃO DOS FILTROS
    # -----------------------------------------------------------------------------
    df_filtered = df.copy()

    if search_query:
        cols_to_search = [c for c in df_filtered.columns if c not in ['id', 'data_atualizacao']]
        mask = pd.Series(False, index=df_filtered.index)
        for col in cols_to_search:
            mask |= df_filtered[col].astype(str).str.lower().str.contains(search_query, na=False)
        df_filtered = df_filtered[mask]

    if sel_categorias and col_cat:
        df_filtered = df_filtered[df_filtered[col_cat].isin(sel_categorias)]

    if sel_tipos and col_tipo:
        df_filtered = df_filtered[df_filtered[col_tipo].isin(sel_tipos)]

    if sel_grupos and col_grupo:
        df_filtered = df_filtered[df_filtered[col_grupo].astype(str).isin(sel_grupos)]

    if sel_cat_pub and col_cat_pub:
        df_filtered = df_filtered[df_filtered[col_cat_pub].astype(str).isin(sel_cat_pub)]

    if sel_roles and col_role:
        df_filtered = df_filtered[df_filtered[col_role].astype(str).isin(sel_roles)]

    if only_with_mac and col_mac:
        df_filtered = df_filtered[is_valid_series(df_filtered[col_mac])]

    if only_with_ip and col_ip:
        df_filtered = df_filtered[is_valid_series(df_filtered[col_ip])]

    if only_without_mac and col_mac:
        df_filtered = df_filtered[~is_valid_series(df_filtered[col_mac])]

    if only_without_ip and col_ip:
        df_filtered = df_filtered[~is_valid_series(df_filtered[col_ip])]

    # -----------------------------------------------------------------------------
    # EXIBIÇÃO DA TABELA & MODAL DE DETALHES (@st.dialog)
    # -----------------------------------------------------------------------------
    st.subheader(f"📋 Lista de Ramais ({len(df_filtered)} de {len(df)})")
    st.caption("💡 Clique na **caixa de seleção (checkbox)** no início de qualquer linha para abrir a Ficha Técnica Completa do Ramal.")

    @st.dialog("📞 Detalhes Completos do Ramal", width="large")
    def show_ramal_details(row):
        ramal_num = get_val(row, "ramal", "directory_number")
        nome = get_val(row, "nome_exibido", "nome_titular", "annu_name", "display_name")
        
        st.markdown(f"### 📞 Ramal **{ramal_num}** — {nome}")
        st.caption("Ficha técnica completa extraída da Central Telefônica Alcatel-Lucent OmniPCX Enterprise (OXE).")

        # 1. Identificação & Titular
        with st.expander("👤 Identificação do Assinante & Titular", expanded=True):
            col1, col2 = st.columns(2)
            with col1:
                st.markdown(f"**Ramal:** `{ramal_num}`")
                st.markdown(f"**Nome Titular:** {get_val(row, 'nome_titular', 'annu_name')}")
                st.markdown(f"**Complemento:** {get_val(row, 'complemento', 'annu_first_name', 'utf8_phone_book_first_name')}")
                st.markdown(f"**Nome Exibido:** {get_val(row, 'nome_exibido', 'phone_book_name', 'display_name')}")
            with col2:
                st.markdown(f"**Login Externo:** `{get_val(row, 'login_externo', 'external_login')}`")
                st.markdown(f"**E-mail:** {get_val(row, 'email', 'mail_address')}")
                st.markdown(f"**Idioma (Language_Id):** `{get_val(row, 'language_id')}`")
                st.markdown(f"**Qtd. Usuários:** `{get_val(row, 'number_of_users')}`")

        # 2. Rede, IP & Hardware
        with st.expander("🌐 Rede, IP & Hardware", expanded=True):
            col1, col2 = st.columns(2)
            with col1:
                st.markdown(f"**Endereço IP:** `{get_val(row, 'endereco_ip', 'ip_address')}`")
                st.markdown(f"**MAC Address:** `{get_val(row, 'mac_address', 'ethernet_address')}`")
                st.markdown(f"**Categoria Dispositivo:** `{get_val(row, 'categoria_dispositivo')}`")
            with col2:
                st.markdown(f"**Endereço PABX:** `{get_val(row, 'endereco_equipamento')}` (Rack `{get_val(row, 'rack', 'equipment_address_rack')}` / Placa `{get_val(row, 'placa', 'equipment_address_board')}` / Terminal `{get_val(row, 'terminal', 'equipment_address_terminal')}`)")
                st.markdown(f"**Criptografia Nativa:** `{get_val(row, 'nativeencryption')}`")
                st.markdown(f"**Suporte a Vídeo:** `{get_val(row, 'videosupportprofile')}`")

        # 3. Telefonia, Grupo de Captura & Regras de Acesso
        with st.expander("📞 Telefonia, Grupo de Captura & Regras de Acesso", expanded=True):
            col1, col2 = st.columns(2)
            with col1:
                st.markdown(f"**Grupo de Captura (Pickup Group):** `{get_val(row, 'grupo_captura', 'pickup_group_name')}`")
                st.markdown(f"**Cat. Rede Pública (Public Network):** `{get_val(row, 'cat_rede_publica', 'public_network_category_id')}`")
                st.markdown(f"**Cat. Encaminhamento Externo:** `{get_val(row, 'external_forwarding_category_id')}`")
                st.markdown(f"**Cat. Medição / Tarifação:** `{get_val(row, 'metering_category')}`")
            with col2:
                st.markdown(f"**Centro de Custo:** `{get_val(row, 'centro_de_custo', 'cost_center_name', 'cost_center_id')}`")
                st.markdown(f"**Função Adm:** `{get_val(row, 'funcao_adm', 'function')}`")
                st.markdown(f"**Cat. Chamador (Caller Category):** `{get_val(row, 'caller_category')}`")
                st.markdown(f"**Cat. Conexão / Dados:** `{get_val(row, 'connection_category_id')}` / `{get_val(row, 'data_connection_category_id')}`")

        # 4. Tipo de Estação, Perfil & Teclado
        with st.expander("⚙️ Tipo de Estação, Perfil & Teclado", expanded=False):
            col1, col2 = st.columns(2)
            with col1:
                st.markdown(f"**Tipo de Estação:** `{get_val(row, 'tipo_de_estacao', 'tipo_estacao', 'station_type')}`")
                st.markdown(f"**Subtipo de Estação:** `{get_val(row, 'subtipo', 'station_sub_type')}`")
                st.markdown(f"**Função / Role:** `{get_val(row, 'funcao_role', 'set_role')}`")
                st.markdown(f"**Perfil DM:** `{get_val(row, 'dm_profile')}` | **Tipo:** `{get_val(row, 'profile_type')}`")
            with col2:
                st.markdown(f"**Módulos de Extensão:** `{get_val(row, 'add_on_module_1')}` / `{get_val(row, 'add_on_module_2')}`")
                st.markdown(f"**Teclado Interno:** `{get_val(row, 'internal_keyboard')}`")
                st.markdown(f"**Entidade / Domínio:** `{get_val(row, 'entity_number')}` / `{get_val(row, 'domain_identifier')}`")
                st.markdown(f"**Licença Opex:** `{get_val(row, 'opex_license')}`")

        # 5. Correio de Voz & Recursos Especiais
        with st.expander("🎙️ Correio de Voz & Recursos Especiais", expanded=False):
            col1, col2 = st.columns(2)
            with col1:
                st.markdown(f"**Tipo Correio de Voz:** `{get_val(row, 'voice_mail_type')}`")
                st.markdown(f"**Ramal Correio de Voz:** `{get_val(row, 'voice_mail_directory_number')}`")
                st.markdown(f"**Grupo ACD:** `{get_val(row, 'acd_group_directory_number')}` ({get_val(row, 'acd_type')})")
            with col2:
                st.markdown(f"**Controle PIN:** `{get_val(row, 'pin_group_control')}` (Grupo `{get_val(row, 'user_pin_group')}`)")
                st.markdown(f"**Atendimento VIP:** `{get_val(row, 'vip')}`")
                st.markdown(f"**Chamada Urgente:** `{get_val(row, 'urgent_call')}`")

        # 6. Todos os Dados Brutos da API
        with st.expander("🔍 Dicionário Completo de Dados Brutos (API OXE)", expanded=False):
            clean_dict = {k: v for k, v in row.to_dict().items() if str(v).strip() and str(v).strip().lower() not in ["none", "nan", "null"]}
            st.json(clean_dict)

        if st.button("Fechar", key=f"close_ramal_modal_{ramal_num}"):
            st.rerun()

    # Mapeamento de colunas amigáveis para exibição
    column_rename_map = {
        "ramal": "Ramal",
        "nome_titular": "Nome / Titular",
        "complemento": "Complemento",
        "nome_exibido": "Nome Exibido",
        "tipo_estacao": "Tipo de Estação",
        "subtipo": "Subtipo",
        "funcao_role": "Função / Role",
        "grupo_captura": "Grupo de Captura",
        "cat_rede_publica": "Cat. Rede Pública",
        "login_externo": "Login Externo",
        "email": "E-mail",
        "rack": "Rack",
        "placa": "Placa",
        "terminal": "Terminal",
        "endereco_equipamento": "Endereço Equipamento",
        "ip_address": "Endereço IP",
        "mac_address": "MAC Address",
        "categoria_dispositivo": "Categoria",
        "status_equipamento": "Status",
        "data_atualizacao": "Última Atualização"
    }

    df_display = df_filtered.rename(columns=column_rename_map)

    cols_order = [
        "Ramal", "Nome Exibido", "Grupo de Captura", "Cat. Rede Pública",
        "Tipo de Estação", "Subtipo", "Endereço IP", "MAC Address",
        "Endereço Equipamento", "Categoria", "Função / Role", "Login Externo", "E-mail"
    ]
    existing_cols = [c for c in cols_order if c in df_display.columns]

    # Paginação
    df_page, current_page, total_pages, total_items = paginate_items(
        df_display[existing_cols],
        page_key="central_oxe",
        items_per_page=items_per_page
    )

    if "last_selected_ramal" not in st.session_state:
        st.session_state["last_selected_ramal"] = None

    selection_event = st.dataframe(
        df_page,
        width='stretch',
        hide_index=True,
        on_select="rerun",
        selection_mode="single-row",
        key="tabela_oxe_datagrid"
    )

    selected_rows = selection_event.selection.rows if hasattr(selection_event, "selection") else []

    if selected_rows:
        current_selected = selected_rows[0]
        if st.session_state["last_selected_ramal"] != current_selected:
            st.session_state["last_selected_ramal"] = current_selected
            row_data = df_filtered.iloc[(current_page - 1) * items_per_page + current_selected]
            show_ramal_details(row_data)
    else:
        st.session_state["last_selected_ramal"] = None

    # Régua de controles de paginação
    render_pagination_controls(
        page_key="central_oxe",
        current_page=current_page,
        total_pages=total_pages,
        total_items=total_items,
        items_per_page=items_per_page
    )

    # -----------------------------------------------------------------------------
    # BOTÃO DE EXPORTAÇÃO
    # -----------------------------------------------------------------------------
    st.markdown("---")
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
        df_display[existing_cols].to_excel(writer, sheet_name="Central Telefônica", index=False)
    
    st.download_button(
        label="📥 Baixar Tabela Filtrada em Excel (.xlsx)",
        data=output.getvalue(),
        file_name=f"Central_Telefonica_Filtrada_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        width='stretch'
    )


