import sys
import subprocess
import pandas as pd
import streamlit as st
from pathlib import Path

root_dir = Path(__file__).parent.parent.parent
sys.path.insert(0, str(root_dir))

from src.config import OUTPUT_DIR_PRONTO
from src.database import (
    add_unidade_manual, 
    get_unidades_manuais, 
    delete_unidade_manual, 
    get_unidades_df, 
    get_ramais_df,
    update_unidade_manual_by_id
)
from src.unidades_scraper import check_unidades_sync_running, read_unidades_last_log_lines
from src.ramais_scraper import check_ramais_sync_running, read_ramais_last_log_lines
from src.components.subtabs import render_subtabs
from src.components.pagination import (
    render_items_per_page_selector,
    paginate_items,
    render_pagination_controls
)

UNIDADES_EXCEL_PATH = OUTPUT_DIR_PRONTO / "Unidades_MPMS.xlsx"


@st.dialog("📋 Detalhes e Gestão da Unidade")
def modal_detalhes_unidade(row_data: dict):
    """Exibe um modal interativo com detalhes da unidade. Se for manual, permite editar ou excluir."""
    manual_id = row_data.get("manual_id")
    is_manual = bool(manual_id and pd.notna(manual_id))
    
    if is_manual:
        st.markdown("### 📌 Editar Setor Manual")
        st.caption("Esta unidade foi cadastrada manualmente no banco de dados local. Você pode alterar seus campos ou excluí-la.")

        m_id = int(manual_id)

        with st.form(f"form_edit_unidade_manual_{m_id}"):
            c1, c2 = st.columns(2)
            with c1:
                cidade = st.text_input("Cidade *", row_data.get("Cidade", ""))
                tipo_val = row_data.get("Tipo", "Setor Interno")
                tipos_lista = ["Setor Interno", "Promotoria", "Procuradoria", "CAO", "PGJ", "GAECO", "Escola do MP", "Outros"]
                tipo_idx = tipos_lista.index(tipo_val) if tipo_val in tipos_lista else 0
                tipo = st.selectbox("Tipo de Unidade *", tipos_lista, index=tipo_idx)
                setor = st.text_input("Nome do Setor / Unidade *", row_data.get("Setor", ""))
                sigla = st.text_input("Sigla", row_data.get("Sigla", ""))
            with c2:
                titular = st.text_input("Titular / Responsável", row_data.get("Titular", ""))
                u_predio = st.text_input("Localidade / Prédio", row_data.get("Unidade (Prédio)", ""))
                telefone = st.text_input("Telefone / Contato", row_data.get("Telefone", ""))
                url = st.text_input("URL / Link", row_data.get("URL", ""))

            st.markdown("<br>", unsafe_allow_html=True)
            col_sub, col_del = st.columns([2, 1])
            with col_sub:
                submitted = st.form_submit_button("💾 Salvar Alterações", type="primary", use_container_width=True)
            with col_del:
                btn_delete = st.form_submit_button("🗑️ Excluir Registro", use_container_width=True)

            if submitted:
                if not setor or not cidade:
                    st.error("Preencha ao menos Cidade e Nome do Setor!")
                else:
                    update_unidade_manual_by_id(m_id, cidade, tipo, setor, sigla, titular, u_predio, telefone, url)
                    st.toast("✅ Registro atualizado com sucesso!", icon="🎉")
                    st.rerun()

            if btn_delete:
                delete_unidade_manual(m_id)
                st.toast("🗑️ Registro manual excluído!", icon="🗑️")
                st.rerun()
    else:
        st.markdown(f"### 🌐 {row_data.get('Setor', 'Unidade')}")
        st.caption("Dados de unidade oficial sincronizada a partir do Portal do MPMS.")

        m1, m2 = st.columns(2)
        with m1:
            st.markdown(f"**🏙️ Cidade:** {row_data.get('Cidade', 'N/A')}")
            st.markdown(f"**🏷️ Tipo:** {row_data.get('Tipo', 'N/A')}")
            st.markdown(f"**🏷️ Sigla:** {row_data.get('Sigla', 'N/A')}")
            st.markdown(f"**👤 Titular:** {row_data.get('Titular', 'N/A')}")
        with m2:
            st.markdown(f"**🏢 Prédio / Localidade:** {row_data.get('Unidade (Prédio)', 'N/A')}")
            st.markdown(f"**📞 Telefone:** {row_data.get('Telefone', 'N/A')}")
            url_p = row_data.get("URL")
            if url_p and str(url_p).strip() and str(url_p).strip() != "None":
                st.markdown(f"**🔗 Link Portal:** [{url_p}]({url_p})")

        st.markdown("<br>", unsafe_allow_html=True)
        st.info("ℹ️ Unidades do Portal são mantidas automaticamente pelo robô de sincronização.")


@st.dialog("➕ Novo Setor Interno (Unidade Manual)")
def modal_novo_setor_manual():
    """Modal nativo do Streamlit (@st.dialog) para cadastro de novos setores manuais."""
    st.write("Preencha as informações do setor interno para salvar no banco de dados:")

    cidade = st.text_input("🏙️ Cidade *", value="Campo Grande")
    tipo = st.selectbox("🏷️ Tipo de Unidade *", ["Setor Interno", "CAO", "PGJ", "GAECO", "Escola do MP", "Outros"])
    setor = st.text_input("🏢 Nome do Setor / Unidade *", placeholder="Ex: Divisão de Suporte de TI")
    sigla = st.text_input("🏷️ Sigla", placeholder="Ex: STI")
    titular = st.text_input("👤 Titular / Responsável", placeholder="Ex: Nome do Chefe")
    unidade_predio = st.text_input("📍 Prédio / Localidade", value="PGJ - Bloco B - 1º Andar")
    telefone = st.text_input("📞 Telefone / Ramal", placeholder="Ex: (67) 3318-3939 / Ramal 3939")
    url = st.text_input("🔗 Link / Portal", placeholder="https://...")

    if st.button("💾 Salvar no Banco Local", type="primary", use_container_width=True):
        if not setor or not cidade:
            st.error("Preencha ao menos Cidade e Nome do Setor!")
            return
        
        add_unidade_manual(
            cidade=cidade,
            tipo=tipo,
            setor=setor,
            sigla=sigla,
            titular=titular,
            unidade_predio=unidade_predio,
            telefone=telefone,
            url=url
        )
        st.toast("✅ Novo setor cadastrado com sucesso!", icon="🎉")
        st.rerun()


def render_unidades_page():
    """Renderiza a página Catálogo de Unidades do MPMS."""
    st.title("🏢 Catálogo de Unidades do MPMS")
    st.caption("Consulte promotorias, procuradorias e setores internos cadastrados no portal do MPMS e no banco local.")

    # --- CHECAGEM DE ROBÔS EM SEGUNDO PLANO ---
    unidades_ativo = check_unidades_sync_running()
    ramais_ativo = check_ramais_sync_running()

    # Notificação quando algum dos robôs conclui
    was_unidades = st.session_state.get("was_unidades_syncing", False)
    if was_unidades and not unidades_ativo:
        st.toast("🎉 Atualização do catálogo de unidades concluída com sucesso!", icon="🏢")
        st.cache_data.clear()
        st.session_state["was_unidades_syncing"] = False

    was_ramais = st.session_state.get("was_ramais_syncing", False)
    if was_ramais and not ramais_ativo:
        st.toast("🎉 Sincronização de ramais concluída com sucesso!", icon="📞")
        st.cache_data.clear()
        st.session_state["was_ramais_syncing"] = False

    if unidades_ativo:
        st.session_state["was_unidades_syncing"] = True
    if ramais_ativo:
        st.session_state["was_ramais_syncing"] = True

    # --- SIDEBAR: BOTÕES DE AÇÃO ---
    st.sidebar.markdown("## ⚙️ Ações e Gestão")

    if st.sidebar.button("➕ Novo Setor Interno", type="primary", use_container_width=True):
        modal_novo_setor_manual()

    st.sidebar.markdown("<br>", unsafe_allow_html=True)

    if unidades_ativo:
        st.sidebar.button("🤖 Unidades em Atualização...", use_container_width=True, disabled=True)
    else:
        if st.sidebar.button("🔄 Rodar Scraper Completo (Web)", use_container_width=True, help="Executa o scraper completo buscando dados atualizados no portal do MPMS em segundo plano."):
            subprocess.Popen([sys.executable, "src/unidades_scraper.py"])
            st.toast("🚀 Scraper Completo de Unidades iniciado em segundo plano!", icon="🤖")
            st.rerun()

        if st.sidebar.button("⚡ Atualização Rápida (Só Manuais)", use_container_width=True, help="Atualiza a base unificada de unidades com as unidades manuais do banco."):
            subprocess.Popen([sys.executable, "src/unidades_scraper.py", "--only-manual"])
            st.toast("⚡ Sincronização rápida de manuais iniciada!", icon="⚡")
            st.rerun()

    if ramais_ativo:
        st.sidebar.button("🤖 Ramais em Sincronização...", use_container_width=True, disabled=True)
    else:
        if st.sidebar.button("🔄 Atualizar Ramais (Intranet)", use_container_width=True, help="Executa o robô de extração de ramais em PDF da Intranet do MPMS em segundo plano."):
            subprocess.Popen([sys.executable, "src/ramais_scraper.py"])
            st.toast("🚀 Sincronização de ramais iniciada em segundo plano!", icon="📞")
            st.rerun()

    st.sidebar.markdown("---")

    # --- ACCORDIONS DE LOGS E PROGRESSO EM SEGUNDO PLANO ---
    if unidades_ativo:
        with st.expander("🤖 Robô de Unidades Rodando em Segundo Plano – Acompanhar Progresso", expanded=True):
            st.info("O robô de unidades está coletando dados no portal do MPMS neste momento. Você pode continuar navegando normalmente!")
            logs_u = read_unidades_last_log_lines(15)
            st.code(logs_u, language="text")
            if st.button("🔄 Atualizar Log de Unidades", key="btn_update_unidades_log"):
                st.rerun()

    if ramais_ativo:
        with st.expander("📞 Robô de Ramais Rodando em Segundo Plano – Acompanhar Progresso", expanded=True):
            st.info("O robô de ramais está baixando os PDFs da Intranet e processando a telefonia neste momento. O uso da aplicação permanece livre!")
            logs_r = read_ramais_last_log_lines(15)
            st.code(logs_r, language="text")
            if st.button("🔄 Atualizar Log de Ramais", key="btn_update_ramais_log"):
                st.rerun()

    # --- CARREGA DADOS DIRETO DO SQLite ---
    df_unidades = get_unidades_df()
    if df_unidades.empty and UNIDADES_EXCEL_PATH.exists():
        try:
            df_excel = pd.read_excel(UNIDADES_EXCEL_PATH)
            df_excel.fillna("", inplace=True)
            from src.database import save_unidades_to_db
            save_unidades_to_db(df_excel)
            df_unidades = get_unidades_df()
        except Exception:
            pass

    # --- FILTROS SIDEBAR ---
    st.sidebar.markdown("## 🔍 Filtros de Busca")

    cidades_opts = ["Todas"]
    tipos_opts = ["Todos"]
    origem_opts = ["Todas", "📌 Manual", "🌐 Portal Web"]

    if not df_unidades.empty:
        if "Cidade" in df_unidades.columns:
            cidades_opts += sorted([str(c).strip() for c in df_unidades["Cidade"].unique() if str(c).strip()])
        if "Tipo" in df_unidades.columns:
            tipos_opts += sorted([str(t).strip() for t in df_unidades["Tipo"].unique() if str(t).strip()])

    selected_cidade = st.sidebar.selectbox("🏙️ Filtrar por Cidade:", cidades_opts)
    selected_tipo = st.sidebar.selectbox("🏷️ Filtrar por Tipo de Unidade:", tipos_opts)
    selected_origem = st.sidebar.selectbox("📌 Origem do Registro:", origem_opts)
    search_query = st.sidebar.text_input("🔎 Buscar (Setor, Sigla, Titular, Prédio):", "").strip().lower()

    items_per_page = render_items_per_page_selector(
        key_prefix="unidades_mp",
        options=[10, 20, 50, 100, "Todos"],
        default_index=1,
        label="📄 Registros por página:"
    )

    # --- SUBTABS DE NAVEGAÇÃO DA PÁGINA ---
    UNIDADES_SUBTAB_MAP = {
        "geral": "🏢 Relação Unificada de Unidades",
        "ramais": "📞 Lista de Ramais (Telefonia)"
    }

    selected_tab = render_subtabs(UNIDADES_SUBTAB_MAP, default_slug="geral", key="unidades_subtab_radio")

    if selected_tab == "🏢 Relação Unificada de Unidades":
        if df_unidades.empty:
            st.warning("⚠️ Nenhuma unidade cadastrada no banco SQLite. Clique em 'Rodar Scraper Completo (Web)' ou 'Atualização Rápida (Só Manuais)' na barra lateral para gerar.")
        else:
            df_filtered = df_unidades.copy()

            if selected_cidade != "Todas":
                df_filtered = df_filtered[df_filtered["Cidade"] == selected_cidade]

            if selected_tipo != "Todos":
                df_filtered = df_filtered[df_filtered["Tipo"] == selected_tipo]

            if selected_origem != "Todas":
                df_filtered = df_filtered[df_filtered["Origem"] == selected_origem]

            if search_query:
                mask = (
                    df_filtered["Setor"].astype(str).str.lower().str.contains(search_query, na=False) |
                    df_filtered["Sigla"].astype(str).str.lower().str.contains(search_query, na=False) |
                    df_filtered["Titular"].astype(str).str.lower().str.contains(search_query, na=False) |
                    df_filtered["Unidade (Prédio)"].astype(str).str.lower().str.contains(search_query, na=False) |
                    df_filtered["Telefone"].astype(str).str.lower().str.contains(search_query, na=False)
                )
                df_filtered = df_filtered[mask]

            st.markdown(f"### 🏢 Relação Unificada de Unidades ({len(df_filtered)} registros)")
            st.caption("Clique em qualquer linha da tabela para abrir o modal de detalhes, edição (setores manuais) ou exclusão.")

            df_page, current_page, total_pages, total_items = paginate_items(
                df_filtered,
                page_key="unidades_mp",
                items_per_page=items_per_page
            )

            selection = st.dataframe(
                df_page,
                column_config={
                    "Origem": st.column_config.TextColumn("Origem"),
                    "Cidade": st.column_config.TextColumn("Cidade"),
                    "Tipo": st.column_config.TextColumn("Tipo"),
                    "Setor": st.column_config.TextColumn("Nome do Setor / Unidade"),
                    "Sigla": st.column_config.TextColumn("Sigla"),
                    "Titular": st.column_config.TextColumn("Titular / Responsável"),
                    "Unidade (Prédio)": st.column_config.TextColumn("Localidade / Prédio"),
                    "Telefone": st.column_config.TextColumn("Telefone / Contato"),
                    "URL": st.column_config.LinkColumn("Link Portal", display_text="🔗 Abrir Portal"),
                },
                column_order=["Origem", "Cidade", "Tipo", "Setor", "Sigla", "Titular", "Unidade (Prédio)", "Telefone", "URL"],
                hide_index=True,
                use_container_width=True,
                on_select="rerun",
                selection_mode="single-row"
            )

            if selection and selection.get("selection") and selection["selection"].get("rows"):
                selected_row_idx = selection["selection"]["rows"][0]
                if selected_row_idx < len(df_page):
                    row_selected = df_page.iloc[selected_row_idx].to_dict()
                    modal_detalhes_unidade(row_selected)

            render_pagination_controls(
                page_key="unidades_mp",
                current_page=current_page,
                total_pages=total_pages,
                total_items=total_items,
                items_per_page=items_per_page
            )

    elif selected_tab == "📞 Lista de Ramais (Telefonia)":
        st.markdown("### 📞 Lista Oficial de Ramais Telefônicos do MPMS")
        st.caption("Dados extraídos dos documentos oficiais de telefonia da Intranet do MPMS.")

        df_ramais = get_ramais_df()

        if df_ramais.empty:
            st.info("Nenhum ramal cadastrado no banco de dados local. Clique no botão '🔄 Atualizar Ramais (Intranet)' na barra lateral para sincronizar.")
        else:
            search_ramal = st.text_input("🔎 Pesquisar Ramal (Setor, Localidade, Nome, Número):", "", key="search_ramais_input").strip().lower()

            df_filtered_r = df_ramais.copy()
            if search_ramal:
                mask_r = (
                    df_filtered_r["localidade"].astype(str).str.lower().str.contains(search_ramal, na=False) |
                    df_filtered_r["setor_nome"].astype(str).str.lower().str.contains(search_ramal, na=False) |
                    df_filtered_r["telefone_ramal"].astype(str).str.lower().str.contains(search_ramal, na=False) |
                    df_filtered_r["tipo"].astype(str).str.lower().str.contains(search_ramal, na=False)
                )
                df_filtered_r = df_filtered_r[mask_r]

            st.markdown(f"**Exibindo {len(df_filtered_r)} de {len(df_ramais)} ramais**")

            df_page_r, current_page_r, total_pages_r, total_items_r = paginate_items(
                df_filtered_r,
                page_key="ramais_telefonia",
                items_per_page=items_per_page
            )

            st.dataframe(
                df_page_r,
                column_config={
                    "id": st.column_config.NumberColumn("ID"),
                    "localidade": st.column_config.TextColumn("Localidade / Prédio / Comarca"),
                    "setor_nome": st.column_config.TextColumn("Setor / Cargo / Membro"),
                    "telefone_ramal": st.column_config.TextColumn("Telefone / Ramal"),
                    "tipo": st.column_config.TextColumn("Abrangência"),
                    "data_atualizacao": st.column_config.TextColumn("Última Atualização"),
                },
                hide_index=True,
                use_container_width=True
            )

            render_pagination_controls(
                page_key="ramais_telefonia",
                current_page=current_page_r,
                total_pages=total_pages_r,
                total_items=total_items_r,
                items_per_page=items_per_page
            )

