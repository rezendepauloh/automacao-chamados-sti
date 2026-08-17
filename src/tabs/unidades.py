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
from src.scrapers.unidades_scraper import check_unidades_sync_running, read_unidades_last_log_lines
from src.scrapers.ramais_scraper import check_ramais_sync_running, read_ramais_last_log_lines
from src.components.status_banner import render_log_expander
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
                submitted = st.form_submit_button("💾 Salvar Alterações", type="primary", width='stretch')
            with col_del:
                btn_delete = st.form_submit_button("🗑️ Excluir Registro", width='stretch')

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

    if st.button("💾 Salvar no Banco Local", type="primary", width='stretch'):
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


@st.dialog("📞 Detalhes do Ramal Telefônico")
def modal_detalhes_ramal(row_data: dict):
    st.markdown(f"### 👤/🏢 {row_data.get('setor_nome', 'N/A')}")
    st.caption("Informação de ramal extraída da Intranet do MPMS.")
    st.markdown("---")
    m1, m2 = st.columns(2)
    with m1:
        st.markdown(f"**📍 Localidade:** {row_data.get('localidade', 'N/A')}")
        st.markdown(f"**🌐 Abrangência:** {row_data.get('tipo', 'N/A')}")
    with m2:
        st.markdown(f"**📞 Telefone / Ramal:** {row_data.get('telefone_ramal', 'N/A')}")

        # Formatação segura da data para o modal
        dt_raw = row_data.get('data_atualizacao')
        dt_formatada = "N/A"
        if pd.notna(dt_raw):
            try:
                dt_formatada = pd.to_datetime(dt_raw).strftime('%d/%m/%Y')
            except:
                dt_formatada = str(dt_raw)
        st.markdown(f"**🔄 Atualizado em:** {dt_formatada}")

    st.markdown("<br>", unsafe_allow_html=True)
    if st.button("Fechar", width='stretch'):
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

    # --- 1. ELEVAÇÃO DAS SUBTABS DE NAVEGAÇÃO DA PÁGINA ---
    UNIDADES_SUBTAB_MAP = {
        "geral": "🏢 Relação Unificada de Unidades",
        "ramais": "📞 Lista de Ramais (Telefonia)"
    }

    selected_tab = render_subtabs(UNIDADES_SUBTAB_MAP, default_slug="geral", key="unidades_subtab_radio")

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

    # --- 2. SIDEBAR: SEGREGAÇÃO DO MENU "AÇÕES E GESTÃO" ---
    st.sidebar.markdown("## ⚙️ Ações e Gestão")

    if selected_tab == "🏢 Relação Unificada de Unidades":
        if st.sidebar.button("➕ Novo Setor Interno", type="primary", width='stretch'):
            modal_novo_setor_manual()

        st.sidebar.markdown("<br>", unsafe_allow_html=True)

        if unidades_ativo:
            st.sidebar.button("🤖 Unidades em Atualização...", width='stretch', disabled=True)
        else:
            creationflags = subprocess.CREATE_NO_WINDOW if sys.platform == "win32" else 0
            if st.sidebar.button("🔄 Rodar Scraper Completo (Web)", width='stretch', help="Executa o scraper completo buscando dados atualizados no portal do MPMS em segundo plano."):
                import time
                subprocess.Popen([sys.executable, "src/scrapers/unidades_scraper.py"], creationflags=creationflags)
                time.sleep(0.8)
                st.toast("🚀 Scraper Completo de Unidades iniciado em segundo plano!", icon="🤖")
                st.rerun()

            if st.sidebar.button("⚡ Atualização Rápida (Só Manuais)", width='stretch', help="Atualiza a base unificada de unidades com as unidades manuais do banco."):
                import time
                subprocess.Popen([sys.executable, "src/scrapers/unidades_scraper.py", "--only-manual"], creationflags=creationflags)
                time.sleep(0.8)
                st.toast("⚡ Sincronização rápida de manuais iniciada!", icon="⚡")
                st.rerun()

    elif selected_tab == "📞 Lista de Ramais (Telefonia)":
        if ramais_ativo:
            st.sidebar.button("🤖 Ramais em Sincronização...", width='stretch', disabled=True)
        else:
            creationflags = subprocess.CREATE_NO_WINDOW if sys.platform == "win32" else 0
            if st.sidebar.button("🔄 Atualizar Ramais (Intranet)", width='stretch', help="Executa o robô de extração de ramais em PDF da Intranet do MPMS em segundo plano."):
                import time
                subprocess.Popen([sys.executable, "src/scrapers/ramais_scraper.py"], creationflags=creationflags)
                time.sleep(0.5)
                st.toast("🚀 Sincronização de ramais iniciada em segundo plano!", icon="📞")
                st.rerun()

    st.sidebar.markdown("---")

    # --- 3 & 4. SIDEBAR: SEGREGAÇÃO DE FILTROS DE BUSCA E PAGINAÇÃO ---
    st.sidebar.markdown("## 🔍 Filtros de Busca")

    if selected_tab == "🏢 Relação Unificada de Unidades":
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

    elif selected_tab == "📞 Lista de Ramais (Telefonia)":
        df_ramais_sb = get_ramais_df()

        localidade_opts = ["Todas"]
        setor_opts = ["Todos"]
        abrangencia_opts = ["Todas"]

        if not df_ramais_sb.empty:
            if "localidade" in df_ramais_sb.columns:
                localidade_opts += sorted([str(x).strip() for x in df_ramais_sb['localidade'].dropna().unique() if str(x).strip()])
            if "setor_nome" in df_ramais_sb.columns:
                setor_opts += sorted([str(x).strip() for x in df_ramais_sb['setor_nome'].dropna().unique() if str(x).strip()])
            if "tipo" in df_ramais_sb.columns:
                abrangencia_opts += sorted([str(x).strip() for x in df_ramais_sb['tipo'].dropna().unique() if str(x).strip()])

        selected_localidade_r = st.sidebar.selectbox("📍 Filtrar por Localidade:", localidade_opts)
        selected_setor_r = st.sidebar.selectbox("🏢 Filtrar por Setor / Membro:", setor_opts)
        selected_abrangencia_r = st.sidebar.selectbox("🌐 Filtrar por Abrangência:", abrangencia_opts)
        st.sidebar.text_input("🔎 Pesquisar Ramal (Setor, Localidade, Nome, Número):", "", key="search_ramais_input")

        items_per_page = render_items_per_page_selector(
            key_prefix="ramais_telefonia",
            options=[10, 20, 50, 100, "Todos"],
            default_index=1,
            label="📄 Registros por página:"
        )

    # --- ACCORDIONS DE LOGS E PROGRESSO EM SEGUNDO PLANO (NO CORPO PRINCIPAL) ---
    render_log_expander(
        "🤖 Robô de Unidades Rodando em Segundo Plano – Acompanhar Progresso",
        unidades_ativo,
        read_unidades_last_log_lines,
        check_unidades_sync_running,
        "O robô de unidades está coletando dados no portal do MPMS neste momento. Você pode continuar navegando normalmente!"
    )

    render_log_expander(
        "📞 Robô de Ramais Rodando em Segundo Plano – Acompanhar Progresso",
        ramais_ativo,
        read_ramais_last_log_lines,
        check_ramais_sync_running,
        "O robô de ramais está baixando os PDFs da Intranet e processando a telefonia neste momento. O uso da aplicação permanece livre!"
    )

    # --- RENDERIZAÇÃO DAS SUBTABS ---
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
                width='stretch',
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
        st.caption("Dados extraídos dos documentos oficiais de telefonia da Intranet do MPMS. Clique em qualquer linha para abrir a ficha de detalhes.")

        df_ramais = get_ramais_df()

        if df_ramais.empty:
            st.info("Nenhum ramal cadastrado no banco de dados local. Clique no botão '🔄 Atualizar Ramais (Intranet)' na barra lateral para sincronizar.")
        else:
            search_ramal = st.session_state.get("search_ramais_input", "").strip().lower()

            df_filtered_r = df_ramais.copy()

            if 'selected_localidade_r' in locals() and selected_localidade_r != "Todas":
                df_filtered_r = df_filtered_r[df_filtered_r["localidade"] == selected_localidade_r]

            if 'selected_setor_r' in locals() and selected_setor_r != "Todos":
                df_filtered_r = df_filtered_r[df_filtered_r["setor_nome"] == selected_setor_r]

            if 'selected_abrangencia_r' in locals() and selected_abrangencia_r != "Todas":
                df_filtered_r = df_filtered_r[df_filtered_r["tipo"] == selected_abrangencia_r]

            if search_ramal:
                mask_r = (
                    df_filtered_r["localidade"].astype(str).str.lower().str.contains(search_ramal, na=False) |
                    df_filtered_r["setor_nome"].astype(str).str.lower().str.contains(search_ramal, na=False) |
                    df_filtered_r["telefone_ramal"].astype(str).str.lower().str.contains(search_ramal, na=False) |
                    df_filtered_r["tipo"].astype(str).str.lower().str.contains(search_ramal, na=False)
                )
                df_filtered_r = df_filtered_r[mask_r]

            df_filtered_r['data_atualizacao'] = pd.to_datetime(df_filtered_r['data_atualizacao'], errors='coerce')

            st.markdown(f"**Exibindo {len(df_filtered_r)} de {len(df_ramais)} ramais**")

            df_page_r, current_page_r, total_pages_r, total_items_r = paginate_items(
                df_filtered_r,
                page_key="ramais_telefonia",
                items_per_page=items_per_page
            )

            selection_r = st.dataframe(
                df_page_r,
                column_config={
                    "id": st.column_config.NumberColumn("ID"),
                    "localidade": st.column_config.TextColumn("Localidade / Prédio / Comarca"),
                    "setor_nome": st.column_config.TextColumn("Setor / Cargo / Membro"),
                    "telefone_ramal": st.column_config.TextColumn("Telefone / Ramal"),
                    "tipo": st.column_config.TextColumn("Abrangência"),
                    "data_atualizacao": st.column_config.DatetimeColumn("Última Atualização", format="DD/MM/YYYY HH:mm"),
                },
                hide_index=True,
                width='stretch',
                on_select="rerun",
                selection_mode="single-row"
            )

            if selection_r and selection_r.get("selection") and selection_r["selection"].get("rows"):
                selected_row_idx_r = selection_r["selection"]["rows"][0]
                if selected_row_idx_r < len(df_page_r):
                    row_selected_r = df_page_r.iloc[selected_row_idx_r].to_dict()
                    modal_detalhes_ramal(row_selected_r)

            render_pagination_controls(
                page_key="ramais_telefonia",
                current_page=current_page_r,
                total_pages=total_pages_r,
                total_items=total_items_r,
                items_per_page=items_per_page
            )

