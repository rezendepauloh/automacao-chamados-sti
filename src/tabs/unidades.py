import sys
import subprocess
import pandas as pd
import streamlit as st
from pathlib import Path

root_dir = Path(__file__).parent.parent.parent
sys.path.insert(0, str(root_dir))

from src.config import OUTPUT_DIR_PRONTO
from src.database import add_unidade_manual, get_unidades_manuais, delete_unidade_manual, get_unidades_df
from src.components.pagination import (
    render_items_per_page_selector,
    paginate_items,
    render_pagination_controls
)

UNIDADES_EXCEL_PATH = OUTPUT_DIR_PRONTO / "Unidades_MPMS.xlsx"


@st.dialog("➕ Novo Setor Interno (Unidade Manual)")
def modal_novo_setor_manual():
    """Modal nativo do Streamlit (@st.dialog) para cadastro de novos setores manuais."""
    st.write("Preencha as informações do setor interno para salvar no banco de dados:")

    cidade = st.text_input("🏙️ Cidade *", value="Campo Grande")
    tipo = st.selectbox("🏷️ Tipo de Unidade *", ["Setor Interno", "CAO", "PGJ", "GAECO", "Escola do MP", "Outros"])
    setor = st.text_input("🏢 Nome do Setor / Unidade *", placeholder="Ex: Divisão de Suporte Técnico")
    sigla = st.text_input("🔖 Sigla", placeholder="Ex: STI-SUP")
    titular = st.text_input("👤 Titular / Chefia", placeholder="Nome do responsável")
    unidade_predio = st.text_input("🏛️ Prédio / Localização Física", placeholder="Ex: Sede PGJ - Bloco A")
    telefone = st.text_input("📞 Telefone / Ramal", placeholder="Ex: (67) 3318-2000")
    url = st.text_input("🌐 Link / URL Portal", placeholder="https://...")

    if st.button("💾 Salvar Setor", type="primary", use_container_width=True):
        if not setor.strip() or not cidade.strip():
            st.error("Por favor, preencha a Cidade e o Nome do Setor.")
            return

        add_unidade_manual(
            cidade=cidade.strip(),
            tipo=tipo,
            setor=setor.strip(),
            sigla=sigla.strip(),
            titular=titular.strip(),
            unidade_predio=unidade_predio.strip(),
            telefone=telefone.strip(),
            url=url.strip()
        )
        st.toast("✅ Novo setor cadastrado com sucesso!", icon="🎉")
        st.rerun()


def render_unidades_page():
    """Renderiza a página Catálogo de Unidades do MPMS."""
    st.title("🏢 Catálogo de Unidades do MPMS")
    st.caption("Consulte promotorias, procuradorias e setores internos cadastrados no portal do MPMS e no banco local.")

    # --- SIDEBAR: BOTÕES DE AÇÃO ---
    st.sidebar.markdown("## ⚙️ Ações e Gestão")

    if st.sidebar.button("➕ Novo Setor Interno", type="primary", use_container_width=True):
        modal_novo_setor_manual()

    st.sidebar.markdown("<br>", unsafe_allow_html=True)

    if st.sidebar.button("🔄 Rodar Scraper Completo (Web)", use_container_width=True, help="Executa o scraper completo buscando dados atualizados no portal do MPMS em segundo plano."):
        subprocess.Popen([sys.executable, "src/unidades_scraper.py"])
        st.toast("🚀 Scraper Completo iniciado em segundo plano!", icon="🤖")

    if st.sidebar.button("🔄 Atualização Rápida (Só Manuais)", use_container_width=True, help="Atualiza a base unificada de unidades com as unidades manuais do banco."):
        subprocess.Popen([sys.executable, "src/unidades_scraper.py", "--only-manual"])
        st.toast("⚡ Sincronização rápida de manuais iniciada!", icon="⚡")

    st.sidebar.markdown("---")

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

    if not df_unidades.empty:
        if "Cidade" in df_unidades.columns:
            cidades_opts += sorted([str(c).strip() for c in df_unidades["Cidade"].unique() if str(c).strip()])
        if "Tipo" in df_unidades.columns:
            tipos_opts += sorted([str(t).strip() for t in df_unidades["Tipo"].unique() if str(t).strip()])

    selected_cidade = st.sidebar.selectbox("🏙️ Filtrar por Cidade:", cidades_opts)
    selected_tipo = st.sidebar.selectbox("🏷️ Filtrar por Tipo de Unidade:", tipos_opts)
    search_query = st.sidebar.text_input("🔎 Buscar (Setor, Sigla, Titular, Prédio):", "").strip().lower()

    items_per_page = render_items_per_page_selector(
        key_prefix="unidades_mp",
        options=[10, 20, 50, 100, "Todos"],
        default_index=1,
        label="📄 Registros por página:"
    )

    # --- SUBTABS DE NAVEGAÇÃO DA PÁGINA ---
    subtab1, subtab2 = st.tabs(["📋 Relação Geral de Unidades", "⚙️ Gestão de Unidades Manuais (Banco Local)"])

    with subtab1:
        if df_unidades.empty:
            st.warning("⚠️ Nenhuma unidade cadastrada no banco SQLite. Clique em 'Rodar Scraper Completo (Web)' ou 'Atualização Rápida (Só Manuais)' na barra lateral para gerar.")
        else:
            df_filtered = df_unidades.copy()

            if selected_cidade != "Todas":
                df_filtered = df_filtered[df_filtered["Cidade"] == selected_cidade]

            if selected_tipo != "Todos":
                df_filtered = df_filtered[df_filtered["Tipo"] == selected_tipo]

            if search_query:
                mask = (
                    df_filtered["Setor"].astype(str).str.lower().str.contains(search_query, na=False) |
                    df_filtered["Sigla"].astype(str).str.lower().str.contains(search_query, na=False) |
                    df_filtered["Titular"].astype(str).str.lower().str.contains(search_query, na=False) |
                    df_filtered["Unidade (Prédio)"].astype(str).str.lower().str.contains(search_query, na=False) |
                    df_filtered["Telefone"].astype(str).str.lower().str.contains(search_query, na=False)
                )
                df_filtered = df_filtered[mask]

            st.markdown(f"### 📋 Relação Unificada de Unidades ({len(df_filtered)} registros)")

            df_page, current_page, total_pages, total_items = paginate_items(
                df_filtered,
                page_key="unidades_mp",
                items_per_page=items_per_page
            )

            st.dataframe(
                df_page,
                column_config={
                    "Cidade": st.column_config.TextColumn("Cidade"),
                    "Tipo": st.column_config.TextColumn("Tipo"),
                    "Setor": st.column_config.TextColumn("Nome do Setor / Unidade"),
                    "Sigla": st.column_config.TextColumn("Sigla"),
                    "Titular": st.column_config.TextColumn("Titular / Responsável"),
                    "Unidade (Prédio)": st.column_config.TextColumn("Localidade / Prédio"),
                    "Telefone": st.column_config.TextColumn("Telefone / Contato"),
                    "URL": st.column_config.LinkColumn("Link Portal", display_text="🔗 Abrir Portal"),
                },
                hide_index=True,
                use_container_width=True
            )

            render_pagination_controls(
                page_key="unidades_mp",
                current_page=current_page,
                total_pages=total_pages,
                total_items=total_items,
                items_per_page=items_per_page
            )

    with subtab2:
        st.markdown("### ⚙️ Setores e Unidades Cadastradas Manualmente")
        df_manuais_db = get_unidades_manuais()

        if df_manuais_db.empty:
            st.info("Nenhum setor cadastrado manualmente no banco de dados SQLite ainda. Clique em '➕ Novo Setor Interno' na barra lateral para adicionar.")
        else:
            st.dataframe(
                df_manuais_db,
                column_config={
                    "id": st.column_config.NumberColumn("ID"),
                    "cidade": st.column_config.TextColumn("Cidade"),
                    "tipo": st.column_config.TextColumn("Tipo"),
                    "setor": st.column_config.TextColumn("Setor"),
                    "sigla": st.column_config.TextColumn("Sigla"),
                    "titular": st.column_config.TextColumn("Titular"),
                    "unidade_predio": st.column_config.TextColumn("Prédio"),
                    "telefone": st.column_config.TextColumn("Telefone"),
                    "url": st.column_config.TextColumn("URL"),
                    "data_atualizacao": st.column_config.TextColumn("Atualizado em"),
                },
                hide_index=True,
                use_container_width=True
            )

            with st.expander("🗑️ Excluir Unidade Manual"):
                del_id = st.number_input("Digite o ID da unidade manual que deseja remover:", min_value=1, step=1)
                if st.button("🗑️ Remover Registro", type="primary"):
                    delete_unidade_manual(int(del_id))
                    st.toast(f"Registro ID {del_id} removido do banco!", icon="🗑️")
                    st.rerun()
