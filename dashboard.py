import streamlit as st
import pandas as pd
import sqlite3
from pathlib import Path
from datetime import datetime
import locale

# Configuração do locale para Português do Brasil
try:
    locale.setlocale(locale.LC_ALL, 'pt_BR.utf8')
except:
    try:
        locale.setlocale(locale.LC_ALL, 'Portuguese_Brazil.1252')
    except:
        pass

# Configuração da página para ocupar a tela toda e ter um título
st.set_page_config(page_title="Painel de Chamados - STI", layout="wide")

# CSS para ocultar apenas o botão Deploy e o rodapé padrão em inglês, mantendo o menu de três pontinhos
st.markdown("""
    <style>
    /* Oculta apenas o botão de Deploy do cabeçalho */
    [data-testid="stAppDeployButton"], .stAppDeployButton {
        display: none !important;
    }
    /* Remove rodapé padrão */
    footer {
        display: none !important;
    }
    </style>
""", unsafe_allow_html=True)

st.title("📊 Painel de Chamados Centralizado")
st.write("Visualize e interaja com os chamados do OTRS e CitSmart.")

DB_PATH = Path("chamados.db")

def load_data():
    if not DB_PATH.exists():
        return pd.DataFrame()
    conn = sqlite3.connect(DB_PATH)
    df = pd.read_sql_query("SELECT * FROM chamados", conn)
    conn.close()
    return df

df = load_data()

if df.empty:
    st.warning("Nenhum dado encontrado no banco de dados. Execute o orquestrador primeiro!")
else:
    # Tratamento de datas para exibição e filtro
    df['datetime_obj'] = pd.to_datetime(df['data_criacao'], errors='coerce')
    df['Data Formatada'] = df['datetime_obj'].dt.strftime('%d/%m/%Y %H:%M:%S')
    df['Data Formatada'] = df['Data Formatada'].fillna(df['data_criacao'])

    # Barra lateral de filtros
    st.sidebar.header("Filtros")
    
    # Filtro de Datas (Calendário)
    min_date = df['datetime_obj'].dropna().min().date() if not df['datetime_obj'].dropna().empty else datetime.now().date()
    max_date = df['datetime_obj'].dropna().max().date() if not df['datetime_obj'].dropna().empty else datetime.now().date()
    
    date_range = st.sidebar.date_input(
        "Intervalo de Datas",
        value=(min_date, max_date),
        min_value=min_date,
        max_value=max_date,
        format="DD/MM/YYYY"
    )
    
    # Filtros Multi-seleção (Sem padrão selecionado para não poluir)
    status_options = list(df['status'].unique())
    selected_status = st.sidebar.multiselect("Status", status_options, placeholder="Escolha as opções...")
    
    tag_options = list(df['tag'].dropna().unique())
    selected_tags = st.sidebar.multiselect("TAG", tag_options, placeholder="Escolha as opções...")
    
    loc_options = list(df['localidade_fisica'].dropna().unique())
    selected_locs = st.sidebar.multiselect("Localidade Física", loc_options, placeholder="Escolha as opções...")
    
    city_options = list(df['cidade_predio'].dropna().unique())
    selected_cities = st.sidebar.multiselect("Cidade - Prédio", city_options, placeholder="Escolha as opções...")
    
    unit_options = list(df['unidade'].dropna().unique())
    selected_units = st.sidebar.multiselect("Unidade", unit_options, placeholder="Escolha as opções...")
    
    # NOVO: Filtro de Base (CitSmart/OTRS)
    base_options = list(df['base'].dropna().unique())
    selected_bases = st.sidebar.multiselect("Base de Origem", base_options, placeholder="Escolha as opções...")
    
    # Filtro de Usuário (Busca por texto)
    user_search = st.sidebar.text_input("Buscar por Usuário", placeholder="Digite o nome do usuário...")
    
    st.sidebar.markdown("---")
    show_ip = st.sidebar.checkbox("🔌 Mostrar IP de Origem na Tabela", value=False)

    # Aplica os filtros
    filtered_df = df.copy()
    
    # Filtro de data
    if isinstance(date_range, tuple) and len(date_range) == 2:
        start_date, end_date = date_range
        filtered_df = filtered_df[
            (filtered_df['datetime_obj'].dt.date >= start_date) & 
            (filtered_df['datetime_obj'].dt.date <= end_date)
        ]
        
    if selected_status:
        filtered_df = filtered_df[filtered_df['status'].isin(selected_status)]
    if selected_tags:
        filtered_df = filtered_df[filtered_df['tag'].isin(selected_tags)]
    if selected_locs:
        filtered_df = filtered_df[filtered_df['localidade_fisica'].isin(selected_locs)]
    if selected_cities:
        filtered_df = filtered_df[filtered_df['cidade_predio'].isin(selected_cities)]
    if selected_units:
        filtered_df = filtered_df[filtered_df['unidade'].isin(selected_units)]
    if selected_bases:
        filtered_df = filtered_df[filtered_df['base'].isin(selected_bases)]
    if user_search:
        filtered_df = filtered_df[filtered_df['usuario'].str.contains(user_search, case=False, na=False)]
        
    # Exibe métricas
    col1, col2, col3 = st.columns(3)
    col1.metric("Total de Chamados", len(filtered_df))
    col2.metric("Abertos", len(filtered_df[filtered_df['status'] == 'Aberto']))
    col3.metric("Fechados", len(filtered_df[filtered_df['status'] == 'Fechado']))
    
    st.write("---")
    
    @st.dialog("Detalhes do Chamado", width="large")
    def show_ticket_details(row):
        # Cabeçalho do chamado
        st.subheader(f"🎫 Chamado #{row['id']}")
        
        # Só exibe o Título se ele não estiver vazio ou nulo
        title = str(row.get('titulo', '')).strip()
        if title and title.lower() not in ["none", "nan", "null", ""]:
            st.info(f"**Título:** {title}")
            
        st.markdown("---")
        
        # Cria as duas colunas para otimizar espaço vertical
        col1, col2 = st.columns(2)
        
        with col1:
            st.markdown("### 👤 Informações do Usuário")
            # Formata o Usuário com o sAMAccountName (id_cliente) se disponível
            user_display = str(row['usuario'])
            client_id = str(row.get('id_cliente', '')).strip()
            if client_id and client_id.lower() not in ["none", "nan", "null", ""]:
                user_display += f" ({client_id})"
                
            st.markdown(f"**Usuário:** {user_display}")
            st.markdown(f"**Localidade:** {row['localidade_fisica']}")
            st.markdown(f"**Base de Origem:** `{row['base']}`")
            st.markdown(f"**IP de Origem:** `{row.get('ip_origem') or 'N/A'}`")
            
        with col2:
            st.markdown("### ⚙️ Classificação & Status")
            # Exibição da TAG com destaque colorido premium
            st.markdown(f"**TAG Inteligente:** :blue-background[{row['tag']}]")
            
            # Alteração de status
            status_options = ["Aberto", "Fechado"]
            current_idx = status_options.index(row['status']) if row['status'] in status_options else 0
            new_status = st.selectbox("Status", status_options, index=current_idx, key="status_select_modal")
            
            if new_status != row['status']:
                if st.button("💾 Salvar Alteração de Status", key="save_status_btn_modal"):
                    conn = sqlite3.connect(DB_PATH)
                    cursor = conn.cursor()
                    now = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
                    cursor.execute("""
                    UPDATE chamados 
                    SET status = ?, data_atualizacao = ?
                    WHERE id = ?
                    """, (new_status, now, row['id']))
                    conn.commit()
                    conn.close()
                    st.success(f"Status atualizado para {new_status}!")
                    st.rerun()
                    
        st.markdown("---")
        
        # Accordion para a Descrição
        with st.expander(f"📝 #1 - {row['Data Formatada']} (Descrição)", expanded=True):
            st.text(row['descricao'])
            
        if st.button("Fechar", key="close_modal_btn"):
            st.rerun()

    # Colunas para exibir
    cols_to_show = [
        'id', 'status', 'tag', 'localidade_fisica', 
        'cidade_predio', 'unidade', 'usuario', 'datetime_obj', 'base'
    ]
    if show_ip:
        cols_to_show.insert(8, 'ip_origem') # Insere antes da base
    
    st.subheader("Lista de Chamados")
    st.write("Dica: Clique no **checkbox (caixinha de seleção)** no início de qualquer linha na tabela abaixo para abrir os Detalhes e Descrição no Modal.")
    
    # Controle de estado para evitar loop do modal
    if "last_selected" not in st.session_state:
        st.session_state["last_selected"] = None
    
    # Configuramos o st.dataframe com seleção nativa (Altamente compatível)
    selection_event = st.dataframe(
        filtered_df[cols_to_show],
        column_config={
            "id": st.column_config.TextColumn("Chamado #"),
            "status": st.column_config.TextColumn("Status"),
            "tag": st.column_config.TextColumn("TAG"),
            "localidade_fisica": st.column_config.TextColumn("Localidade Física"),
            "cidade_predio": st.column_config.TextColumn("Cidade - Prédio"),
            "unidade": st.column_config.TextColumn("Unidade"),
            "usuario": st.column_config.TextColumn("Usuário"),
            "datetime_obj": st.column_config.DatetimeColumn("Data Criação", format="DD/MM/YYYY HH:mm:ss"),
            "ip_origem": st.column_config.TextColumn("IP de Origem"),
            "base": st.column_config.TextColumn("Base"),
        },
        hide_index=True,
        use_container_width=True,
        on_select="rerun",
        selection_mode="single-row"
    )
    
    # Lógica para exibir o Modal baseado na seleção da linha
    selected_rows = selection_event.selection.rows if hasattr(selection_event, "selection") else []
    
    if selected_rows:
        current_selected = selected_rows[0]
        if st.session_state["last_selected"] != current_selected:
            st.session_state["last_selected"] = current_selected
            row_data = filtered_df.iloc[current_selected]
            show_ticket_details(row_data)
    else:
        st.session_state["last_selected"] = None
