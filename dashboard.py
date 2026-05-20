import sys
import asyncio

# Silencia o aviso WinError 10054 (Connection Reset) comum no Windows asyncio
if sys.platform == 'win32':
    try:
        asyncio.set_event_loop_policy(asyncio.WindowsSelectorEventLoopPolicy())
    except:
        pass

import streamlit as st
import streamlit.components.v1 as components
import pandas as pd
import sqlite3
from pathlib import Path
from datetime import datetime
import locale

# Declara o componente customizado do Select Multiple nativo com clique-e-arraste
custom_select_dir = Path(__file__).parent / "custom_select_component"
custom_select = components.declare_component("custom_select", path=str(custom_select_dir))

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
    /* Borda fina vermelha e glow suave para o modal (st.dialog) */
    div[data-testid="stDialog"] > div:first-child,
    div[role="dialog"] {
        border: 2px solid #ff4b4b !important;
        box-shadow: 0 0 15px rgba(255, 75, 75, 0.3) !important;
        border-radius: 8px !important;
    }
    </style>
""", unsafe_allow_html=True)

st.title("📊 Painel de Chamados Centralizado")
st.write("Visualize e interaja com os chamados do OTRS e CitSmart.")

DB_PATH = Path("chamados.db")

# Cores oficiais das TAGs para uso unificado na tabela e no modal
TAG_COLORS = {
    "BACKUP": "#dd5358",
    "EVENTO": "#ce66ce",
    "FORMATAÇÃO": "#d38a62",
    "GARANTIA": "#518bbb",
    "IMPRESSORA": "#C6EFCE",
    "INSTALAÇÃO HARDWARE": "#FCE4D6",
    "INSTALAÇÃO SOFTWARE": "#86BEEE",
    "MANUTENÇÃO": "#E9CF69",
    "MONITOR": "#cbdd6f",
    "MUDANÇA": "#21ffe0",
    "PREPARAÇÃO COMPUTADORES": "#f09c72",
    "REDE": "#B7F391",
    "SOLICITAÇÃO SSD": "#f5a89b",
    "SUPORTE": "#FFE699",
    "TELEFONIA FIXA": "#e273a1",
    "VIAGEM": "#61e7c6",
    "VISTORIA CPDS": "#b2740e",
}

def load_data():
    if not DB_PATH.exists():
        return pd.DataFrame()
    conn = sqlite3.connect(DB_PATH)
    df = pd.read_sql_query("SELECT * FROM chamados", conn)
    conn.close()
    
    # Limpa " - Sede" de forma inteligente na coluna de exibição da Localidade Física
    if 'localidade_fisica' in df.columns:
        import re
        df['localidade_fisica'] = df['localidade_fisica'].apply(
            lambda x: re.sub(r'\s*-\s*Sede\b', '', str(x), flags=re.IGNORECASE).strip() if pd.notna(x) else x
        )
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
    
    # Componente Customizado Select Multiple nativo com suporte a clique e arraste (drag-to-select) e Shift+Click
    loc_options = sorted(list(df['localidade_fisica'].dropna().unique()))
    
    st.sidebar.markdown("---")
    st.sidebar.subheader("📍 Seleção de Localidades")
    st.sidebar.write("Arraste o mouse sobre as opções ou segure **Shift** para selecionar múltiplas de uma vez:")
    
    # Executa o componente customizado passando as opções e recuperando a lista selecionada
    with st.sidebar:
        selected_locs = custom_select(
            options=loc_options,
            default=[],
            key="custom_loc_selection"
        )
    if selected_locs is None:
        selected_locs = []
    
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
        # Cabeçalho do chamado unificado em um Expander aberto por padrão para economizar espaço se necessário
        title = str(row.get('titulo', '')).strip()
        if title and title.lower() not in ["none", "nan", "null", ""]:
            header_text = f"🎫 Chamado #{row['id']} – {title}"
        else:
            header_text = f"🎫 Chamado #{row['id']}"
            
        with st.expander(header_text, expanded=True):
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
                # Exibição da TAG com destaque colorido premium baseado nas cores oficiais
                tag_name = str(row['tag']).upper().strip()
                bg_color = TAG_COLORS.get(tag_name, "#262730")
                
                # Calcula contraste excelente calculando a luminância da cor de fundo
                hex_color = bg_color.lstrip('#')
                try:
                    r, g, b = int(hex_color[0:2], 16), int(hex_color[2:4], 16), int(hex_color[4:6], 16)
                    luminance = (0.299 * r + 0.587 * g + 0.114 * b) / 255
                    text_color = "#ffffff" if luminance < 0.6 else "#212529"
                except:
                    text_color = "#ffffff"
                    
                tag_html = f'<span style="background-color: {bg_color}; color: {text_color}; padding: 3px 8px; border-radius: 4px; font-weight: bold; font-family: inherit; font-size: 13px;">{row["tag"]}</span>'
                st.markdown(f"**TAG Inteligente:** {tag_html}", unsafe_allow_html=True)
                
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
                        
                # Botão para abrir o chamado original
                link_url = row.get('link')
                if link_url:
                    st.markdown("---")
                    st.link_button("🔗 Abrir Chamado Original", link_url, width="stretch")
        
        # Accordion para a Descrição
        with st.expander(f"📝 #1 - {row['Data Formatada']} (Descrição)", expanded=True):
            st.text(row['descricao'])
            
        # Comentários / Notas históricas
        from database import get_comments_by_ticket
        comments = get_comments_by_ticket(row['id'])
        if comments:
            st.markdown("### 💬 Histórico de Notas e Acompanhamentos")
            # Exibe cada comentário em um expander elegante
            for i, c in enumerate(comments, start=2):
                header = f"🕒 #{i} – {c['data']} – por {c['autor']}"
                with st.expander(header):
                    st.text(c['texto'])
            
        if st.button("Fechar", key="close_modal_btn"):
            st.rerun()

    # Colunas para exibir por padrão (Ordem padrão)
    cols_to_show = [
        'id', 'status', 'tag', 'localidade_fisica', 
        'cidade_predio', 'unidade', 'usuario', 'datetime_obj', 'base'
    ]
        
    # Inclui colunas necessárias para geração de links e para estarem disponíveis no picker (como ip_origem)
    cols_to_generate = list(cols_to_show)
    if 'link' not in cols_to_generate and 'link' in filtered_df.columns:
        cols_to_generate.append('link')
    if 'ip_origem' not in cols_to_generate and 'ip_origem' in filtered_df.columns:
        cols_to_generate.append('ip_origem')
        
    df_display = filtered_df[cols_to_generate].copy()
    
    def format_id_link(row):
        cid = str(row['id']).strip()
        link = str(row.get('link', '')).strip() if 'link' in row else ''
        if not link or link.lower() in ["none", "nan", "null", ""]:
            # Fallbacks seguros
            if row['base'] == 'CitSmart':
                link = f"https://suporte.mpms.mp.br/citsmart/pages/serviceRequestIncident/serviceRequestIncident.load?iframe=true&language=pt-BR#/request?idRequest={cid}"
            else:
                link = "https://central.mpms.mp.br/otrs/index.pl"
        return f"{link}#id:{cid}"

    df_display['id'] = df_display.apply(format_id_link, axis=1)
    
    # As colunas finais passadas incluem ip_origem para estar disponível no picker
    cols_for_dataframe = list(cols_to_show)
    if 'ip_origem' in df_display.columns:
        cols_for_dataframe.append('ip_origem')
        
    df_final_display = df_display[cols_for_dataframe]
    
    st.subheader("Lista de Chamados")
    st.write("Dica: Clique no **checkbox (caixinha de seleção)** no início de qualquer linha na tabela abaixo para abrir os Detalhes e Descrição no Modal.")
    
    # Controle de estado para evitar loop do modal
    if "last_selected" not in st.session_state:
        st.session_state["last_selected"] = None
    
    # Função para colorir as linhas do DataFrame de acordo com as TAGs e suas cores oficiais do Excel
    def style_dataframe(row):
        tag = str(row.get('tag', '')).upper().strip()
        bg_color = TAG_COLORS.get(tag, "")
        if bg_color:
            # Garante contraste excelente calculando a luminância da cor de fundo
            hex_color = bg_color.lstrip('#')
            r, g, b = int(hex_color[0:2], 16), int(hex_color[2:4], 16), int(hex_color[4:6], 16)
            luminance = (0.299 * r + 0.587 * g + 0.114 * b) / 255
            text_color = "#ffffff" if luminance < 0.6 else "#212529"
            style = f"background-color: {bg_color}; color: {text_color};"
        else:
            style = ""
            
        return [style] * len(row)

    # Configuramos o st.dataframe com seleção nativa e estilização de cores (Altamente compatível)
    selection_event = st.dataframe(
        df_final_display.style.apply(style_dataframe, axis=1),
        column_order=cols_to_show, # Especifica quais colunas aparecem por padrão (oculta ip_origem)
        column_config={
            "id": st.column_config.LinkColumn("Chamado #", display_text=r".*#id:(.*)"),
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
        width="stretch",
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

    # NOVO: Seção para geração rápida de resumo para WhatsApp com interpretação de IA e problema completo
    st.markdown("---")
    st.subheader("📲 Compartilhar Fila por WhatsApp")
    with st.expander("💬 Gerar Resumo Formatado (Pronto para copiar e enviar)", expanded=False):
        if filtered_df.empty:
            st.info("Nenhum chamado na fila filtrada.")
        else:
            def get_ai_diagnostico(tag, desc):
                tag = str(tag).upper().strip()
                desc = str(desc).strip()
                
                # Limpeza de saudações iniciais para extrair sintoma puro
                import re
                clean_desc = re.sub(r'^(bom dia|boa tarde|boa noite|ola|prezados|favor|solicito|gostaria de)\b.*?\n', '', desc, flags=re.IGNORECASE)
                clean_desc = clean_desc.strip()
                
                sentences = re.split(r'[.!?\n]', clean_desc)
                first_sentence = ""
                for s in sentences:
                    s = s.strip()
                    if len(s) > 10:
                        first_sentence = s
                        break
                if not first_sentence:
                    first_sentence = desc[:100]
                    if len(desc) > 100:
                        first_sentence += "..."
                        
                diagnosticos = {
                    "BACKUP": "Cópia de segurança ou restauração de arquivos pendente.",
                    "EVENTO": "Suporte técnico para eventos ou solenidades institucionais.",
                    "FORMATAÇÃO": "Computador com lentidão extrema/travamento exigindo formatação e reinstalação de OS.",
                    "GARANTIA": "Defeito físico de fábrica em equipamento que exige acionamento de suporte terceirizado.",
                    "IMPRESSORA": "Instabilidade na fila de impressão local, papel atolado ou configuração de nova impressora de rede.",
                    "INSTALAÇÃO HARDWARE": "Necessidade de substituição física ou acréscimo de componente de hardware na máquina.",
                    "INSTALAÇÃO SOFTWARE": "Instalação, licenciamento ou atualização corretiva de programas corporativos.",
                    "MANUTENÇÃO": "Necessidade de intervenção mecânica/elétrica, limpeza interna ou reaperto de conexões físicas.",
                    "MONITOR": "Sem sinal de vídeo, tela preta, piscando ou distorcendo imagens de saída.",
                    "MUDANÇA": "Deslocamento físico completo de equipamentos de informática entre salas ou comarcas.",
                    "PREPARAÇÃO COMPUTADORES": "Configuração inicial de máquinas novas e perfis de rede para novos servidores.",
                    "REDE": "Ausência total de internet, falha de rede física ou lentidão no tráfego de dados locais.",
                    "SOLICITAÇÃO SSD": "Melhoria de desempenho físico de máquina lenta via substituição por disco de estado sólido (SSD).",
                    "SUPORTE": "Instruções de uso básico ou esclarecimento de dúvidas técnicas em sistemas internos.",
                    "TELEFONIA FIXA": "Aparelho de telefone mudo, ramal com ruídos/chiado ou necessidade de transferência de ramal.",
                    "VIAGEM": "Deslocamento programado da equipe STI para atendimento em promotoria regional externa.",
                    "VISTORIA CPDS": "Check-up preventivo completo nos servidores e no centro de processamento de dados local."
                }
                
                diag = diagnosticos.get(tag, "Análise e resolução de ticket técnico STI.")
                return f"🧠 *Possível Problema:* {diag}\n🩺 *Sintoma:* _{first_sentence}_"

            from database import get_comments_by_ticket

            lines = []
            lines.append("📋 *LISTA DE CHAMADOS STI - MPMS* 📋\n")
            for _, row in filtered_df.iterrows():
                cid = str(row['id']).strip()
                link = str(row.get('link', '')).strip()
                if not link or link.lower() in ["none", "nan", "null", ""]:
                    if row['base'] == 'CitSmart':
                        link = f"https://suporte.mpms.mp.br/citsmart/pages/serviceRequestIncident/serviceRequestIncident.load?iframe=true&language=pt-BR#/request?idRequest={cid}"
                    else:
                        link = "https://central.mpms.mp.br/otrs/index.pl"
                
                user = str(row['usuario'])
                loc = str(row['localidade_fisica'])
                tag = str(row['tag'])
                desc = str(row['descricao']).strip()
                
                # Recupera os comentários históricos do banco para enviar junto
                comments_list = get_comments_by_ticket(row['id'])
                comments_text = ""
                if comments_list:
                    comments_text = "💬 *Histórico de Acompanhamento:*"
                    for i, c in enumerate(comments_list, start=1):
                        comments_text += f"\n  • #{i} [{c['data']}] – {c['autor']}: {c['texto']}"
                
                # Gera diagnóstico inteligente
                diagnostico_ia = get_ai_diagnostico(tag, desc)
                
                lines.append(f"🎫 *Chamado #{cid}* ({row['base']})")
                lines.append(f"👤 *Usuário:* {user}")
                lines.append(f"📍 *Local:* {loc}")
                lines.append(f"🏷️ *TAG:* {tag}")
                lines.append(f"{diagnostico_ia}")
                lines.append(f"📝 *Problema Completo:*")
                lines.append(f"{desc}")
                if comments_text:
                    lines.append(comments_text)
                lines.append(f"🔗 *Link Direto:* {link}")
                lines.append("--------------------------------------------------")
            
            whats_text = "\n".join(lines)
            st.write("Dica: Use o botão de **copiar** no canto superior direito do bloco de código abaixo:")
            st.code(whats_text, language="text")
