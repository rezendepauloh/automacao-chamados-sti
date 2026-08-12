import sys
from pathlib import Path

root_dir = Path(__file__).parent.parent.parent

import asyncio
import os
import re
from datetime import datetime
import pandas as pd
import sqlite3
import streamlit as st
import streamlit.components.v1 as components

from src.components.status_banner import check_orquestrador_running, read_last_log_lines, render_log_expander
from src.database import load_data, get_comments_by_ticket, update_ticket_location_details, update_ticket_tag, update_ticket_andamento, update_ticket_title
from src.components.pagination import (
    render_items_per_page_selector,
    paginate_items,
    render_pagination_controls
)

DB_PATH = root_dir / "chamados.db"

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

@st.cache_resource
def load_spacy_model():
    """Carrega o modelo spaCy local em português, com fallback caso falhe."""
    import spacy
    try:
        return spacy.load("pt_core_news_sm")
    except Exception:
        return None

@st.cache_data(show_spinner=False)
def summarize_ticket_locally(description: str, comments: str, max_sentences: int = 2) -> str:
    """
    Resume o chamado técnico localmente usando Processamento de Linguagem Natural (spaCy).
    """
    from collections import Counter
    from src.config import clean_otrs_description
    
    description = clean_otrs_description(description)
    
    def clean_text(t: str) -> str:
        if not t:
            return ""
        t = re.sub(
            r'^\s*(?:prezados?|prezadas?|caros?|caras?|olá|ola|bom\s+dia|boa\s+tarde|boa\s+noite|prezada\s+equipe|prezada\s+sti)\b(?:[^\n\.\?]*[\n\.,\?])?',
            '', t, flags=re.IGNORECASE
        )
        t = re.sub(
            r'^\s*(?:tudo\s+bem\??|espero\s+que\s+esteja\s+tudo\s+bem\??|espero\s+que\s+sim\??)',
            '', t, flags=re.IGNORECASE
        )
        t = re.sub(
            r'\b(?:gostaria\s+de\s+|venho\s+(?:por\s+meio\s+deste\s+)?|favor\s+|por\s+gentileza\s+|gentileza\s+)\b',
            '', t, flags=re.IGNORECASE
        )
        t = re.sub(
            r'\b(?:solicito\s+providências\s+para\s+|solicito\s+(?:a|o|que|os|as)?\s+|encaminho\s+para\s+providências\s*(?:[ao]s?|para|de)?\s+|encaminho\s+para\s+|segue\s+para\s+|segue\s+o\s+chamado\s+(?:para\s+)?)\b',
            '', t, flags=re.IGNORECASE
        )
        t = re.sub(
            r'\b(?:conforme|como|conforme\s+mostra\s+a|ver|veja)\s+(?:imagem\s+)?(?:em\s+)?anexo\b',
            '', t, flags=re.IGNORECASE
        )
        t = re.sub(
            r'\b(?:segue[m]?\s+)?(?:em\s+)?anexo\b',
            '', t, flags=re.IGNORECASE
        )
        t = re.sub(
            r'\b(?:fico|ficamos)\s+(?:à|a)\s+disposição\s+para\s+(?:eventuais|quaisquer)\s+(?:esclarecimentos|dúvidas)\b\.?',
            '', t, flags=re.IGNORECASE
        )
        t = re.sub(
            r'\b(?:desde\s+já\s+)?agradeço[s]?\b\.?',
            '', t, flags=re.IGNORECASE
        )
        t = re.sub(
            r'\b(?:atenciosamente|grato|obrigado|fico\s+no\s+aguardo|aguardo\s+retorno|sem\s+mais)\b\.?',
            '', t, flags=re.IGNORECASE
        )
        t = re.sub(r'\s+', ' ', t)
        t = re.sub(r'^\s*[,\.\-\:\/]+\s*', '', t)
        return t.strip()

    desc_clean = clean_text(str(description))
    
    comments_clean_list = []
    if comments:
        for line in str(comments).split('\n'):
            line_payload = re.sub(r'^[-\s]*[\d/:\s\[\]\-\#\.]+(?:[\w\s\(\)]+)?:\s*', '', line).strip()
            line_clean = clean_text(line_payload)
            if line_clean and len(line_clean) > 8:
                comments_clean_list.append(line_clean)
                
    text_parts = []
    if desc_clean:
        text_parts.append(desc_clean)
    if comments_clean_list:
        text_parts.append(" ".join(comments_clean_list))
        
    combined_text = " ".join(text_parts).strip()
    
    if not combined_text:
        return "Sem descrição detalhada."
        
    nlp = load_spacy_model()
    
    if nlp is None:
        sentences = [s.strip() for s in combined_text.split('.') if len(s.strip()) > 8]
        if sentences:
            res = ". ".join(sentences[:max_sentences])
            if not res.endswith('.'):
                res += '.'
            return res
        return combined_text[:140] + "..." if len(combined_text) > 140 else combined_text

    doc = nlp(combined_text)
    
    TECHNICAL_BOOST = {
        "ssd", "hd", "windows", "formatação", "formatar", "lentidão", "travamento", "travando",
        "impressora", "imprimir", "rede", "conexão", "erro", "falha", "sistema", "configurar",
        "configuração", "instalação", "instalar", "senha", "usuário", "computador", "máquina",
        "notebook", "monitor", "teclado", "mouse", "backup", "servidor", "internet", "cabo",
        "wi-fi", "wifi", "login", "acesso", "workstation", "driver", "inicialização", "boot",
        "perfil", "outlook", "email", "e-mail", "toner", "cartucho", "suporte", "atualizar",
        "atualização", "office", "word", "excel", "pasta", "rede", "compartilhamento"
    }
    
    keywords = []
    for token in doc:
        if token.is_stop or token.is_punct or token.is_space:
            continue
        if token.pos_ in ["NOUN", "VERB", "ADJ", "PROPN"]:
            keywords.append(token.text.lower())
            
    if not keywords:
        sentences = list(doc.sents)
        return " ".join([s.text.strip() for s in sentences[:max_sentences]])
        
    word_freq = Counter(keywords)
    max_freq = max(word_freq.values())
    for word in word_freq:
        word_freq[word] = word_freq[word] / max_freq
        
    sent_scores = {}
    sentences = list(doc.sents)
    
    for idx, sent in enumerate(sentences):
        words = [t for t in sent if not t.is_punct and not t.is_space]
        if len(words) < 3:
            continue
            
        score = 0
        for token in sent:
            word_lower = token.text.lower()
            if word_lower in word_freq:
                score += word_freq[word_lower]
            if word_lower in TECHNICAL_BOOST:
                score += 3.0
                
        word_count = len(words)
        if 8 <= word_count <= 25:
            score *= 1.3
        elif word_count > 30:
            score *= 0.6
        elif word_count < 6:
            score *= 0.7
            
        if idx == 0:
            score += 2.0
            
        if idx == len(sentences) - 1 and len(sentences) > 1:
            score += 1.0
            
        sent_scores[sent] = score
        
    if not sent_scores:
        res = " ".join([s.text.strip() for s in sentences[:max_sentences]])
        return res
        
    sorted_sents = sorted(sent_scores.keys(), key=lambda x: sent_scores[x], reverse=True)
    top_sents = sorted_sents[:max_sentences]
    top_sents = sorted(top_sents, key=lambda x: x.start)
    
    formatted_sentences = []
    for s in top_sents:
        sent_text = s.text.strip()
        if not sent_text:
            continue
        sent_text = sent_text[0].upper() + sent_text[1:]
        sent_text = re.sub(r'^[\s,\.\-\:\/]+', '', sent_text)
        if not sent_text.endswith(('.', '!', '?')):
            sent_text += '.'
        formatted_sentences.append(sent_text)
        
    summary = " ".join(formatted_sentences)
    return summary


def render_chamados_page():
    """Renderiza a página principal do Painel de Chamados."""
    col_title, col_btn = st.columns([3, 1])
    with col_title:
        st.title("📊 Painel de Chamados Centralizado")
        st.write("Visualize e interaja com os chamados do OTRS e CitSmart.")

    robo_ativo = check_orquestrador_running()

    with col_btn:
        st.markdown("<div style='height: 15px;'></div>", unsafe_allow_html=True)
        if robo_ativo:
            st.button("🤖 Robô em Execução...", use_container_width=True, disabled=True)
        else:
            run_orquestrador = st.button(
                "🔄 Atualizar Chamados", 
                use_container_width=True, 
                help="Executa o orquestrador completo em segundo plano.",
                type="primary"
            )
            if run_orquestrador:
                import subprocess
                import time
                subprocess.Popen([sys.executable, "orquestrador.py"])
                time.sleep(0.8)
                st.toast("🚀 Robô iniciado em segundo plano!", icon="🤖")
                st.cache_data.clear()
                st.rerun()

    render_log_expander(
        "🤖 Robô Rodando em Segundo Plano – Acompanhar Progresso",
        robo_ativo,
        read_last_log_lines,
        check_orquestrador_running,
        "O robô está coletando novos chamados e classificando com IA neste momento. Você pode continuar usando o painel normalmente!"
    )

    df = load_data()

    if df.empty:
        st.warning("Nenhum dado encontrado no banco de dados. Execute o orquestrador primeiro!")
        return

    # Conversão de data garantindo formato brasileiro e ISO correto
    def parse_dates_safely(val):
        if pd.isna(val) or not str(val).strip():
            return pd.NaT
        s = str(val).strip()
        # Se estiver no padrão ISO YYYY-MM-DD HH:MM:SS
        if re.match(r'^\d{4}-\d{2}-\d{2}', s):
            return pd.to_datetime(s, errors='coerce')
        # Se estiver no formato brasileiro DD/MM/YYYY HH:MM:SS
        return pd.to_datetime(s, dayfirst=True, errors='coerce')

    df['datetime_obj'] = df['data_criacao'].apply(parse_dates_safely)
    df['Data Formatada'] = df['datetime_obj'].dt.strftime('%d/%m/%Y %H:%M:%S')
    df['Data Formatada'] = df['Data Formatada'].fillna(df['data_criacao'])

    # Ordenação padrão decrescente pela data correta
    df = df.sort_values(by='datetime_obj', ascending=False)

    min_date = df['datetime_obj'].dropna().min().date() if not df['datetime_obj'].dropna().empty else datetime.now().date()
    max_date = df['datetime_obj'].dropna().max().date() if not df['datetime_obj'].dropna().empty else datetime.now().date()
    
    if "f_date_range" not in st.session_state:
        st.session_state["f_date_range"] = (min_date, max_date)
    if "f_status" not in st.session_state:
        st.session_state["f_status"] = []
    if "f_tags" not in st.session_state:
        st.session_state["f_tags"] = []
    if "custom_loc_selection" not in st.session_state:
        st.session_state["custom_loc_selection"] = []
    if "f_cities" not in st.session_state:
        st.session_state["f_cities"] = []
    if "f_units" not in st.session_state:
        st.session_state["f_units"] = []
    if "f_bases" not in st.session_state:
        st.session_state["f_bases"] = []
    if "f_user" not in st.session_state:
        st.session_state["f_user"] = ""
    if "f_mode" not in st.session_state:
        st.session_state["f_mode"] = "🟢 Manter Selecionados"
    if "f_ticket_ids" not in st.session_state:
        st.session_state["f_ticket_ids"] = []

    def get_filtered_options(col_name: str) -> list:
        temp_df = df.copy()
        is_exclude_mode = (st.session_state.get("f_mode") == "🔴 Ocultar Selecionados")
        
        dr = st.session_state.get("f_date_range", (min_date, max_date))
        if isinstance(dr, tuple) and len(dr) == 2:
            start_date, end_date = dr
            temp_df = temp_df[
                (temp_df['datetime_obj'].dt.date >= start_date) & 
                (temp_df['datetime_obj'].dt.date <= end_date)
            ]
            
        if col_name != 'status' and st.session_state.get("f_status"):
            if is_exclude_mode:
                temp_df = temp_df[~temp_df['status'].isin(st.session_state["f_status"])]
            else:
                temp_df = temp_df[temp_df['status'].isin(st.session_state["f_status"])]
            
        if col_name != 'tag' and st.session_state.get("f_tags"):
            if is_exclude_mode:
                temp_df = temp_df[~temp_df['tag'].isin(st.session_state["f_tags"])]
            else:
                temp_df = temp_df[temp_df['tag'].isin(st.session_state["f_tags"])]
            
        if col_name != 'localidade_fisica' and st.session_state.get("custom_loc_selection"):
            if is_exclude_mode:
                temp_df = temp_df[~temp_df['localidade_fisica'].isin(st.session_state["custom_loc_selection"])]
            else:
                temp_df = temp_df[temp_df['localidade_fisica'].isin(st.session_state["custom_loc_selection"])]
            
        if col_name != 'cidade_predio' and st.session_state.get("f_cities"):
            if is_exclude_mode:
                temp_df = temp_df[~temp_df['cidade_predio'].isin(st.session_state["f_cities"])]
            else:
                temp_df = temp_df[temp_df['cidade_predio'].isin(st.session_state["f_cities"])]
            
        if col_name != 'unidade' and st.session_state.get("f_units"):
            if is_exclude_mode:
                temp_df = temp_df[~temp_df['unidade'].isin(st.session_state["f_units"])]
            else:
                temp_df = temp_df[temp_df['unidade'].isin(st.session_state["f_units"])]
            
        if col_name != 'base' and st.session_state.get("f_bases"):
            if is_exclude_mode:
                temp_df = temp_df[~temp_df['base'].isin(st.session_state["f_bases"])]
            else:
                temp_df = temp_df[temp_df['base'].isin(st.session_state["f_bases"])]
            
        if col_name != 'usuario' and st.session_state.get("f_user"):
            if is_exclude_mode:
                temp_df = temp_df[~temp_df['usuario'].str.contains(st.session_state["f_user"], case=False, na=False)]
            else:
                temp_df = temp_df[temp_df['usuario'].str.contains(st.session_state["f_user"], case=False, na=False)]
            
        options = sorted(list(temp_df[col_name].dropna().unique()))
        
        key_map = {
            'status': 'f_status',
            'tag': 'f_tags',
            'localidade_fisica': 'custom_loc_selection',
            'cidade_predio': 'f_cities',
            'unidade': 'f_units',
            'base': 'f_bases',
            'usuario': 'f_user'
        }
        session_key = key_map.get(col_name)
        if session_key:
            current_selection = st.session_state.get(session_key, [])
            if current_selection:
                if isinstance(current_selection, list):
                    for val in current_selection:
                        if val not in options:
                            options.append(val)
                elif current_selection not in options:
                    options.append(current_selection)
                    
        return options

    st.sidebar.markdown("### 🔍 Filtros de Chamados")
    
    filter_mode = st.sidebar.radio(
        "Modo de Filtragem:",
        options=["🟢 Manter Selecionados", "🔴 Ocultar Selecionados"],
        key="f_mode",
        horizontal=True,
        help="🟢 Manter: Mostra apenas os itens escolhidos.\n🔴 Ocultar: Exibe tudo EXCETO os itens escolhidos."
    )
    
    status_options = get_filtered_options('status')
    tag_options = get_filtered_options('tag')
    city_options = get_filtered_options('cidade_predio')
    unit_options = get_filtered_options('unidade')
    base_options = get_filtered_options('base')

    col_btn_sel1, col_btn_sel2 = st.sidebar.columns(2)
    with col_btn_sel1:
        if st.button("☑️ Marcar Todos", use_container_width=True, help="Seleciona todas as opções"):
            st.session_state["f_status"] = list(status_options)
            st.session_state["f_tags"] = list(tag_options)
            st.session_state["f_cities"] = list(city_options)
            st.session_state["f_units"] = list(unit_options)
            st.session_state["f_bases"] = list(base_options)
            st.rerun()
    with col_btn_sel2:
        if st.button("🧹 Limpar Tudo", use_container_width=True, help="Limpa todos os filtros ativos"):
            st.session_state["f_date_range"] = (min_date, max_date)
            st.session_state["f_status"] = []
            st.session_state["f_tags"] = []
            st.session_state["custom_loc_selection"] = []
            st.session_state["f_cities"] = []
            st.session_state["f_units"] = []
            st.session_state["f_bases"] = []
            st.session_state["f_user"] = ""
            st.session_state["f_ticket_ids"] = []
            st.rerun()

    st.sidebar.markdown("---")

    date_range = st.sidebar.date_input(
        "Intervalo de Datas",
        value=st.session_state["f_date_range"],
        min_value=min_date,
        max_value=max_date,
        format="DD/MM/YYYY",
        key="f_date_range"
    )
    
    selected_status = st.sidebar.multiselect(
        "Status", 
        options=status_options, 
        key="f_status", 
        placeholder="Escolha as opções..."
    )
    
    selected_tags = st.sidebar.multiselect(
        "TAG (Categoria de IA)", 
        options=tag_options, 
        key="f_tags", 
        placeholder="Escolha as opções..."
    )
    
    loc_options = get_filtered_options('localidade_fisica')
    
    st.sidebar.markdown("---")
    st.sidebar.subheader("📍 Seleção de Localidades")
    
    selected_locs = st.sidebar.multiselect(
        "Localidade Física",
        options=loc_options,
        default=st.session_state.get("custom_loc_selection", []),
        key="custom_loc_selection",
        placeholder="Escolha as localidades..."
    )
    
    selected_cities = st.sidebar.multiselect(
        "Cidade - Prédio", 
        options=city_options, 
        key="f_cities", 
        placeholder="Escolha as opções..."
    )
    
    selected_units = st.sidebar.multiselect(
        "Unidade", 
        options=unit_options, 
        key="f_units", 
        placeholder="Escolha as opções..."
    )
    
    selected_bases = st.sidebar.multiselect(
        "Base de Origem", 
        options=base_options, 
        key="f_bases", 
        placeholder="Escolha as opções..."
    )
    
    user_search = st.sidebar.text_input(
        "Buscar por Usuário", 
        key="f_user", 
        placeholder="Digite o nome do usuário..."
    )

    filtered_df = df.copy()
    is_exclude_mode = (filter_mode == "🔴 Ocultar Selecionados")
    
    if isinstance(date_range, tuple) and len(date_range) == 2:
        start_date, end_date = date_range
        filtered_df = filtered_df[
            (filtered_df['datetime_obj'].dt.date >= start_date) & 
            (filtered_df['datetime_obj'].dt.date <= end_date)
        ]
        
    if selected_status:
        if is_exclude_mode:
            filtered_df = filtered_df[~filtered_df['status'].isin(selected_status)]
        else:
            filtered_df = filtered_df[filtered_df['status'].isin(selected_status)]

    if selected_tags:
        if is_exclude_mode:
            filtered_df = filtered_df[~filtered_df['tag'].isin(selected_tags)]
        else:
            filtered_df = filtered_df[filtered_df['tag'].isin(selected_tags)]

    if selected_locs:
        if is_exclude_mode:
            filtered_df = filtered_df[~filtered_df['localidade_fisica'].isin(selected_locs)]
        else:
            filtered_df = filtered_df[filtered_df['localidade_fisica'].isin(selected_locs)]

    if selected_cities:
        if is_exclude_mode:
            filtered_df = filtered_df[~filtered_df['cidade_predio'].isin(selected_cities)]
        else:
            filtered_df = filtered_df[filtered_df['cidade_predio'].isin(selected_cities)]

    if selected_units:
        if is_exclude_mode:
            filtered_df = filtered_df[~filtered_df['unidade'].isin(selected_units)]
        else:
            filtered_df = filtered_df[filtered_df['unidade'].isin(selected_units)]

    if selected_bases:
        if is_exclude_mode:
            filtered_df = filtered_df[~filtered_df['base'].isin(selected_bases)]
        else:
            filtered_df = filtered_df[filtered_df['base'].isin(selected_bases)]

    if user_search:
        if is_exclude_mode:
            filtered_df = filtered_df[~filtered_df['usuario'].str.contains(user_search, case=False, na=False)]
        else:
            filtered_df = filtered_df[filtered_df['usuario'].str.contains(user_search, case=False, na=False)]

    st.sidebar.markdown("---")
    
    source_df = filtered_df if not filtered_df.empty else df
    df_tickets_list = source_df[['id', 'titulo']].copy()
    ids_clean = df_tickets_list['id'].astype(str).str.replace(r'\.0$', '', regex=True).str.strip()
    titles_clean = df_tickets_list['titulo'].fillna("Sem Título").astype(str).str.strip()
    titles_clean = titles_clean.replace(["nan", "NaN", "None", "null", ""], "Sem Título")
    labels = ids_clean + " - " + titles_clean
    ticket_options = sorted(labels.unique().tolist())
    
    current_sel_tickets = st.session_state.get("f_ticket_ids", [])
    for sel_t in current_sel_tickets:
        if sel_t not in ticket_options:
            ticket_options.append(sel_t)

    selected_tickets = st.sidebar.multiselect(
        "🎫 Chamados Específicos (por ID / Título)",
        options=ticket_options,
        key="f_ticket_ids",
        placeholder="Selecione chamados individuais..."
    )

    items_per_page = render_items_per_page_selector(
        key_prefix="chamados_v5",
        options=[10, 50, 100, "Todos"],
        default_index=2,
        label="📄 Chamados por página:"
    )

    st.sidebar.markdown("---")


    if selected_tickets:
        selected_ids = [t.split(" - ")[0].strip() for t in selected_tickets]
        clean_df_ids = filtered_df['id'].astype(str).str.replace(r'\.0$', '', regex=True).str.strip()
        if is_exclude_mode:
            filtered_df = filtered_df[~clean_df_ids.isin(selected_ids)]
        else:
            filtered_df = filtered_df[clean_df_ids.isin(selected_ids)]
        
    col1, col2, col3 = st.columns(3)
    total_ch = len(filtered_df)
    abertos_ch = len(filtered_df[filtered_df['status'] == 'Aberto']) if 'status' in filtered_df.columns else 0
    fechados_ch = len(filtered_df[filtered_df['status'] == 'Fechado']) if 'status' in filtered_df.columns else 0

    with col1:
        st.markdown(f"""
            <div class="metric-card" style="border-left-color: #3b82f6;">
                <div class="metric-title">TOTAL DE CHAMADOS</div>
                <div class="metric-value">{total_ch}</div>
            </div>
        """, unsafe_allow_html=True)

    with col2:
        st.markdown(f"""
            <div class="metric-card" style="border-left-color: #f59e0b;">
                <div class="metric-title">ABERTOS</div>
                <div class="metric-value" style="color: #f59e0b;">{abertos_ch}</div>
            </div>
        """, unsafe_allow_html=True)

    with col3:
        st.markdown(f"""
            <div class="metric-card" style="border-left-color: #10b981;">
                <div class="metric-title">FECHADOS</div>
                <div class="metric-value" style="color: #10b981;">{fechados_ch}</div>
            </div>
        """, unsafe_allow_html=True)


    
    # ABAS INTERNAS DE VISUALIZAÇÃO COM QUERY PARAMETERS (?subtab=slug)
    CHAMADOS_SUBTAB_MAP = {
        "tabela": "📋 Tabela Geral de Chamados",
        "graficos": "📈 Gráficos & Estatísticas do Painel"
    }
    CHAMADOS_SUBTAB_REVERSE = {v: k for k, v in CHAMADOS_SUBTAB_MAP.items()}

    url_subtab = st.query_params.get("subtab", "tabela")
    default_title = CHAMADOS_SUBTAB_MAP.get(url_subtab, "📋 Tabela Geral de Chamados")
    options = list(CHAMADOS_SUBTAB_MAP.values())
    default_idx = options.index(default_title) if default_title in options else 0

    selected_subtab = st.radio(
        "Navegação do Painel:",
        options=options,
        index=default_idx,
        horizontal=True,
        label_visibility="collapsed",
        key="chamados_subtab_radio"
    )

    new_slug = CHAMADOS_SUBTAB_REVERSE.get(selected_subtab, "tabela")
    if st.query_params.get("subtab") != new_slug:
        st.query_params["subtab"] = new_slug

    st.markdown("<br>", unsafe_allow_html=True)

    if selected_subtab == "📈 Gráficos & Estatísticas do Painel":

        st.subheader("📊 Análise Estatística dos Chamados Filtrados")
        st.write("Visualização consolidada de abertura de chamados por Prédio, Unidade, Categorias (TAGs) e Usuários.")
        st.markdown("<br>", unsafe_allow_html=True)

        if not filtered_df.empty:
            g_col1, g_col2 = st.columns(2)
            with g_col1:
                st.markdown("#### 🏢 Top Prédios / Cidades com Mais Chamados")
                city_counts = filtered_df['cidade_predio'].value_counts().head(10).reset_index()
                city_counts.columns = ['Prédio / Cidade', 'Quantidade']
                st.bar_chart(city_counts.set_index('Prédio / Cidade'), use_container_width=True)

            with g_col2:
                st.markdown("#### 🏛️ Top Unidades / Setores Mais Demandantes")
                unit_counts = filtered_df['unidade'].value_counts().head(10).reset_index()
                unit_counts.columns = ['Unidade / Setor', 'Quantidade']
                st.bar_chart(unit_counts.set_index('Unidade / Setor'), use_container_width=True)

            st.markdown("---")

            g_col3, g_col4 = st.columns(2)
            with g_col3:
                st.markdown("#### 🏷️ Distribuição por Categoria (TAG de IA)")
                tag_counts = filtered_df['tag'].value_counts().reset_index()
                tag_counts.columns = ['Categoria (TAG)', 'Quantidade']
                st.bar_chart(tag_counts.set_index('Categoria (TAG)'), use_container_width=True)

            with g_col4:
                st.markdown("#### 👤 Top Usuários que Mais Abrem Chamados")
                user_counts = filtered_df['usuario'].value_counts().head(10).reset_index()
                user_counts.columns = ['Usuário', 'Quantidade']
                st.bar_chart(user_counts.set_index('Usuário'), use_container_width=True)
                
            st.markdown("---")

            g_col5, g_col6 = st.columns(2)
            with g_col5:
                st.markdown("#### 🔄 Origem dos Chamados (Base)")
                base_counts = filtered_df['base'].value_counts().reset_index()
                base_counts.columns = ['Base de Origem', 'Quantidade']
                st.bar_chart(base_counts.set_index('Base de Origem'), use_container_width=True)

            with g_col6:
                st.markdown("#### 📍 Status por Prédio / Cidade (Abertos x Fechados)")
                status_city = filtered_df.groupby(['cidade_predio', 'status']).size().unstack(fill_value=0)
                st.bar_chart(status_city.head(10), use_container_width=True)
        else:
            st.info("Sem chamados no filtro selecionado para renderizar gráficos.")

    else:
        @st.dialog("Detalhes do Chamado", width="large")
        def show_ticket_details(row):
            title = str(row.get('titulo', '')).strip()
            if title and title.lower() not in ["none", "nan", "null", ""]:
                header_text = f"🎫 Chamado #{row['id']} – {title}"
            else:
                header_text = f"🎫 Chamado #{row['id']}"
            
            with st.expander(header_text, expanded=True):
                col1, col2 = st.columns(2)
                
                with col1:
                    st.markdown("### 👤 Informações do Usuário")
                    user_display = str(row['usuario'])
                    client_id = str(row.get('id_cliente', '')).strip()
                    if client_id and client_id.lower() not in ["none", "nan", "null", ""]:
                        user_display += f" ({client_id})"
                        
                    st.markdown(f"**Usuário:** {user_display}")
                    st.markdown(f"**Localidade:** {row['localidade_fisica']}")
                    st.markdown(f"**Base de Origem:** `{row['base']}`")
                    st.markdown(f"**IP de Origem:** `{row.get('ip_origem') or 'N/A'}`")
                    st.markdown(f"**Hostname:** `{row.get('hostname') or 'N/A'}`")
                    
                    with st.expander("✏️ Editar Título do Chamado", expanded=False):
                        curr_title = str(row.get('titulo', '')).strip()
                        if curr_title.lower() in ["none", "nan", "null", "sem título"]:
                            curr_title = ""
                        new_titulo = st.text_input("Título do Chamado:", value=curr_title, key=f"edit_titulo_{row['id']}")
                        if st.button("💾 Salvar Título", key=f"save_title_btn_{row['id']}"):
                            update_ticket_title(row['id'], new_titulo)
                            st.success("Título do chamado atualizado com sucesso! (Fechar para atualizar a tabela)")
                            st.cache_data.clear()

                    with st.expander("📍 Editar Localização Manual", expanded=False):
                        new_cidade = st.text_input("Cidade - Prédio", value=str(row.get('cidade_predio', '')), key=f"edit_cidade_{row['id']}")
                        new_unidade = st.text_input("Unidade", value=str(row.get('unidade', '')), key=f"edit_unidade_{row['id']}")
                        new_localidade = st.text_input("Localidade Física", value=str(row.get('localidade_fisica', '')), key=f"edit_localidade_{row['id']}")
                        if st.button("💾 Salvar Localização", key=f"save_loc_btn_{row['id']}"):
                            update_ticket_location_details(row['id'], new_localidade, new_cidade, new_unidade)
                            st.success("Localização salva! (Fechar para atualizar a tabela)")
                            st.cache_data.clear()
                    
                with col2:
                    st.markdown("### ⚙️ Classificação & Status")
                    tag_name = str(row['tag']).upper().strip()
                    bg_color = TAG_COLORS.get(tag_name, "#262730")
                    
                    hex_color = bg_color.lstrip('#')
                    try:
                        r, g, b = int(hex_color[0:2], 16), int(hex_color[2:4], 16), int(hex_color[4:6], 16)
                        luminance = (0.299 * r + 0.587 * g + 0.114 * b) / 255
                        text_color = "#ffffff" if luminance < 0.6 else "#212529"
                    except:
                        text_color = "#ffffff"
                        
                    tag_html = f'<span style="background-color: {bg_color}; color: {text_color}; padding: 3px 8px; border-radius: 4px; font-weight: bold; font-family: inherit; font-size: 13px;">{row["tag"]}</span>'
                    st.markdown(f"**TAG Atual:** {tag_html}", unsafe_allow_html=True)
                    
                    tag_options = sorted(list(TAG_COLORS.keys()))
                    try:
                        default_idx = tag_options.index(tag_name)
                    except ValueError:
                        default_idx = 0
                        
                    new_tag = st.selectbox("🏷️ Alterar TAG Manualmente", options=tag_options, index=default_idx, key=f"select_tag_{row['id']}")
                    if new_tag != tag_name:
                        if st.button("💾 Salvar Nova TAG", key=f"save_tag_btn_{row['id']}"):
                            update_ticket_tag(row['id'], new_tag)
                            st.success(f"TAG alterada com sucesso para {new_tag}! (Atualizará na tabela ao fechar o modal)")
                            st.cache_data.clear()
                    
                    link_url = row.get('link')
                    if link_url:
                        st.markdown("---")
                        st.link_button("🔗 Abrir Chamado Original", link_url, width="stretch")
            
            with st.expander("📝 Andamento / Nota de Atendimento", expanded=True):
                current_andamento = str(row.get('andamento', '')).strip()
                if current_andamento.lower() in ["none", "nan", "null", ""]:
                    current_andamento = ""
                new_andamento = st.text_area("Nota rápida sobre o andamento do chamado:", value=current_andamento, key="andamento_modal_ta")
                if st.button("💾 Salvar Nota de Andamento", key="save_andamento_modal_btn"):
                    update_ticket_andamento(row['id'], new_andamento)
                    st.success("Nota de andamento atualizada com sucesso! (Atualizará na tabela ao fechar o modal)")
                    st.cache_data.clear()
            
            with st.expander(f"📝 #1 - {row['Data Formatada']} (Descrição)", expanded=True):
                st.text(row['descricao'])
                
            comments = get_comments_by_ticket(row['id'])
            if comments:
                st.markdown("### 💬 Histórico de Notas e Acompanhamentos")
                for i, c in enumerate(comments, start=2):
                    header = f"🕒 #{i} – {c['data']} – por {c['autor']}"
                    with st.expander(header):
                        st.text(c['texto'])
                
            if st.button("Fechar", key="close_modal_btn"):
                st.rerun()

        cols_to_show = [
            'id', 'status', 'tag', 'andamento', 'localidade_fisica', 
            'cidade_predio', 'unidade', 'usuario', 'datetime_obj', 'base'
        ]
            
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
                if row['base'] == 'CitSmart':
                    link = f"https://suporte.mpms.mp.br/citsmart/pages/serviceRequestIncident/serviceRequestIncident.load?iframe=true&language=pt-BR#/request?idRequest={cid}"
                else:
                    link = "https://central.mpms.mp.br/otrs/index.pl"
            return f"{link}#id:{cid}"

        df_display['id'] = df_display.apply(format_id_link, axis=1)
        
        cols_for_dataframe = list(cols_to_show)
        if 'ip_origem' in df_display.columns:
            cols_for_dataframe.append('ip_origem')
            
        df_final_display = df_display[cols_for_dataframe]
        
        col_tbl_head, col_tbl_cap = st.columns([3, 2])


        with col_tbl_head:
            st.subheader(f"📋 Lista de Chamados ({len(filtered_df)} registros)")
            st.write("Dica: Clique no **checkbox (caixinha de seleção)** no início de qualquer linha para abrir os Detalhes no Modal.")
        with col_tbl_cap:
            components.html("""
            <script src="https://cdnjs.cloudflare.com/ajax/libs/html2canvas/1.4.1/html2canvas.min.js"></script>
            <div style="text-align: right; padding-top: 5px;">
                <button id="btn-cap-tbl" onclick="captureTable()" style="
                    background: linear-gradient(135deg, #10b981 0%, #059669 100%);
                    color: white;
                    border: none;
                    padding: 8px 14px;
                    font-size: 0.85rem;
                    font-weight: bold;
                    border-radius: 6px;
                    cursor: pointer;
                    box-shadow: 0 2px 6px rgba(0,0,0,0.3);
                    transition: all 0.2s ease;
                ">
                    📸 Baixar Tabela em Alta Resolução (PNG)
                </button>
            </div>
            <script>
            function captureTable() {
                const btn = document.getElementById("btn-cap-tbl");
                btn.innerText = "⏳ Gerando PNG em alta resolução...";
                btn.disabled = true;

                const tableEl = window.parent.document.querySelector('div[data-testid="stDataFrame"]') || 
                                window.parent.document.querySelector('.stDataFrame');

                if (!tableEl) {
                    alert("Não foi possível encontrar a tabela na tela.");
                    btn.innerText = "📸 Baixar Tabela em Alta Resolução (PNG)";
                    btn.disabled = false;
                    return;
                }

                html2canvas(tableEl, {
                    scale: 2.5,
                    useCORS: true,
                    backgroundColor: "#0e1117",
                    logging: false
                }).then(canvas => {
                    const link = document.createElement("a");
                    const d = new Date();
                    const dateStr = d.getFullYear() + "-" + String(d.getMonth()+1).padStart(2,'0') + "-" + String(d.getDate()).padStart(2,'0');
                    link.download = `tabela_chamados_sti_${dateStr}.png`;
                    link.href = canvas.toDataURL("image/png");
                    link.click();

                    btn.innerText = "✅ Imagem Salva!";
                    setTimeout(() => {
                        btn.innerText = "📸 Baixar Tabela em Alta Resolução (PNG)";
                        btn.disabled = false;
                    }, 3000);
                }).catch(err => {
                    console.error("Erro no html2canvas:", err);
                    alert("Erro ao capturar tabela: " + err);
                    btn.innerText = "📸 Baixar Tabela em Alta Resolução (PNG)";
                    btn.disabled = false;
                });
            }
            </script>
            """, height=65)
        
        if "last_selected" not in st.session_state:
            st.session_state["last_selected"] = None
        
        def style_dataframe(row):
            tag = str(row.get('tag', '')).upper().strip()
            bg_color = TAG_COLORS.get(tag, "")
            if bg_color:
                hex_color = bg_color.lstrip('#')
                r, g, b = int(hex_color[0:2], 16), int(hex_color[2:4], 16), int(hex_color[4:6], 16)
                luminance = (0.299 * r + 0.587 * g + 0.114 * b) / 255
                text_color = "#ffffff" if luminance < 0.6 else "#212529"
                style = f"background-color: {bg_color}; color: {text_color};"
            else:
                style = ""
                
            return [style] * len(row)

        df_page, current_page, total_pages, total_items = paginate_items(
            df_final_display,
            page_key="chamados",
            items_per_page=items_per_page
        )

        table_height = "content" if items_per_page >= 999999 else 600

        selection_event = st.dataframe(
            df_page.style.apply(style_dataframe, axis=1),
            column_order=cols_to_show,
            column_config={
                "id": st.column_config.LinkColumn("Chamado #", display_text=r".*#id:(.*)"),
                "status": st.column_config.TextColumn("Status"),
                "tag": st.column_config.TextColumn("TAG"),
                "andamento": st.column_config.TextColumn("Andamento"),
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
            height=table_height,
            on_select="rerun",
            selection_mode="single-row",
            key="tabela_chamados_datagrid"
        )

        selected_rows = selection_event.selection.rows if hasattr(selection_event, "selection") else []
        
        if selected_rows:
            current_selected = selected_rows[0]
            if st.session_state["last_selected"] != current_selected:
                st.session_state["last_selected"] = current_selected
                row_data = filtered_df.iloc[(current_page - 1) * items_per_page + current_selected]
                show_ticket_details(row_data)
        else:
            st.session_state["last_selected"] = None

        render_pagination_controls(
            page_key="chamados",
            current_page=current_page,
            total_pages=total_pages,
            total_items=total_items,
            items_per_page=items_per_page
        )


        st.markdown("---")
        st.subheader("📲 Compartilhar Fila por WhatsApp")
        with st.expander("💬 Gerar Resumo Formatado (Pronto para copiar e enviar)", expanded=False):
            if filtered_df.empty:
                st.info("Nenhum chamado na fila filtrada.")
            else:
                usar_resumo_ia = st.checkbox(
                    "✨ Usar Resumos Inteligentes (NLP Local)", 
                    value=True, 
                    help="Usa Processamento de Linguagem Natural (spaCy) rodando totalmente local para resumir o chamado em poucas palavras."
                )
                
                def get_ai_diagnostico(tag, desc):
                    tag = str(tag).upper().strip()
                    desc = str(desc).strip()
                    
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
                    
                    comments_list = get_comments_by_ticket(row['id'])
                    comments_text = ""
                    comments_summary_input = ""
                    if comments_list:
                        comments_text = "💬 *Histórico de Acompanhamento:*"
                        comments_summary_input = "\n".join([f"- {c['data']} ({c['autor']}): {c['texto']}" for c in comments_list])
                        for i, c in enumerate(comments_list, start=1):
                            comments_text += f"\n  • #{i} [{c['data']}] – {c['autor']}: {c['texto']}"
                    
                    diagnostico_ia = get_ai_diagnostico(tag, desc)
                    
                    lines.append(f"🎫 *Chamado #{cid}* ({row['base']})")
                    lines.append(f"👤 *Usuário:* {user}")
                    lines.append(f"📍 *Local:* {loc}")
                    lines.append(f"🏷️ *TAG:* {tag}")
                    lines.append(f"{diagnostico_ia}")
                    
                    if usar_resumo_ia:
                        resumo_nlp = summarize_ticket_locally(desc, comments_summary_input)
                        lines.append(f"📝 *Resumo Inteligente:* {resumo_nlp}")
                    else:
                        lines.append(f"📝 *Problema Completo:*")
                        lines.append(f"{desc}")
                        if comments_text:
                            lines.append(comments_text)
                    
                    lines.append(f"🔗 *Link Direto:* {link}")
                    lines.append("--------------------------------------------------")
                
                whats_text = "\n".join(lines)
                st.write("Dica: Use o botão de **copiar** no canto superior direito do bloco de código abaixo:")
                st.code(whats_text, language="text")
