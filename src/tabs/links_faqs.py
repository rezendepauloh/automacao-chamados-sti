import json
import sqlite3
from pathlib import Path
import pandas as pd
import streamlit as st
from src.config import setup_logging, DEBUG_DIR_FAQ

logger = setup_logging(DEBUG_DIR_FAQ / "faq.log", "faq")


def render_faq_page():
    """Renderiza a página de FAQs, Tutoriais do SharePoint e Links Úteis da Bancada."""
    st.title("📚 FAQ, Tutoriais & Links Úteis da Bancada")
    st.write("Base de conhecimento centralizada com tutoriais da equipe e atalhos rápidos para sistemas externos.")
    st.markdown("---")

    root_dir = Path(__file__).parent.parent.parent
    db_path = root_dir / "chamados.db"
    json_faq_path = root_dir / "temp" / "faqs_template.json"
    json_links_path = root_dir / "temp" / "links_uteis_template.json"
    
    try:
        conn = sqlite3.connect(db_path)
        cursor = conn.cursor()
        cursor.execute("""
            CREATE TABLE IF NOT EXISTS faqs (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                titulo TEXT NOT NULL,
                tipo_faq TEXT NOT NULL,
                url TEXT NOT NULL UNIQUE,
                conteudo TEXT,
                data_atualizacao DATETIME DEFAULT CURRENT_TIMESTAMP
            )
        """)
        
        cursor.execute("PRAGMA table_info(faqs)")
        cols_db = [col[1] for col in cursor.fetchall()]
        if "conteudo" not in cols_db:
            cursor.execute("ALTER TABLE faqs ADD COLUMN conteudo TEXT")
            conn.commit()

        cursor.execute("SELECT COUNT(*) FROM faqs")
        count = cursor.fetchone()[0]

        if count == 0 and json_faq_path.exists():
            with open(json_faq_path, "r", encoding="utf-8") as f:
                faqs_json = json.load(f)
            
            for item in faqs_json:
                cursor.execute("""
                    INSERT OR IGNORE INTO faqs (titulo, tipo_faq, url, conteudo)
                    VALUES (?, ?, ?, ?)
                """, (item.get("titulo"), item.get("tipo_faq", "Geral"), item.get("url"), item.get("conteudo")))
            conn.commit()

        df_faqs = pd.read_sql_query("SELECT id, titulo, tipo_faq, url, conteudo FROM faqs", conn)
        conn.close()
    except Exception as e:
        st.error(f"Erro ao carregar banco de dados de FAQs: {e}")
        df_faqs = pd.DataFrame()

    links_uteis = []
    if json_links_path.exists():
        try:
            with open(json_links_path, "r", encoding="utf-8") as f:
                links_uteis = json.load(f)
        except Exception as e:
            logger.error(f"Erro ao ler links_uteis_template.json: {e}")

    tab_faqs, tab_links = st.tabs(["📚 FAQs & Tutoriais (SharePoint)", "🔗 Links Úteis da Bancada"])

    with tab_faqs:
        if df_faqs.empty:
            st.info("Nenhum FAQ cadastrado no momento.")
        else:
            @st.dialog("📖 Leitor de FAQ", width="large")
            def open_faq_modal(faq_id):
                faq_item = df_faqs[df_faqs['id'] == faq_id].iloc[0]
                
                st.markdown("""
                <style>
                .faq-container {
                    font-family: 'Segoe UI', system-ui, -apple-system, sans-serif;
                    color: #e0e0e0;
                    line-height: 1.6;
                }
                .faq-container h1, .faq-container h2, .faq-container h3, .faq-container h4 {
                    color: #ffffff !important;
                    margin-top: 1.5rem;
                    margin-bottom: 0.75rem;
                    font-weight: 600;
                }
                .faq-container p {
                    margin-bottom: 1rem;
                    font-size: 0.95rem;
                }
                .faq-container img {
                    max-width: 100% !important;
                    height: auto !important;
                    border-radius: 8px !important;
                    margin: 16px 0 !important;
                    box-shadow: 0 4px 12px rgba(0,0,0,0.4) !important;
                    border: 1px solid #343541 !important;
                }
                .faq-container ol, .faq-container ul {
                    padding-left: 1.5rem;
                    margin-bottom: 1rem;
                }
                .faq-container li {
                    margin-bottom: 0.4rem;
                }
                .faq-container code {
                    background-color: #2a2b36;
                    color: #ff4b4b;
                    padding: 2px 6px;
                    border-radius: 4px;
                    font-size: 0.9rem;
                }
                </style>
                """, unsafe_allow_html=True)

                st.subheader(faq_item['titulo'])
                st.caption(f"Categoria: **{faq_item['tipo_faq']}**")
                st.markdown("---")
                
                if faq_item['conteudo'] and str(faq_item['conteudo']).strip():
                    st.markdown(f'<div class="faq-container">{faq_item["conteudo"]}</div>', unsafe_allow_html=True)
                else:
                    st.info("O conteúdo detalhado deste FAQ ainda não foi sincronizado localmente.")
                    st.write("Você pode visualizar o tutorial completo diretamente no SharePoint pelo botão abaixo.")
                    
                st.markdown("---")
                st.markdown(f'<a href="{faq_item["url"]}" target="_blank" style="display: inline-block; background-color: #ff4b4b; color: white; text-decoration: none; font-weight: bold; padding: 8px 16px; border-radius: 6px;">🔗 Abrir no SharePoint (Nova Aba) ↗</a>', unsafe_allow_html=True)

            st.sidebar.markdown("## 🔍 Filtros do FAQ")
            search_query = st.sidebar.text_input("Buscar por palavra-chave:", "", key="faq_search")
            
            tipos_disponiveis = ["Todos"] + sorted(df_faqs['tipo_faq'].dropna().unique().tolist())
            selected_tipo = st.sidebar.selectbox("📂 Categoria:", tipos_disponiveis, key="faq_cat")

            filtered_df = df_faqs.copy()
            if search_query:
                filtered_df = filtered_df[filtered_df['titulo'].str.contains(search_query, case=False, na=False)]
            if selected_tipo != "Todos":
                filtered_df = filtered_df[filtered_df['tipo_faq'] == selected_tipo]

            st.markdown(f"**Exibindo {len(filtered_df)} de {len(df_faqs)} FAQs / Tutoriais**")
            st.markdown("<br>", unsafe_allow_html=True)

            cols = st.columns(2)
            for index, row in filtered_df.iterrows():
                col_target = cols[index % 2]
                with col_target:
                    with st.container(border=True):
                        st.caption(f"📌 {row['tipo_faq']}")
                        st.subheader(row['titulo'])
                        
                        c_btn1, c_btn2 = st.columns([1, 1])
                        with c_btn1:
                            if st.button("📖 Ler Tutorial", key=f"btn_read_{row['id']}", use_container_width=True):
                                open_faq_modal(row['id'])
                        with c_btn2:
                            st.markdown(f'<a href="{row["url"]}" target="_blank" style="display: block; text-align: center; background-color: #2a2b36; border: 1px solid #343541; color: white; text-decoration: none; font-size: 0.85rem; padding: 6px; border-radius: 6px; font-weight: bold;">🔗 SharePoint ↗</a>', unsafe_allow_html=True)

    with tab_links:
        st.subheader("🌐 Links e Atalhos Rápidos da Bancada")
        st.write("Acesso direto aos sistemas operacionais, filas de atendimento e ferramentas externas.")
        st.markdown("<br>", unsafe_allow_html=True)

        if not links_uteis:
            st.info("Nenhum link útil cadastrado em `temp/links_uteis_template.json`.")
        else:
            search_link = st.text_input("🔍 Pesquisar por Nome do Sistema / Link:", "", key="search_link_input")
            
            filtered_links = links_uteis
            if search_link:
                filtered_links = [l for l in links_uteis if search_link.lower() in l.get("titulo", "").lower()]

            link_cols = st.columns(3)
            for idx, item in enumerate(filtered_links):
                col_lk = link_cols[idx % 3]
                with col_lk:
                    with st.container(border=True):
                        st.markdown(f"#### 🚀 {item.get('titulo', 'Sem Título')}")
                        st.caption(item.get("url", ""))
                        st.markdown("<br>", unsafe_allow_html=True)
                        st.markdown(
                            f'<a href="{item.get("url")}" target="_blank" style="display: block; text-align: center; background-color: #ff4b4b; color: white; text-decoration: none; font-weight: bold; padding: 10px; border-radius: 6px;">🔗 Acessar Sistema ↗</a>', 
                            unsafe_allow_html=True
                        )
