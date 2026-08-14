import os
import re
import io
from pathlib import Path
import pandas as pd
import requests
import streamlit as st
from src.config import ATOS_NORMAS_API_URL, ATOS_NORMAS_DOWNLOAD_URL
from src.components.pagination import (
    render_items_per_page_selector,
    paginate_items,
    render_pagination_controls
)
from src.components.status_banner import render_log_expander
from src.syncs.sync_fiscalizacao import check_fiscalizacao_sync_running, read_fiscalizacao_last_log_lines
from src.database import (
    get_fiscalizacao_indicacoes_df,
    get_fiscalizacao_publicacoes_df,
    get_fiscalizacao_contador_df
)


def _formatar_texto_portaria(texto: str, nomes_destacar: list[str] | None = None) -> str:
    """Formata o texto bruto da portaria retornado pela API do MPMS para exibição legível."""
    texto = (
        texto
        .replace("\u0096", "–")
        .replace("\u0093", "\u201c")
        .replace("\u0094", "\u201d")
        .replace("\u0092", "\u2019")
        .replace("\u0091", "\u2018")
        .replace("\r\n", "\n")
        .replace("\r", "\n")
    )

    texto = re.sub(r'<br\s*/?>', '\n', texto, flags=re.IGNORECASE)
    texto = re.sub(r'<hr\s*/?>', '\n---\n', texto, flags=re.IGNORECASE)

    corte = re.search(
        r'\n\s*(Procuradoria-Geral de Justi[cç]a|Minist[ée]rio P[úu]b[Ll]ico)',
        texto
    )
    if corte:
        texto = texto[:corte.start()]

    texto = re.sub(r'^[,.\s/\$–r#„\d]{5,}$', '', texto, flags=re.MULTILINE)
    texto = re.sub(r'^\s*:\s*$', '', texto, flags=re.MULTILINE)
    texto = re.sub(r'^\s*[A-Za-z]\s*$', '', texto, flags=re.MULTILINE)
    texto = re.sub(r'\n{3,}', '\n\n', texto)

    if nomes_destacar:
        for nome in nomes_destacar:
            nome = nome.strip()
            if nome and len(nome) > 3:
                padrao = re.compile(re.escape(nome), re.IGNORECASE)
                texto = padrao.sub(f"**🟢 {nome}**", texto)

    return texto.strip()


@st.dialog("📄 Detalhes da Portaria — MPMS", width="large")
def _consultar_portaria_mpms(nome_portaria: str, fiscal_titular: str = "", fiscal_suplente: str = ""):
    """Consulta a API pública do MPMS e exibe os detalhes da portaria em um modal."""
    st.markdown(f"**Consultando:** `{nome_portaria}`")

    url = ATOS_NORMAS_API_URL
    params = {
        "atotit": nome_portaria,
        "atotipcod[]": "1",
        "atocod": "",
    }

    headers = {"User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64)"}

    with st.spinner("Consultando API do MPMS..."):
        try:
            response = requests.get(url, params=params, headers=headers, timeout=15)
            response.raise_for_status()
            data = response.json()
        except requests.exceptions.Timeout:
            st.error("⏱️ A consulta ao MPMS excedeu o tempo limite. Tente novamente mais tarde.")
            return
        except requests.exceptions.ConnectionError:
            st.error("🔌 Não foi possível conectar ao servidor do MPMS. Verifique sua conexão.")
            return
        except Exception as e:
            st.error(f"❌ Erro inesperado na consulta: {e}")
            return

    atos = data.get("atos", [])
    total = data.get("total", 0)

    if not atos:
        st.warning("⚠️ Nenhum resultado encontrado no MPMS para essa portaria.")
        return

    st.success(f"✅ **{total}** resultado(s) encontrado(s).")
    st.markdown("---")

    for idx, ato in enumerate(atos):
        if len(atos) > 1:
            st.markdown(f"### Resultado {idx + 1}")

        col_info1, col_info2 = st.columns(2)
        with col_info1:
            st.markdown(f"**Número do Ato:** `{ato.get('atonum', 'N/A')}`")
            st.markdown(f"**Data do Ato:** {ato.get('atodta', 'N/A')}")
            st.markdown(f"**Data de Publicação:** {ato.get('atodtapub', 'N/A')}")
        with col_info2:
            st.markdown(f"**Origem:** {ato.get('atoorigem', 'N/A')}")
            tipo_info = ato.get("atotipcod") or {}
            subtipo_info = ato.get("atosubtipcod") or {}
            st.markdown(f"**Tipo:** {tipo_info.get('atotipnom', 'N/A')}")
            st.markdown(f"**Subtipo:** {subtipo_info.get('atosubtipnom', 'N/A')}")

        sit_info = ato.get("sitcod") or {}
        if sit_info.get("sitnom"):
            st.markdown(f"**Setor:** {sit_info.get('sitnom')} — {sit_info.get('sitdes', '')}")

        descricao = ato.get("atotit", "")
        if descricao:
            nomes = [n for n in [fiscal_titular, fiscal_suplente] if n.strip()]
            descricao_formatada = _formatar_texto_portaria(descricao, nomes_destacar=nomes if nomes else None)
            with st.expander(f"📝 Portaria nº {ato.get('atonum', 'N/A')} (Descrição)", expanded=True):
                st.markdown(descricao_formatada)

        atocod = ato.get("atocod")
        anx = ato.get("anxcod") or {}
        nome_arquivo = anx.get("anxlin", "")

        if atocod and nome_arquivo:
            st.markdown("---")
            download_url = f"{ATOS_NORMAS_DOWNLOAD_URL}{atocod}"

            col_anx_info, col_anx_btn = st.columns([3, 1])
            with col_anx_info:
                st.info(f"📎 **Arquivo anexo:** {nome_arquivo}")
            with col_anx_btn:
                try:
                    headers_dl = {"User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64)"}
                    resp_dl = requests.get(download_url, headers=headers_dl, timeout=15)
                    resp_dl.raise_for_status()
                    st.download_button(
                        label="⬇️ Baixar PDF",
                        data=resp_dl.content,
                        file_name=nome_arquivo,
                        mime="application/pdf",
                        width='stretch',
                        key=f"dl_portaria_{atocod}_{idx}",
                    )
                except Exception:
                    st.link_button("🔗 Abrir no MPMS", download_url, width='stretch')
        elif nome_arquivo:
            st.markdown("---")
            st.info(f"📎 **Arquivo anexo:** {nome_arquivo}")

        if len(atos) > 1 and idx < len(atos) - 1:
            st.markdown("---")


def render_contracts_page():
    """Renderiza a página de Fiscalização de Contratos a partir da planilha oficial do OneDrive/SharePoint."""
    st.title("📜 Fiscalização de Contratos & Processos SAJ")
    st.write("Acompanhamento das indicações de fiscais titulares, suplentes, processos SAJ e portarias publicadas.")
    st.markdown("---")

    fiscalizacao_ativo = check_fiscalizacao_sync_running()

    if "was_fiscalizacao_syncing" not in st.session_state:
        st.session_state["was_fiscalizacao_syncing"] = False

    if st.session_state["was_fiscalizacao_syncing"] and not fiscalizacao_ativo:
        st.session_state["was_fiscalizacao_syncing"] = False
        st.toast("🎉 Sincronização de fiscais concluída com sucesso!", icon="✅")
        st.rerun()

    if fiscalizacao_ativo:
        st.session_state["was_fiscalizacao_syncing"] = True

    render_log_expander(
        "🤖 Sincronização de Fiscais em Segundo Plano",
        fiscalizacao_ativo,
        read_fiscalizacao_last_log_lines,
        check_fiscalizacao_sync_running,
        "O robô está realizando a leitura segura da planilha no OneDrive. O painel permanece livre para uso!"
    )

    df_indicacoes = get_fiscalizacao_indicacoes_df()
    df_publicacoes = get_fiscalizacao_publicacoes_df()
    df_contador = get_fiscalizacao_contador_df()

    if not df_indicacoes.empty:
        df_indicacoes.columns = [str(col).strip() for col in df_indicacoes.columns]
    
    if not df_publicacoes.empty:
        df_publicacoes.columns = [str(col).strip() for col in df_publicacoes.columns]

    fiscais_foco = [
        "Paulo Henrique Gonçalves Rezende",
        "Reginaldo da Silva Bandeira",
        "Luiz Leonardo Villalba"
    ]

    st.subheader("📊 Resumo de Contratos por Fiscal")
    kpi_cols = st.columns(3)

    for i, fiscal in enumerate(fiscais_foco):
        count_titular = 0
        count_suplente = 0
        
        if not df_indicacoes.empty:
            if "Fiscal titular" in df_indicacoes.columns:
                count_titular = (df_indicacoes["Fiscal titular"].astype(str).str.strip() == fiscal).sum()
            if "Fiscal suplente" in df_indicacoes.columns:
                count_suplente = (df_indicacoes["Fiscal suplente"].astype(str).str.strip() == fiscal).sum()
        
        total_fiscal = count_titular + count_suplente
        primeiro_nome = fiscal.split()[0] + " " + fiscal.split()[-1]

        with kpi_cols[i % 3]:
            st.markdown(f"""
            <div class="metric-card" style="border-left-color: #ff4b4b; text-align: center;">
                <div class="metric-title" style="color:#ff4b4b;">👤 {primeiro_nome}</div>
                <div class="metric-value">{total_fiscal} <span style="font-size:0.9rem; opacity: 0.7;">processos</span></div>
                <div style="display:flex; justify-content:space-around; margin-top:8px; font-size:0.8rem;">
                    <span>📌 Titular: <b>{count_titular}</b></span>
                    <span>🔄 Suplente: <b>{count_suplente}</b></span>
                </div>
            </div>
            """, unsafe_allow_html=True)


    st.markdown("<br>", unsafe_allow_html=True)

    st.sidebar.markdown("## 🔍 Filtros de Contratos")
    
    opcoes_fiscais = ["Todos"] + fiscais_foco
    selected_fiscal_filter = st.sidebar.selectbox("👤 Filtrar por Fiscal:", opcoes_fiscais)
    
    search_text = st.sidebar.text_input("🔍 Buscar por Nº SAJ, Objeto ou Contrato:", "")

    st.sidebar.markdown("---")
    st.sidebar.markdown("## ⚙️ Ações e Sincronização")
    if fiscalizacao_ativo:
        st.sidebar.button("🤖 Sincronizando...", width='stretch', disabled=True)
    else:
        if st.sidebar.button("🔄 Sincronizar Planilha", type="primary", width='stretch', help="Busca atualizações na planilha do SharePoint em segundo plano."):
            import sys, subprocess, time
            subprocess.Popen([sys.executable, "src/syncs/sync_fiscalizacao.py"])
            time.sleep(0.5)
            st.toast("🚀 Sincronização iniciada em segundo plano!", icon="🤖")
            st.rerun()

    st.sidebar.markdown("---")

    items_per_page = render_items_per_page_selector(
        key_prefix="fiscalizacao",
        options=[10, 25, 50, 100, "Todos"],
        default_index=1,
        label="📄 Contratos por página:"
    )


    df_filtered_ind = df_indicacoes.copy()
    
    if selected_fiscal_filter != "Todos":
        cond_titular = df_filtered_ind["Fiscal titular"].astype(str).str.strip() == selected_fiscal_filter if "Fiscal titular" in df_filtered_ind.columns else False
        cond_suplente = df_filtered_ind["Fiscal suplente"].astype(str).str.strip() == selected_fiscal_filter if "Fiscal suplente" in df_filtered_ind.columns else False
        df_filtered_ind = df_filtered_ind[cond_titular | cond_suplente]

    if search_text:
        mask = pd.Series(False, index=df_filtered_ind.index)
        for col in df_filtered_ind.columns:
            mask = mask | df_filtered_ind[col].astype(str).str.contains(search_text, case=False, na=False)
        df_filtered_ind = df_filtered_ind[mask]

    # ABAS INTERNAS DE VISUALIZAÇÃO COM QUERY PARAMETERS (?subtab=slug)
    FISCAL_SUBTAB_MAP = {
        "indicacoes": "📋 Indicações de Fiscais",
        "graficos": "📈 Gráficos & Estatísticas",
        "publicacoes": "📰 Publicações & Portarias",
        "contadora": "📊 Tabela Contadora"
    }
    FISCAL_SUBTAB_REVERSE = {v: k for k, v in FISCAL_SUBTAB_MAP.items()}

    url_subtab = st.query_params.get("subtab", "indicacoes")
    default_title = FISCAL_SUBTAB_MAP.get(url_subtab, "📋 Indicações de Fiscais")
    options = list(FISCAL_SUBTAB_MAP.values())
    default_idx = options.index(default_title) if default_title in options else 0

    selected_subtab = st.radio(
        "Navegação:",
        options=options,
        index=default_idx,
        horizontal=True,
        label_visibility="collapsed",
        key="fiscalizacao_subtab_radio"
    )

    new_slug = FISCAL_SUBTAB_REVERSE.get(selected_subtab, "indicacoes")
    if st.query_params.get("subtab") != new_slug:
        st.query_params["subtab"] = new_slug

    st.markdown("<br>", unsafe_allow_html=True)

    if selected_subtab == "📋 Indicações de Fiscais":

        c_head1, c_head2 = st.columns([3, 1])
        with c_head1:
            st.subheader(f"📋 Processos e Indicações ({len(df_filtered_ind)} registros)")
        with c_head2:
            if not df_filtered_ind.empty:
                output_buffer = io.BytesIO()
                with pd.ExcelWriter(output_buffer, engine='openpyxl') as writer:
                    df_filtered_ind.to_excel(writer, index=False, sheet_name='Fiscais')
                excel_bytes = output_buffer.getvalue()
                
                st.download_button(
                    label="📥 Exportar Excel",
                    data=excel_bytes,
                    file_name="contratos_fiscais_filtrados.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    width='stretch'
                )

        if not df_filtered_ind.empty:
            df_page_ind, current_page, total_pages, total_items = paginate_items(
                df_filtered_ind,
                page_key="fiscal_ind",
                items_per_page=items_per_page
            )

            st.dataframe(
                df_page_ind,
                width='stretch',
                hide_index=True,
                column_config={
                    "nº Saj": st.column_config.TextColumn("Nº SAJ"),
                    "Fiscal titular": st.column_config.TextColumn("Fiscal Titular"),
                    "Fiscal suplente": st.column_config.TextColumn("Fiscal Suplente"),
                    "Objeto": st.column_config.TextColumn("Objeto / Descrição"),
                    "Contrato": st.column_config.TextColumn("Contrato / Empenho"),
                }
            )

            render_pagination_controls(
                page_key="fiscal_ind",
                current_page=current_page,
                total_pages=total_pages,
                total_items=total_items,
                items_per_page=items_per_page
            )

        else:
            st.info("Nenhum registro encontrado para os filtros selecionados.")

    elif selected_subtab == "📈 Gráficos & Estatísticas":
        st.subheader("📈 Visão Geral da Carga de Trabalho dos Fiscais")
        if not df_indicacoes.empty and "Fiscal titular" in df_indicacoes.columns and "Fiscal suplente" in df_indicacoes.columns:
            df_t = df_indicacoes["Fiscal titular"].dropna().astype(str).str.strip().value_counts().reset_index()
            df_t.columns = ["Fiscal", "Como Titular"]
            
            df_s = df_indicacoes["Fiscal suplente"].dropna().astype(str).str.strip().value_counts().reset_index()
            df_s.columns = ["Fiscal", "Como Suplente"]
            
            df_comp = pd.merge(df_t, df_s, on="Fiscal", how="outer").fillna(0)
            df_comp["Como Titular"] = df_comp["Como Titular"].astype(int)
            df_comp["Como Suplente"] = df_comp["Como Suplente"].astype(int)
            df_comp["Total Processos"] = df_comp["Como Titular"] + df_comp["Como Suplente"]
            df_comp = df_comp.sort_values(by="Total Processos", ascending=False)
            
            g_col1, g_col2 = st.columns(2)
            with g_col1:
                st.markdown("#### 📌 Distribuição de Titularidades")
                st.bar_chart(df_comp.set_index("Fiscal")[["Como Titular"]], width='stretch')
            with g_col2:
                st.markdown("#### 🔄 Distribuição de Suplências")
                st.bar_chart(df_comp.set_index("Fiscal")[["Como Suplente"]], width='stretch')
                
            st.markdown("---")
            st.markdown("#### 📊 Carga Total Comparativa de Fiscais")
            st.bar_chart(df_comp.set_index("Fiscal")[["Como Titular", "Como Suplente"]], width='stretch')

            st.markdown("---")
            st.subheader("📦 Agrupamento por Tipo de Objeto / Equipamento")
            
            if "Objeto" in df_indicacoes.columns:
                df_obj = df_indicacoes["Objeto"].dropna().astype(str).str.strip().str.lower()
                
                def categorizar_objeto(desc):
                    if "monitor" in desc:
                        return "🖥️ Monitores"
                    elif "desktop" in desc or "computador" in desc:
                        return "💻 Desktops / Computadores"
                    elif "fone" in desc or "headset" in desc:
                        return "🎧 Fones / Headsets"
                    elif "webcam" in desc or "mouse" in desc:
                        return "🖱️ Periféricos (Webcam/Mouse/Teclado)"
                    elif "notebook" in desc or "laptop" in desc:
                        return "💻 Notebooks"
                    elif "telefone" in desc or "ramal" in desc:
                        return "📞 Telefonia / Ramais"
                    elif "hd" in desc or "ssd" in desc:
                        return "💾 Armazenamento (HD/SSD)"
                    elif "internet" in desc or "satélite" in desc:
                        return "📡 Internet / Conectividade"
                    elif "scanner" in desc:
                        return "🖨️ Scanners / Impressão"
                    elif "tablet" in desc:
                        return "📱 Tablets"
                    else:
                        return "📦 Outros Suprimentos / Serviços"

                df_indicacoes_cats = df_indicacoes.copy()
                df_indicacoes_cats["Categoria_Objeto"] = df_indicacoes_cats["Objeto"].astype(str).apply(categorizar_objeto)
                
                counts_obj = df_indicacoes_cats["Categoria_Objeto"].value_counts().reset_index()
                counts_obj.columns = ["Categoria", "Quantidade"]
                
                o_col1, o_col2 = st.columns([2, 1])
                with o_col1:
                    st.bar_chart(counts_obj.set_index("Categoria"), width='stretch')
                with o_col2:
                    st.markdown("##### 📌 Quantidade por Tipo:")
                    for _, r in counts_obj.iterrows():
                        st.markdown(f"- **{r['Categoria']}**: `{r['Quantidade']}` processos")
        else:
            st.info("Dados insuficientes para renderização dos gráficos.")

    elif selected_subtab == "📰 Publicações & Portarias":
        st.subheader("📰 Publicações em Diário Oficial & Portarias")
        df_filtered_pub = df_publicacoes.copy()
        
        if selected_fiscal_filter != "Todos" and not df_filtered_pub.empty:
            cond_t = df_filtered_pub["Fiscal titular"].astype(str).str.strip() == selected_fiscal_filter if "Fiscal titular" in df_filtered_pub.columns else False
            cond_s = df_filtered_pub["Fiscal suplente"].astype(str).str.strip() == selected_fiscal_filter if "Fiscal suplente" in df_filtered_pub.columns else False
            df_filtered_pub = df_filtered_pub[cond_t | cond_s]

        if search_text and not df_filtered_pub.empty:
            mask_p = pd.Series(False, index=df_filtered_pub.index)
            for col in df_filtered_pub.columns:
                mask_p = mask_p | df_filtered_pub[col].astype(str).str.contains(search_text, case=False, na=False)
            df_filtered_pub = df_filtered_pub[mask_p]

        if not df_filtered_pub.empty:
            if "pub_last_selected" not in st.session_state:
                st.session_state["pub_last_selected"] = None

            df_page_pub, current_page_pub, total_pages_pub, total_items_pub = paginate_items(
                df_filtered_pub,
                page_key="fiscal_pub",
                items_per_page=items_per_page
            )

            selection_pub = st.dataframe(
                df_page_pub,
                width='stretch',
                hide_index=True,
                column_config={
                    "nº Saj": st.column_config.TextColumn("Nº SAJ"),
                    "Fiscal titular": st.column_config.TextColumn("Fiscal Titular"),
                    "Fiscal suplente": st.column_config.TextColumn("Fiscal Suplente"),
                    "Objeto": st.column_config.TextColumn("Objeto / Nota de Empenho"),
                    "Portaria": st.column_config.TextColumn("Portaria"),
                    "Data Portaria": st.column_config.DateColumn("Data Portaria", format="DD/MM/YYYY"),
                },
                on_select="rerun",
                selection_mode="single-row",
                key="tabela_publicacoes_datagrid",
            )

            selected_pub_rows = selection_pub.selection.rows if hasattr(selection_pub, "selection") else []

            if selected_pub_rows:
                current_pub_selected = selected_pub_rows[0]
                if st.session_state["pub_last_selected"] != current_pub_selected:
                    st.session_state["pub_last_selected"] = current_pub_selected
                    row_pub = df_filtered_pub.iloc[(current_page_pub - 1) * items_per_page + current_pub_selected]
                    portaria_valor = str(row_pub.get("Portaria", "")).strip() if "Portaria" in df_filtered_pub.columns else ""
                    fiscal_t = str(row_pub.get("Fiscal titular", "")).strip() if "Fiscal titular" in df_filtered_pub.columns else ""
                    fiscal_s = str(row_pub.get("Fiscal suplente", "")).strip() if "Fiscal suplente" in df_filtered_pub.columns else ""
                    if portaria_valor:
                        _consultar_portaria_mpms(portaria_valor, fiscal_titular=fiscal_t, fiscal_suplente=fiscal_s)
                    else:
                        st.toast("⚠️ Essa linha não possui uma portaria para consultar.", icon="⚠️")
            else:
                st.session_state["pub_last_selected"] = None

            render_pagination_controls(
                page_key="fiscal_pub",
                current_page=current_page_pub,
                total_pages=total_pages_pub,
                total_items=total_items_pub,
                items_per_page=items_per_page
            )

        else:
            st.info("Nenhuma publicação/portaria encontrada.")

    elif selected_subtab == "📊 Tabela Contadora":
        st.subheader("📊 Tabela de Contagem Geral")
        if not df_contador.empty:
            st.dataframe(df_contador, width='stretch', hide_index=True)
        else:
            st.info("Aba Contador indisponível ou vazia na planilha.")

