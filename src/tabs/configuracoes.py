import streamlit as st
import os
from pathlib import Path
from src.database.settings_db import (
    get_all_settings,
    set_setting,
    seed_settings_from_env_if_empty
)
from src.database.connection import DB_TYPE, DB_PATH
from src.components.subtabs import render_subtabs

CONFIG_SUBTAB_MAP = {
    "rede": "🔑 Rede & AD (OTRS / CitSmart)",
    "sccm": "💻 Administrador SCCM",
    "papercut": "🖨️ PaperCut (Impressoras)",
    "oxe": "📞 Central Telefônica (OXE)",
    "urls": "🌐 Portais & Links Web",
    "sharepoint": "📂 SharePoint & Planilhas",
    "ia": "🤖 Inteligência Artificial (Gemini)",
    "whatsapp": "📱 WhatsApp & Alertas (Evolution)",
    "schedules": "⏰ Agendamentos & Cron Jobs"
}

def render_configuracoes_page():
    """Renderiza a página visual de configurações do sistema e cofre de credenciais."""
    
    # Garante que as configurações estejam inicializadas a partir do .env caso tabela esteja vazia
    seed_settings_from_env_if_empty(force=False)
    
    # Cabeçalho da página
    st.markdown(r"""
        <div style="background: var(--metric-bg, #1e293b); padding: 18px 24px; border-radius: 12px; border-left: 6px solid #3b82f6; border-top: 1px solid var(--metric-border, #2d3139); border-right: 1px solid var(--metric-border, #2d3139); border-bottom: 1px solid var(--metric-border, #2d3139); margin-bottom: 20px; box-shadow: 0 2px 8px rgba(0,0,0,0.08);">
            <div style="display: flex; justify-content: space-between; align-items: center;">
                <h2 style="color: var(--metric-value-color, #ffffff); margin: 0; font-size: 24px; font-weight: 700;">⚙️ Configurações & Credenciais do Sistema</h2>
                <span style="background-color: rgba(59, 130, 246, 0.15); color: #38bdf8; font-size: 13px; font-weight: 600; padding: 4px 12px; border-radius: 20px; border: 1px solid #0284c7;">
                    🔒 Cofre Criptografado (AES-Fernet)
                </span>
            </div>
            <p style="color: var(--metric-title-color, #94a3b8); margin: 6px 0 0 0; font-size: 14px;">
                Gerencie todos os parâmetros de conexão, contas de serviço, links e credenciais de acesso de forma centralizada, segura e persistente.
            </p>
        </div>
    """, unsafe_allow_html=True)

    # Status do Banco de Dados
    col_db_info, col_btn_sync = st.columns([3, 1])
    with col_db_info:
        db_desc = "🐘 PostgreSQL (Red Hat / Produção)" if DB_TYPE in ["postgres", "postgresql"] else f"💾 SQLite Local (`{DB_PATH}`)"
        st.caption(f"**Armazenamento Ativo:** {db_desc} | **Criptografia:** Ativa com proteção de dados em repouso")
    with col_btn_sync:
        if st.button("🔄 Reimportar do .env", help="Recarrega as variáveis e senhas existentes no arquivo .env e Keyring para o banco de dados"):
            seed_settings_from_env_if_empty(force=True)
            st.toast("✅ Configurações reimportadas do .env com sucesso!", icon="🔄")
            st.rerun()

    # Carrega as configurações atuais do banco (descriptografadas para o formulário)
    settings = get_all_settings(decrypt=True)
    
    # Subtabs sincronizadas com URL Search Params (?subtab=slug)
    selected_subtab = render_subtabs(CONFIG_SUBTAB_MAP, default_slug="rede", key="config_subtabs_radio")
    st.markdown("<br>", unsafe_allow_html=True)

    form_values = {}

    # -------------------------------------------------------------------------
    # TAB 1: Rede & AD
    # -------------------------------------------------------------------------
    if selected_subtab == "🔑 Rede & AD (OTRS / CitSmart)":
        st.markdown("#### 🔑 Credenciais do Active Directory & Rede Corporativa")
        st.info("Estas credenciais são utilizadas para a raspagem dos chamados do **OTRS** e **CitSmart**, bem como nas rotinas de automação.")
        
        c1, c2 = st.columns(2)
        with c1:
            form_values["AD_USER"] = (
                st.text_input("Usuário da Rede / AD", value=settings.get("AD_USER", {}).get("value", "paulogoncalves"), key="cfg_ad_user"),
                False, "rede", "Usuário de Rede / Active Directory"
            )
        with c2:
            form_values["AD_PASSWORD"] = (
                st.text_input("Senha da Rede / AD", value=settings.get("AD_PASSWORD", {}).get("value", ""), type="password", key="cfg_ad_pass", help="Armazenada de forma criptografada no banco"),
                True, "rede", "Senha de Rede / Active Directory (OTRS & CitSmart)"
            )

        c3, c4 = st.columns(2)
        with c3:
            form_values["AD_DOMAIN"] = (
                st.text_input("Domínio FQDN", value=settings.get("AD_DOMAIN", {}).get("value", "in.mpe.ms.gov.br"), key="cfg_ad_domain"),
                False, "rede", "Domínio FQDN do Active Directory"
            )
            form_values["AD_SHORT"] = (
                st.text_input("Nome NetBIOS Curto", value=settings.get("AD_SHORT", {}).get("value", "MPE"), key="cfg_ad_short"),
                False, "rede", "Nome NetBIOS curto do Domínio"
            )
        with c4:
            form_values["AD_MMC"] = (
                st.text_input("Base DN / MMC", value=settings.get("AD_MMC", {}).get("value", "DC=in,DC=mpe,DC=ms,DC=gov,DC=br"), key="cfg_ad_mmc"),
                False, "rede", "Base DN / MMC do Active Directory"
            )
            form_values["AD_EMAIL"] = (
                st.text_input("Sufixo de E-mail Institucional", value=settings.get("AD_EMAIL", {}).get("value", "mpms.mp.br"), key="cfg_ad_email"),
                False, "rede", "Sufixo de e-mail institucional"
            )

    # -------------------------------------------------------------------------
    # TAB 2: SCCM
    # -------------------------------------------------------------------------
    elif selected_subtab == "💻 Administrador SCCM":
        st.markdown("#### 💻 Conta Administrativa do SCCM / MECM")
        st.info("Utilizada para consultar IP e Hostname das máquinas dos usuários e gerar credenciais administrativas seguras (`cred_admin.xml`).")
        
        c1, c2 = st.columns(2)
        with c1:
            form_values["SCCM_ADMIN_USER"] = (
                st.text_input("Usuário Administrador (ex: paulo_admin)", value=settings.get("SCCM_ADMIN_USER", {}).get("value", "paulo_admin"), key="cfg_sccm_user"),
                False, "sccm", "Conta de Administrador para consultas SCCM"
            )
        with c2:
            form_values["SCCM_ADMIN_PASSWORD"] = (
                st.text_input("Senha da Conta Administradora", value=settings.get("SCCM_ADMIN_PASSWORD", {}).get("value", ""), type="password", key="cfg_sccm_pass", help="Armazenada de forma criptografada"),
                True, "sccm", "Senha da conta Administradora do SCCM"
            )

        c3, c4 = st.columns(2)
        with c3:
            form_values["SCCM_SERVER"] = (
                st.text_input("Servidor do SCCM (FQDN)", value=settings.get("SCCM_SERVER", {}).get("value", "srv-1046.in.mpe.ms.gov.br"), key="cfg_sccm_server"),
                False, "sccm", "Servidor FQDN do SCCM"
            )
        with c4:
            form_values["SCCM_SITE_CODE"] = (
                st.text_input("Código do Site", value=settings.get("SCCM_SITE_CODE", {}).get("value", "PGJ"), key="cfg_sccm_site"),
                False, "sccm", "Código do Site do SCCM"
            )

    # -------------------------------------------------------------------------
    # TAB 3: PaperCut
    # -------------------------------------------------------------------------
    elif selected_subtab == "🖨️ PaperCut (Impressoras)":
        st.markdown("#### 🖨️ Servidor de Impressão PaperCut MF")
        st.info("Credenciais e rotas para raspagem e monitoramento das impressoras e suprimentos.")

        c1, c2 = st.columns(2)
        with c1:
            form_values["PAPERCUT_USER"] = (
                st.text_input("Usuário do PaperCut", value=settings.get("PAPERCUT_USER", {}).get("value", "admin"), key="cfg_pc_user"),
                False, "papercut", "Usuário Administrador do PaperCut"
            )
        with c2:
            form_values["PAPERCUT_PASS"] = (
                st.text_input("Senha do PaperCut", value=settings.get("PAPERCUT_PASS", {}).get("value", ""), type="password", key="cfg_pc_pass"),
                True, "papercut", "Senha do Administrador do PaperCut"
            )

        form_values["PAPERCUT_URL"] = (
            st.text_input("URL Principal do PaperCut", value=settings.get("PAPERCUT_URL", {}).get("value", "http://impressora.mpms.mp.br:9191/admin"), key="cfg_pc_url"),
            False, "papercut", "URL do painel administrativo do PaperCut"
        )
        c3, c4 = st.columns(2)
        with c3:
            form_values["PAPERCUT_PRINTER_LIST_URL"] = (
                st.text_input("URL da Lista de Impressoras", value=settings.get("PAPERCUT_PRINTER_LIST_URL", {}).get("value", "http://impressora.mpms.mp.br:9191/app?service=page/PrinterList"), key="cfg_pc_print_url"),
                False, "papercut", "URL da listagem de impressoras"
            )
        with c4:
            form_values["PAPERCUT_DEVICE_LIST_URL"] = (
                st.text_input("URL da Lista de Dispositivos", value=settings.get("PAPERCUT_DEVICE_LIST_URL", {}).get("value", "http://impressora.mpms.mp.br:9191/app?service=page/DeviceList"), key="cfg_pc_dev_url"),
                False, "papercut", "URL da listagem de dispositivos multifuncionais"
            )

    # -------------------------------------------------------------------------
    # TAB 4: Telefonia (OXE)
    # -------------------------------------------------------------------------
    elif selected_subtab == "📞 Central Telefônica (OXE)":
        st.markdown("#### 📞 Central Telefônica Alcatel-Lucent OmniPCX Enterprise (OXE)")
        st.info("Parâmetros de conexão e credenciais de acesso à central telefônica para consulta de ramais.")

        c1, c2 = st.columns(2)
        with c1:
            form_values["OXE_USER"] = (
                st.text_input("Usuário do OXE", value=settings.get("OXE_USER", {}).get("value", "mtcl"), key="cfg_oxe_user"),
                False, "oxe", "Usuário de acesso à Central Telefônica OXE"
            )
        with c2:
            form_values["OXE_PASS"] = (
                st.text_input("Senha do OXE", value=settings.get("OXE_PASS", {}).get("value", ""), type="password", key="cfg_oxe_pass"),
                True, "oxe", "Senha de acesso à Central Telefônica OXE"
            )
        form_values["OXE_URL"] = (
            st.text_input("URL da Central Telefônica", value=settings.get("OXE_URL", {}).get("value", "https://10.12.32.30"), key="cfg_oxe_url"),
            False, "oxe", "URL da Central Telefônica OXE"
        )

    # -------------------------------------------------------------------------
    # TAB 5: Portais & Links Web
    # -------------------------------------------------------------------------
    elif selected_subtab == "🌐 Portais & Links Web":
        st.markdown("#### 🌐 Portais Institucionais e Endpoints de Consulta")
        st.info("Links dos sistemas corporativos para integração e atalhos rápidos.")

        form_values["CITSMART_LINK"] = (
            st.text_input("URL do CitSmart", value=settings.get("CITSMART_LINK", {}).get("value", "https://suporte.mpms.mp.br"), key="cfg_citsmart_link"),
            False, "urls", "URL principal do CitSmart"
        )
        form_values["CITSMART_LINK_NOVO"] = (
            st.text_input("URL Formulário Novo CitSmart", value=settings.get("CITSMART_LINK_NOVO", {}).get("value", "https://suporte.mpms.mp.br/inbox/lowcode/form/copilot_novo/default"), key="cfg_citsmart_novo"),
            False, "urls", "URL do formulário de abertura de chamados do CitSmart"
        )
        form_values["OTRS_LINK"] = (
            st.text_input("URL do OTRS", value=settings.get("OTRS_LINK", {}).get("value", "https://central.mpms.mp.br"), key="cfg_otrs_link"),
            False, "urls", "URL do OTRS"
        )
        c1, c2 = st.columns(2)
        with c1:
            form_values["ATOS_NORMAS_API_URL"] = (
                st.text_input("API de Atos e Normas", value=settings.get("ATOS_NORMAS_API_URL", {}).get("value", "https://www.mpms.mp.br/atos-e-normas/listAll"), key="cfg_atos_api"),
                False, "urls", "Endpoint da API de Atos e Normas MPMS"
            )
        with c2:
            form_values["ATOS_NORMAS_DOWNLOAD_URL"] = (
                st.text_input("URL Download de Atos e Normas", value=settings.get("ATOS_NORMAS_DOWNLOAD_URL", {}).get("value", "https://www.mpms.mp.br/atos-e-normas/download/"), key="cfg_atos_dl"),
                False, "urls", "URL base para download de Atos e Normas"
            )

    # -------------------------------------------------------------------------
    # TAB 6: SharePoint & Planilhas
    # -------------------------------------------------------------------------
    elif selected_subtab == "📂 SharePoint & Planilhas":
        st.markdown("#### 📂 Sincronização SharePoint / OneDrive")
        st.info("Caminhos locais sincronizados ou links compartilhados do SharePoint.")

        form_values["SHAREPOINT_RELATIVE_PATH"] = (
            st.text_area("Caminho Relativo da Planilha de Chamados Unificados", value=settings.get("SHAREPOINT_RELATIVE_PATH", {}).get("value", r"OneDrive - Ministerio Público do Estado de Mato Grosso do Sul\Documentos SharePoint DIT-Manutenção\Chamados\Chamados_Unificados_Final.xlsx"), key="cfg_sp_chamados", height=70),
            False, "sharepoint", "Caminho relativo da Planilha de Chamados no OneDrive/SharePoint"
        )
        c1, c2 = st.columns(2)
        with c1:
            form_values["DONATIONS_EXCEL_RELATIVE_PATH"] = (
                st.text_input("URL Planilha Doações e Baixas", value=settings.get("DONATIONS_EXCEL_RELATIVE_PATH", {}).get("value", ""), key="cfg_sp_doacoes"),
                False, "sharepoint", "URL/Caminho da Planilha de Doações e Baixas"
            )
            form_values["FISCAL_EXCEL_RELATIVE_PATH"] = (
                st.text_input("URL Planilha Fiscalização de Contratos", value=settings.get("FISCAL_EXCEL_RELATIVE_PATH", {}).get("value", ""), key="cfg_sp_fiscal"),
                False, "sharepoint", "URL/Caminho da Planilha de Fiscalização de Contratos"
            )
        with c2:
            form_values["WARRANTY_EXCEL_RELATIVE_PATH"] = (
                st.text_input("URL Planilha Controle de Garantia", value=settings.get("WARRANTY_EXCEL_RELATIVE_PATH", {}).get("value", ""), key="cfg_sp_garantia"),
                False, "sharepoint", "URL/Caminho da Planilha de Garantia"
            )
            form_values["VIAGENS_EXCEL_RELATIVE_PATH"] = (
                st.text_input("URL Planilha de Viagens da Bancada", value=settings.get("VIAGENS_EXCEL_RELATIVE_PATH", {}).get("value", "https://ministeriopublicoms.sharepoint.com/:x:/s/dit-manutencao/IQBQonhEup-USL9RTQTk8W6vAS9SWFl5jhGXAOR_bezve5E?e=PYfupT"), key="cfg_sp_viagens"),
                False, "sharepoint", "URL/Caminho da Planilha de Viagens da Bancada"
            )
            form_values["SHAREPOINT_MATUTINO_URL"] = (
                st.text_input("URL Planilha Escala Matutina", value=settings.get("SHAREPOINT_MATUTINO_URL", {}).get("value", ""), key="cfg_sp_matutino"),
                False, "sharepoint", "URL da Planilha de Escala Matutina"
            )

        c3, c4 = st.columns(2)
        with c3:
            form_values["VIDEO_FAQ_PATH"] = (
                st.text_input("Pasta de Vídeos FAQ (SharePoint)", value=settings.get("VIDEO_FAQ_PATH", {}).get("value", ""), key="cfg_sp_videos"),
                False, "sharepoint", "URL da pasta de Vídeos FAQ no SharePoint"
            )
        with c4:
            form_values["IMAGE_FAQ_PATH"] = (
                st.text_input("Pasta de Imagens FAQ (SharePoint)", value=settings.get("IMAGE_FAQ_PATH", {}).get("value", ""), key="cfg_sp_imagens"),
                False, "sharepoint", "URL da pasta de Imagens FAQ no SharePoint"
            )

    # -------------------------------------------------------------------------
    # TAB 7: IA Gemini
    # -------------------------------------------------------------------------
    elif selected_subtab == "🤖 Inteligência Artificial (Gemini)":
        st.markdown("#### 🤖 Google Gemini AI")
        st.info("Chave de API do modelo generativo para enriquecimento e categorização de chamados.")

        form_values["GEMINI_API_KEY"] = (
            st.text_input("Chave de API do Gemini (API Key)", value=settings.get("GEMINI_API_KEY", {}).get("value", ""), type="password", key="cfg_gemini_key", help="Armazenada de forma criptografada no banco"),
            True, "ia", "Chave da API do Google Gemini"
        )

    # -------------------------------------------------------------------------
    # TAB 8: WhatsApp & Alertas (Evolution API)
    # -------------------------------------------------------------------------
    elif selected_subtab == "📱 WhatsApp & Alertas (Evolution)":
        st.markdown("#### 📱 Integração WhatsApp (Evolution API)")
        st.info("Conecte o telefone institucional da Bancada STI (+55 67 98478-2034) para envio de alertas automáticos D-1 (12h nos dias úteis) para a equipe.")

        from src.services.evolution_client import (
            get_connection_status,
            get_qr_code,
            disconnect_instance,
            send_whatsapp_text
        )
        from src.database import (
            get_whatsapp_destinatarios,
            toggle_whatsapp_destinatario_status,
            get_whatsapp_disparos_log
        )

        col_w1, col_w2 = st.columns(2)
        with col_w1:
            form_values["EVOLUTION_API_URL"] = (
                st.text_input("URL Base da Evolution API", value=settings.get("EVOLUTION_API_URL", {}).get("value", "http://evolution-api:8080"), key="cfg_evo_url"),
                False, "whatsapp", "URL do serviço Evolution API"
            )
            form_values["EVOLUTION_INSTANCE_NAME"] = (
                st.text_input("Nome da Instância", value=settings.get("EVOLUTION_INSTANCE_NAME", {}).get("value", "bancada_sti"), key="cfg_evo_inst"),
                False, "whatsapp", "Nome da Instância do WhatsApp na Evolution"
            )
        with col_w2:
            form_values["EVOLUTION_API_KEY"] = (
                st.text_input("Chave de Autenticação (API Key)", value=settings.get("EVOLUTION_API_KEY", {}).get("value", "bancada_secret_token_123"), type="password", key="cfg_evo_key"),
                True, "whatsapp", "Token de autenticação da Evolution API"
            )
            form_values["WHATSAPP_INSTITUCIONAL_NUMERO"] = (
                st.text_input("Telefone Institucional da Bancada", value=settings.get("WHATSAPP_INSTITUCIONAL_NUMERO", {}).get("value", "+55 67 98478-2034"), key="cfg_evo_tel"),
                False, "whatsapp", "Número de telefone institucional da bancada"
            )

        st.markdown("---")
        st.markdown("##### 📡 Status da Sessão & Pareamento QR Code")

        # Verifica status em tempo real
        with st.spinner("Consultando status na Evolution API..."):
            status_data = get_connection_status()

        state = status_data.get("state", "offline")
        is_open = state == "open"
        is_connecting = state == "connecting"

        col_st_badge, col_btn_qr, col_btn_dc = st.columns([2, 1, 1])
        with col_st_badge:
            if is_open:
                st.success("🟢 **WhatsApp Conectado e Pronto!** (Instância ativa)")
            elif is_connecting:
                st.warning("🟡 **Aguardando Leitura do QR Code...**")
            elif status_data.get("online"):
                st.info(f"⚪ **Status Atual:** `{state.upper()}` (Desconectado)")
            else:
                st.error(f"🔴 **Evolution API Offline ou Inacessível** ({status_data.get('error', 'Sem resposta')})")

        with col_btn_qr:
            if st.button("📲 Gerar / Exibir QR Code", use_container_width=True):
                with st.spinner("Solicitando novo QR Code à Evolution API..."):
                    qr_res = get_qr_code()
                    if qr_res.get("success") and qr_res.get("base64"):
                        st.session_state["whatsapp_qr_base64"] = qr_res["base64"]
                        st.session_state["whatsapp_pairing_code"] = qr_res.get("pairing_code", "")
                        st.toast("QR Code gerado com sucesso!", icon="📱")
                    elif qr_res.get("success") and qr_res.get("count", 0) > 0:
                        st.info("Aguardando geração do QR Code... tente novamente em alguns segundos.")
                    else:
                        st.error(f"Erro ao obter QR Code: {qr_res.get('error', 'Sem resposta')}")

        with col_btn_dc:
            if is_open and st.button("🚪 Desconectar Sessão", type="secondary", use_container_width=True):
                if disconnect_instance():
                    st.session_state.pop("whatsapp_qr_base64", None)
                    st.toast("Sessão do WhatsApp encerrada com sucesso.", icon="🚪")
                    st.rerun()
                else:
                    st.error("Falha ao desconectar instância.")

        # Renderização do QR Code se disponível na sessão
        if "whatsapp_qr_base64" in st.session_state and not is_open:
            b64_data = st.session_state["whatsapp_qr_base64"]
            if b64_data.startswith("data:image"):
                img_src = b64_data
            else:
                img_src = f"data:image/png;base64,{b64_data}"

            st.markdown(r"""
                <div style="text-align: center; background: #0f172a; padding: 20px; border-radius: 12px; border: 1px solid #334155; margin: 15px 0;">
                    <h4 style="color: #38bdf8; margin-bottom: 8px;">Aponte a câmera do WhatsApp para conectar</h4>
                    <p style="color: #94a3b8; font-size: 13px;">Abra o WhatsApp no celular > Aparelhos Conectados > Conectar um Aparelho</p>
                </div>
            """, unsafe_allow_html=True)
            
            c_qr_left, c_qr_center, c_qr_right = st.columns([1, 1, 1])
            with c_qr_center:
                st.image(img_src, caption="QR Code da Bancada STI", width=280)
                if st.session_state.get("whatsapp_pairing_code"):
                    st.code(f"Código de Pareamento: {st.session_state['whatsapp_pairing_code']}")
                if st.button("🔄 Atualizar Status de Conexão", use_container_width=True):
                    st.rerun()

        st.markdown("---")
        st.markdown("##### 👥 Destinatários Autorizados da Bancada")
        st.caption("Apenas os integrantes abaixo recebem os disparos automáticos de plantões, viagens e portarias:")

        dest_df = get_whatsapp_destinatarios(only_active=False)
        if not dest_df.empty:
            for _, r in dest_df.iterrows():
                d_id = int(r["id"])
                d_nome = r["nome"]
                d_tel = r["telefone"]
                d_ativo = bool(r["ativo"])

                c_d1, c_d2, c_d3, c_d4 = st.columns([3, 2, 2, 2])
                with c_d1:
                    st.markdown(f"**{d_nome}**")
                with c_d2:
                    st.markdown(f"`+{d_tel[:2]} {d_tel[2:4]} {d_tel[4:9]}-{d_tel[9:]}`" if len(d_tel) == 13 else f"`+{d_tel}`")
                with c_d3:
                    new_active = st.checkbox("Ativo p/ Alertas", value=d_ativo, key=f"dest_act_{d_id}")
                    if new_active != d_ativo:
                        toggle_whatsapp_destinatario_status(d_id, new_active)
                        st.toast(f"Status de {d_nome} atualizado!", icon="👥")
                        st.rerun()
                with c_d4:
                    if st.button("📨 Testar Envio", key=f"btn_test_w_{d_id}", disabled=not is_open):
                        with st.spinner("Enviando teste..."):
                            t_msg = f"🔔 *Teste de Conexão - Bancada STI*\n\nOlá, *{d_nome.split()[0]}*! Esta é uma mensagem de teste enviada pela Evolution API da Bancada de Atendimento STI."
                            send_res = send_whatsapp_text(d_tel, t_msg)
                            if send_res.get("success"):
                                st.toast(f"Mensagem de teste enviada para {d_nome.split()[0]}!", icon="✅")
                            else:
                                st.error(f"Erro ao enviar: {send_res.get('error')}")

        st.markdown("---")
        st.markdown("##### ⏰ Execução Manual do Scheduler D-1 (12:00)")
        c_sch1, c_sch2 = st.columns([2, 1])
        with c_sch1:
            st.caption("Dispara manualmente a verificação das escalas do dia seguinte útil (Sexta-feira cobre Sáb/Dom/Seg).")
        with c_sch2:
            if st.button("🚀 Executar Alertas D-1 Agora", use_container_width=True, disabled=not is_open):
                with st.spinner("Processando e enviando alertas WhatsApp..."):
                    from src.syncs.sync_whatsapp_scheduler import run_whatsapp_scheduler
                    res_sch = run_whatsapp_scheduler(dry_run=False, force=False)
                    st.toast(f"Scheduler finalizado! Enviados: {res_sch.get('sent_count', 0)}", icon="🚀")
                    st.success(f"Resultado: {res_sch}")

        # Histórico de disparos
        with st.expander("📋 Ver Histórico Recente de Disparos WhatsApp"):
            logs_df = get_whatsapp_disparos_log(limit=25)
            if not logs_df.empty:
                st.dataframe(logs_df, use_container_width=True)
            else:
                st.info("Nenhum disparo registrado até o momento.")

    # -------------------------------------------------------------------------
    # TAB 9: Agendamentos & Cron Jobs
    # -------------------------------------------------------------------------
    elif selected_subtab == "⏰ Agendamentos & Cron Jobs":
        st.markdown("#### ⏰ Agendador de Tarefas Periódicas (Cron Jobs)")
        st.write("Configure a frequência de execução em segundo plano para sincronizações de planilhas, varreduras de portarias e disparos do WhatsApp.")

        from src.database import (
            get_cron_schedules,
            get_cron_schedule_by_id,
            update_cron_schedule,
            get_recent_cron_logs
        )
        from src.services.cron_scheduler import get_cron_daemon
        from src.components.status_banner import read_log_lines
        import pandas as pd

        def format_br_datetime(val) -> str:
            """Formata datas ISO para padrão brasileiro DD/MM/AAAA HH:MM:SS."""
            if not val or pd.isna(val) or str(val).strip().lower() in ["none", "nan", "nunca", ""]:
                return "Nunca"
            try:
                dt = pd.to_datetime(val)
                return dt.strftime("%d/%m/%Y %H:%M:%S")
            except Exception:
                return str(val)

        # Mapeamento de logs específicos para cada tarefa
        TASK_LOG_PATH_MAP = {
            "sync_portarias": Path("debug_logs") / "faq" / "sync_portarias.log",
            "sync_viagens": Path("debug_logs") / "viagens" / "viagens_sync.log",
            "sync_plantoes": Path("debug_logs") / "plantoes" / "sync_alerts.log",
            "sync_fiscalizacao": Path("debug_logs") / "fiscalizacao" / "sync.log",
            "orquestrador_chamados": Path("debug_logs") / "orquestrador" / "orquestrador.log",
            "whatsapp_d1": Path("debug_logs") / "plantoes" / "whatsapp_scheduler.log",
        }

        daemon = get_cron_daemon()
        is_daemon_alive = daemon.is_alive()

        # Barra de Status do Motor
        col_m1, col_m2 = st.columns([3, 1])
        with col_m1:
            if is_daemon_alive:
                st.success("🟢 **Motor de Agendamento Ativo** (Executando em segundo plano no container)")
            else:
                st.warning("🟡 **Motor de Agendamento Inativo**")
        with col_m2:
            if not is_daemon_alive:
                if st.button("▶️ Iniciar Motor", use_container_width=True):
                    daemon.start()
                    st.toast("Motor Cron iniciado!", icon="🟢")
                    st.rerun()

        st.markdown("---")
        st.markdown("##### ⚙️ Configuração das Rotinas Automáticas")

        schedules_df = get_cron_schedules()

        if not schedules_df.empty:
            for _, r in schedules_df.iterrows():
                task_id = str(r["task_id"])
                nome = str(r["nome"])
                cat = str(r["categoria"])
                ativo = bool(r["ativo"])
                tipo = str(r["tipo_agendamento"])
                int_val = int(r["intervalo_valor"] or 2)
                int_uni = str(r["intervalo_unidade"] or "horas")
                hora_fixa = str(r["horario_fixo"] or "12:00")
                dias_uteis = bool(r["apenas_dias_uteis"])
                ult_exec_raw = r.get("ultima_execucao")
                ult_exec = format_br_datetime(ult_exec_raw)
                ult_status = str(r.get("ultimo_status") or "pendente")
                ult_log = str(r.get("ultimo_log") or "")
                desc = str(r.get("descricao") or "")

                is_task_executing = (ult_status.lower() == "executando") or (task_id in daemon._executing_tasks)
                if task_id == "orquestrador_chamados" and not is_task_executing:
                    from src.components.status_banner import check_orquestrador_running
                    is_task_executing = check_orquestrador_running()

                status_color = "#3b82f6" if is_task_executing else ("#22c55e" if ult_status == "sucesso" else ("#ef4444" if ult_status == "erro" else "#f59e0b"))

                with st.container(border=True):
                    # Cabeçalho da Tarefa
                    c_h1, c_h2, c_h3 = st.columns([4, 2, 2])
                    with c_h1:
                        st.markdown(f"**{nome}**")
                        st.caption(f"📁 Categoria: `{cat}` | {desc}")
                    with c_h2:
                        st.caption(f"Última Execução: **{ult_exec}**")
                        st.markdown(f"Status: <span style='color:{status_color}; font-weight:bold;'>{ult_status.upper()}</span>", unsafe_allow_html=True)
                    with c_h3:
                        if st.button("🚀 Executar Agora", key=f"btn_run_cron_{task_id}", use_container_width=True, disabled=is_task_executing):
                            with st.spinner(f"Iniciando {task_id}..."):
                                if daemon.trigger_task_now(task_id):
                                    st.toast(f"Tarefa {task_id} disparada com sucesso!", icon="🚀")
                                else:
                                    st.warning("Esta tarefa já está em execução.")
                                st.rerun()

                    # Controles de Frequência
                    col_f1, col_f2, col_f3, col_f4 = st.columns([2, 3, 3, 2])
                    with col_f1:
                        new_ativo = st.checkbox("Habilitada", value=ativo, key=f"cron_act_{task_id}")
                    with col_f2:
                        tipo_opts = ["intervalo", "horario_fixo"]
                        new_tipo = st.selectbox("Modo", tipo_opts, index=0 if tipo == "intervalo" else 1, format_func=lambda x: "Recorrente (Intervalo)" if x == "intervalo" else "Horário Fixo Diário", key=f"cron_tipo_{task_id}")

                    if new_tipo == "intervalo":
                        with col_f3:
                            new_val = st.number_input("A cada", min_value=1, max_value=720, value=int_val, key=f"cron_val_{task_id}")
                        with col_f4:
                            uni_opts = ["minutos", "horas", "dias"]
                            idx_uni = uni_opts.index(int_uni) if int_uni in uni_opts else 1
                            new_uni = st.selectbox("Unidade", uni_opts, index=idx_uni, key=f"cron_uni_{task_id}")
                        new_hora = hora_fixa
                        new_dias = dias_uteis
                    else:
                        with col_f3:
                            new_hora = st.text_input("Horário (HH:MM)", value=hora_fixa, max_chars=5, key=f"cron_hf_{task_id}", help="Exemplo: 12:00")
                        with col_f4:
                            new_dias = st.checkbox("Apenas Dias Úteis", value=dias_uteis, key=f"cron_uteis_{task_id}")
                        new_val = int_val
                        new_uni = int_uni

                    # Botão para salvar alterações desta tarefa
                    if (new_ativo != ativo or new_tipo != tipo or new_val != int_val or
                        new_uni != int_uni or new_hora != hora_fixa or new_dias != dias_uteis):
                        if st.button("💾 Salvar Agendamento", key=f"btn_save_cron_{task_id}", type="primary"):
                            update_cron_schedule(task_id, new_ativo, new_tipo, new_val, new_uni, new_hora, new_dias)
                            st.toast(f"Agendamento de {nome} atualizado!", icon="💾")
                            st.rerun()

                    if ult_log and not is_task_executing and str(ult_log).lower() not in ["none", "nan", ""]:
                        st.caption(f"📝 **Último Registro:** `{ult_log[:200]}`")

                    # Accordion em tempo real quando o status estiver como EXECUTANDO
                    if is_task_executing:
                        with st.expander(f"⏳ Processando {nome} em segundo plano...", expanded=True):
                            st.info("A tarefa está sendo executada pelo motor agendador. Acompanhe os logs abaixo:")
                            
                            @st.fragment(run_every="3s")
                            def show_task_live_logs(t_id=task_id):
                                # Verifica se a tarefa já concluiu
                                current_sched = get_cron_schedule_by_id(t_id)
                                sched_done = current_sched and str(current_sched.get("ultimo_status", "")).lower() != "executando"
                                daemon_done = t_id not in daemon._executing_tasks
                                orq_done = True
                                if t_id == "orquestrador_chamados":
                                    from src.components.status_banner import check_orquestrador_running
                                    orq_done = not check_orquestrador_running()

                                if sched_done and daemon_done and orq_done:
                                    # Concluiu! Dá um rerun completo na página para fechar o accordion e reativar o botão
                                    st.rerun()
                                    return

                                log_p = TASK_LOG_PATH_MAP.get(t_id, Path("debug_logs") / "sync" / "cron_scheduler.log")
                                log_txt = read_log_lines(log_p, n=20)
                                st.code(log_txt, language="text")
                                if st.button("🔄 Atualizar Visualização", key=f"btn_refresh_exec_{t_id}"):
                                    st.rerun()
                                    
                            show_task_live_logs()

        # Histórico de Execuções
        st.markdown("---")
        with st.expander("📋 Ver Histórico Recente de Execuções Automáticas (Logs)", expanded=False):
            logs_cron_df = get_recent_cron_logs(limit=30)
            if not logs_cron_df.empty:
                # Formata Início e Fim para padrão brasileiro DD/MM/AAAA HH:MM:SS
                df_display = logs_cron_df.copy()
                if "inicio" in df_display.columns:
                    df_display["inicio"] = df_display["inicio"].apply(format_br_datetime)
                if "fim" in df_display.columns:
                    df_display["fim"] = df_display["fim"].apply(format_br_datetime)
                st.dataframe(df_display, use_container_width=True)
            else:
                st.info("Nenhuma execução registrada no agendador até o momento.")

    # Botão de Ação Global para Salvar a aba ativa (apenas se houver campos de formulário)
    if form_values:
        st.markdown("---")
        col_save, col_spacer = st.columns([1, 4])
        with col_save:
            if st.button("💾 Salvar Configurações", type="primary", use_container_width=True):
                salvos = 0
                for key, (val, is_sec, cat, desc) in form_values.items():
                    if set_setting(key, val, is_secret=is_sec, category=cat, description=desc):
                        salvos += 1
                
                st.success(f"🎉 **Configurações salvas com sucesso!**")
                st.toast("Configurações atualizadas e senhas criptografadas no banco!", icon="🔒")
                st.rerun()
