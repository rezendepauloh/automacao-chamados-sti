import os
import sys
import re
import time
import shutil
import threading
import subprocess
import webbrowser
import keyring
from pathlib import Path
import streamlit as st
import streamlit.components.v1 as components
from src.components.status_banner import render_log_expander

from src.config import (
    PS_SCRIPT_ANALISADOR,
    PS_SCRIPT_MANUTENCAO,
    PS_SCRIPT_REMOVER_USUARIOS,
    USER_HOME,
    DEBUG_DIR_SCRIPTS,
    setup_logging
)

logger = setup_logging(DEBUG_DIR_SCRIPTS / "scripts_automacao.log", __name__)

ANSI_ESCAPE = re.compile(r'\x1b(?:[@-Z\\-_]|\[[0-?]*[ -/]*[@-~])')

# Armazenamento global persistente de tarefas em segundo plano no processo Python
if not hasattr(sys, "_ps_background_jobs"):
    sys._ps_background_jobs = {}

_BACKGROUND_JOBS = sys._ps_background_jobs

SUBTAB_MAP = {
    "analisador": "📊 Analisador de Dispositivos",
    "manutencao": "🧹 Manutenção e Limpeza Remota",
    "perfis": "👥 Remoção de Perfis de Usuário"
}
SUBTAB_REVERSE = {v: k for k, v in SUBTAB_MAP.items()}


def _get_powershell_exe() -> str:
    """
    Detecta se o PowerShell Core 7+ (pwsh.exe) está disponível no sistema.
    Caso contrário, faz o fallback para o Windows PowerShell 5.1 (powershell.exe).
    """
    pwsh_path = shutil.which("pwsh")
    if pwsh_path:
        return pwsh_path
    return "powershell.exe"


def _read_file_safe_utf8(file_path: Path) -> str:
    """
    Lê um arquivo de texto garantindo a decodificação correta de caracteres acentuados
    e corrigindo eventuais problemas de codificação dupla (ex: Ã³ -> ó).
    """
    if not file_path or not file_path.exists():
        return ""

    with open(file_path, "rb") as f:
        raw_bytes = f.read()

    try:
        content = raw_bytes.decode("utf-8-sig")
    except UnicodeDecodeError:
        try:
            content = raw_bytes.decode("cp1252")
        except UnicodeDecodeError:
            content = raw_bytes.decode("latin-1", errors="replace")

    if "Ã" in content or "â€" in content:
        try:
            content = content.encode("latin-1").decode("utf-8")
        except Exception:
            pass

    return content


def _ensure_cred_admin_xml():
    """
    Verifica e regenera o cred_admin.xml em TODAS as pastas de scripts caso o arquivo
    não possa ser descriptografado pelo usuário atual do Windows (falha DPAPI).
    Utiliza as credenciais salvas do SCCM_ADMIN_USER no cofre do Windows (keyring).
    """
    admin_user = os.getenv("SCCM_ADMIN_USER", "")
    if not admin_user:
        return

    script_paths = [PS_SCRIPT_ANALISADOR, PS_SCRIPT_MANUTENCAO, PS_SCRIPT_REMOVER_USUARIOS]
    target_xmls = set()
    for sp in script_paths:
        if sp and sp.exists():
            target_xmls.add(sp.parent / "cred_admin.xml")

    ps_exe = _get_powershell_exe()
    creationflags = subprocess.CREATE_NO_WINDOW if sys.platform == "win32" else 0

    admin_pass = None

    for target_xml in target_xmls:
        need_regenerate = False
        if not target_xml.exists():
            logger.info(f"O arquivo '{target_xml}' não existe. Será gerado...")
            need_regenerate = True
        else:
            test_ps = f"try {{ $c = Import-Clixml -Path '{target_xml}'; if ($c.UserName) {{ exit 0 }} else {{ exit 1 }} }} catch {{ exit 1 }}"
            res = subprocess.run(
                [ps_exe, "-NonInteractive", "-NoProfile", "-Command", test_ps],
                capture_output=True,
                creationflags=creationflags
            )
            if res.returncode != 0:
                logger.warning(f"⚠️ cred_admin.xml em '{target_xml}' não pode ser descriptografado. Necessário regenerar.")
                need_regenerate = True

        if need_regenerate:
            if not admin_pass:
                admin_pass = keyring.get_password("sccm_admin", admin_user) or keyring.get_password("sccm", admin_user)

            if not admin_pass:
                logger.error("❌ Senha do SCCM_ADMIN_USER não encontrada no keyring. Não é possível gerar cred_admin.xml.")
                continue

            domain_user = admin_user if ("\\" in admin_user or "@" in admin_user) else f"mpe\\{admin_user}"
            escaped_pass = admin_pass.replace('"', '`"').replace('$', '`$')

            gen_ps = (
                f'$sec = ConvertTo-SecureString "{escaped_pass}" -AsPlainText -Force; '
                f'$cred = New-Object System.Management.Automation.PSCredential ("{domain_user}", $sec); '
                f'$cred | Export-Clixml -Path "{target_xml}" -Force'
            )

            gen_res = subprocess.run(
                [ps_exe, "-NonInteractive", "-NoProfile", "-Command", gen_ps],
                capture_output=True,
                text=True,
                creationflags=creationflags
            )

            if gen_res.returncode == 0:
                logger.info(f"✅ cred_admin.xml regenerado com sucesso em '{target_xml}'!")
            else:
                logger.error(f"❌ Falha ao gerar cred_admin.xml em '{target_xml}': {gen_res.stderr}")


def _get_latest_html_report(output_dir: Path) -> Path | None:
    """Busca o relatório HTML mais recente gerado no diretório especificado."""
    if not output_dir.exists():
        return None
    
    html_files = list(output_dir.glob("*.html"))
    if not html_files:
        return None
    
    html_files.sort(key=lambda f: f.stat().st_mtime, reverse=True)
    return html_files[0]


def start_background_ps_job(job_id: str, script_name: str, host: str, script_path: Path, args: list, out_folder: Path = None):
    """
    Dispara o script PowerShell em uma thread/processo em segundo plano desacoplada da sessão do Streamlit.
    """
    if not script_path or not script_path.exists():
        st.error(f"❌ Arquivo de script não encontrado no caminho: `{script_path}`")
        return None

    _ensure_cred_admin_xml()
    ps_exe = _get_powershell_exe()

    formatted_args = [f'"{a}"' if " " in str(a) else str(a) for a in args]
    args_str = " ".join(formatted_args)
    ps_cmd_str = (
        f"[Console]::OutputEncoding = [System.Text.Encoding]::UTF8; "
        f"$OutputEncoding = [System.Text.Encoding]::UTF8; "
        f"& '{script_path}' {args_str}"
    )

    cmd = [ps_exe, "-NonInteractive", "-ExecutionPolicy", "Bypass", "-Command", ps_cmd_str]

    job_data = {
        "job_id": job_id,
        "script_name": script_name,
        "host": host,
        "script_path": script_path,
        "out_folder": out_folder,
        "status": "running",
        "logs": [f"🚀 [{time.strftime('%H:%M:%S')}] Processo iniciado em segundo plano ({ps_exe})..."],
        "start_time": time.time(),
        "end_time": None,
        "return_code": None
    }
    # Limita o histórico de robôs/scripts a no máximo 3 itens (remove o mais antigo ao adicionar o 4º)
    while len(_BACKGROUND_JOBS) >= 3:
        oldest_key = next(iter(_BACKGROUND_JOBS))
        _BACKGROUND_JOBS.pop(oldest_key, None)

    _BACKGROUND_JOBS[job_id] = job_data

    def _worker():
        logger.info(f"🚀 [BACKGROUND TASK START] {job_id} ({script_name}) host={host}")
        try:
            creationflags = subprocess.CREATE_NO_WINDOW if sys.platform == "win32" else 0
            process = subprocess.Popen(
                cmd,
                stdout=subprocess.PIPE,
                stderr=subprocess.STDOUT,
                text=True,
                encoding="utf-8",
                errors="replace",
                bufsize=1,
                creationflags=creationflags
            )

            while True:
                line = process.stdout.readline()
                if not line and process.poll() is not None:
                    break
                if line:
                    cleaned_line = ANSI_ESCAPE.sub('', line.rstrip())
                    if "Ã" in cleaned_line or "â€" in cleaned_line:
                        try:
                            cleaned_line = cleaned_line.encode("latin-1").decode("utf-8")
                        except Exception:
                            pass
                    job_data["logs"].append(cleaned_line)
                    logger.info(f"[{job_id}] {cleaned_line}")

            rc = process.poll()
            job_data["return_code"] = rc
            job_data["end_time"] = time.time()

            if rc == 0:
                job_data["status"] = "complete"
                job_data["logs"].append("✅ Execução concluída com sucesso!")
                logger.info(f"✅ [BACKGROUND TASK COMPLETE] {job_id}")
            else:
                job_data["status"] = "error"
                job_data["logs"].append(f"⚠️ Script finalizado com código de retorno: {rc}")
                logger.warning(f"⚠️ [BACKGROUND TASK ERROR] {job_id} return_code={rc}")

        except Exception as e:
            job_data["status"] = "error"
            job_data["end_time"] = time.time()
            job_data["logs"].append(f"❌ Exceção na thread do script: {e}")
            logger.error(f"❌ [BACKGROUND TASK EXCEPTION] {job_id}: {e}")

    t = threading.Thread(target=_worker, daemon=True)
    t.start()
    return job_data


def render_background_jobs_widget():
    """
    Renderiza o widget expansível com o status e logs de scripts rodando em segundo plano (máximo 3).
    """
    # Garante limite máximo de 3 accordions ativas, removendo a mais antiga se houver excedente
    while len(_BACKGROUND_JOBS) > 3:
        oldest_key = next(iter(_BACKGROUND_JOBS))
        _BACKGROUND_JOBS.pop(oldest_key, None)

    if not _BACKGROUND_JOBS:
        return

    st.markdown("### 🤖 Robôs & Scripts em Segundo Plano")

    for job_id, job in list(_BACKGROUND_JOBS.items()):
        status = job["status"]
        script_name = job["script_name"]
        host = job["host"]
        out_folder = job.get("out_folder")

        if status == "running":
            header_label = f"⏳ Rodando: {script_name} em '{host}'"
            exp_state = True
        elif status == "complete":
            header_label = f"✅ Concluído: {script_name} em '{host}'"
            exp_state = False
        else:
            header_label = f"❌ Erro: {script_name} em '{host}'"
            exp_state = True

        with st.expander(header_label, expanded=exp_state):
            if status == "running":
                st.info("⚡ O script está sendo executado em segundo plano. Você pode continuar usando o painel normalmente!")

                # Fragmento de auto-atualização para scripts em execução
                @st.fragment(run_every="3s")
                def auto_refresh_job_logs():
                    displayed_logs = "\n".join(job["logs"][-150:])
                    st.code(displayed_logs, language="powershell")
                    if st.button("🔄 Atualizar Logs", key=f"btn_refresh_{job_id}"):
                        st.rerun()
                auto_refresh_job_logs()
            else:
                if status == "complete":
                    st.success("🎉 Execução finalizada com sucesso!")
                else:
                    st.error("⚠️ O script foi finalizado com erros.")

                displayed_logs = "\n".join(job["logs"][-150:])
                st.code(displayed_logs, language="powershell")

            col_actions, col_clear = st.columns([3, 1])
            with col_clear:
                if st.button("🗑️ Limpar / Dispensar", key=f"btn_clear_{job_id}"):
                    _BACKGROUND_JOBS.pop(job_id, None)
                    st.rerun()

            # Exibe botões do relatório HTML se for o Analisador de Dispositivos e tiver concluído
            if out_folder and Path(out_folder).exists():
                latest_html = _get_latest_html_report(Path(out_folder))
                if latest_html and latest_html.exists():
                    st.markdown("---")
                    st.markdown(f"#### 📄 Relatório HTML Gerado (`{latest_html.name}`)")
                    col_btn1, col_btn2 = st.columns(2)
                    with col_btn1:
                        if st.button("🌐 Abrir no Navegador", key=f"bg_open_html_{job_id}", width='stretch'):
                            try:
                                os.startfile(str(latest_html))
                            except Exception:
                                webbrowser.open(f"file:///{latest_html}")
                            st.toast("Relatório HTML aberto no navegador!", icon="🌐")
                    with col_btn2:
                        html_content = _read_file_safe_utf8(latest_html)
                        st.download_button(
                            label="📥 Baixar Relatório HTML",
                            data=html_content,
                            file_name=latest_html.name,
                            mime="text/html",
                            width='stretch',
                            key=f"bg_download_html_{job_id}"
                        )
                    with st.expander("👁️ Pré-visualizar Relatório HTML no Painel", expanded=False):
                        html_content = _read_file_safe_utf8(latest_html)
                        st.components.v1.html(html_content, height=700, scrolling=True)


def render_scripts_automacao_page():
    """Renderiza a página principal de execução dos scripts de automação PowerShell."""
    ps_engine = _get_powershell_exe()
    engine_badge = "⚡ PowerShell 7+ (pwsh)" if "pwsh" in ps_engine.lower() else "💻 Windows PowerShell 5.1"

    st.markdown(f"""
        <div style="background: var(--metric-bg, #1e293b); padding: 20px; border-radius: 12px; border-left: 6px solid #3b82f6; border-top: 1px solid var(--metric-border, #2d3139); border-right: 1px solid var(--metric-border, #2d3139); border-bottom: 1px solid var(--metric-border, #2d3139); margin-bottom: 20px; box-shadow: 0 2px 8px rgba(0,0,0,0.08);">
            <div style="display: flex; justify-content: space-between; align-items: center;">
                <h2 style="color: var(--metric-value-color, #ffffff); margin: 0; font-size: 24px; font-weight: 700;">⚡ Scripts de Automação PowerShell</h2>
                <span style="background-color: var(--secondary-background-color, rgba(56, 189, 248, 0.15)); color: var(--text-color, #38bdf8); font-size: 12px; font-weight: 600; padding: 4px 10px; border-radius: 20px; border: 1px solid #0284c7;">
                    {engine_badge}
                </span>
            </div>
            <p style="color: var(--metric-title-color, #94a3b8); margin: 6px 0 0 0; font-size: 14px;">
                Execute rotinas remotas em segundo plano com suporte a navegação livre, F5 e acompanhamento de logs em tempo real.
            </p>
        </div>
    """, unsafe_allow_html=True)


    # Renderiza o widget de tarefas em segundo plano (se houver)
    render_background_jobs_widget()

    # Suporte a URL query param ?subtab=slug
    url_subtab = st.query_params.get("subtab", "analisador")
    default_subtab_title = SUBTAB_MAP.get(url_subtab, "📊 Analisador de Dispositivos")
    subtab_list = list(SUBTAB_MAP.values())
    default_index = subtab_list.index(default_subtab_title) if default_subtab_title in subtab_list else 0

    selected_subtab_title = st.radio(
        "Seleção de Script",
        options=subtab_list,
        index=default_index,
        horizontal=True,
        key="nav_scripts_subtab_radio",
        label_visibility="collapsed"
    )

    new_sub_slug = SUBTAB_REVERSE.get(selected_subtab_title, "analisador")
    if st.query_params.get("subtab") != new_sub_slug:
        st.query_params["subtab"] = new_sub_slug

    st.markdown("<div style='margin-bottom: 15px;'></div>", unsafe_allow_html=True)

    # =========================================================================
    # TAB 1: Analisador de Dispositivos
    # =========================================================================
    if selected_subtab_title == "📊 Analisador de Dispositivos":
        st.markdown("### 📊 Analisador de Dispositivos de Máquina Remota")
        st.caption("Coleta inventário completo de hardware, drivers, serviços e gera relatórios em HTML, PDF e Excel.")

        col1, col2 = st.columns([2, 1])
        with col1:
            comp_analisador = st.text_input("💻 Nome ou IP da Máquina Remota", key="input_analisador_host", placeholder="Ex: PJCHA-54491 ou 10.x.x.x")
        with col2:
            default_out = str(USER_HOME / "DeviceReports")
            out_folder = st.text_input("📁 Pasta de Destino dos Relatórios", value=default_out, key="input_analisador_out")

        col_opt1, col_opt2 = st.columns([1, 2])
        with col_opt1:
            timeout_sec = st.number_input("⏱️ Timeout CIM (segundos)", min_value=10, max_value=300, value=30, step=5)
        with col_opt2:
            skip_major = st.checkbox("⚡ Coleta Rápida (Ignorar Drivers, Programas e Serviços)", value=False)

        st.markdown("---")
        if st.button("🚀 Executar Análise de Dispositivo em Segundo Plano", type="primary", key="btn_run_analisador", width='stretch'):
            if not comp_analisador.strip():
                st.warning("⚠️ Por favor, informe o Nome ou IP da máquina remota.")
            else:
                args = ["-ComputerName", comp_analisador.strip(), "-OutputFolder", out_folder.strip(), "-TimeoutSec", str(timeout_sec)]
                if skip_major:
                    args.append("-SkipMajorData")

                job_id = f"analisador_{comp_analisador.strip()}_{int(time.time())}"
                start_background_ps_job(
                    job_id=job_id,
                    script_name="Analisador de Dispositivos",
                    host=comp_analisador.strip(),
                    script_path=PS_SCRIPT_ANALISADOR,
                    args=args,
                    out_folder=Path(out_folder.strip())
                )
                st.toast(f"🚀 Análise da máquina {comp_analisador.strip()} iniciada em segundo plano!", icon="🤖")
                st.rerun()

    # =========================================================================
    # TAB 2: Manutenção e Limpeza Remota
    # =========================================================================
    elif selected_subtab_title == "🧹 Manutenção e Limpeza Remota":
        st.markdown("### 🧹 Manutenção e Limpeza Remota de Estação")
        st.caption("Executa limpeza de arquivos temporários, caches, defrag/otimização de disco e reparo de volumes via WinRM.")

        comp_manutencao = st.text_input("💻 Nome ou IP da Máquina Remota", key="input_manutencao_host", placeholder="Ex: MPE-58063 ou 10.x.x.x")

        st.markdown("#### ⚙️ Opções de Limpeza Profunda")
        col_m1, col_m2, col_m3, col_m4 = st.columns(4)
        with col_m1:
            clean_win_old = st.checkbox("🗑️ Windows.old", value=False, help="Remove pasta C:\\Windows.old de atualizações anteriores")
        with col_m2:
            clean_do = st.checkbox("📦 Cache Delivery Opt.", value=False, help="Limpa o cache de Otimização de Entrega")
        with col_m3:
            clean_dumps = st.checkbox("💥 Crash Dumps & WER", value=False, help="Remove relatórios de erro do Windows e minidumps")
        with col_m4:
            verbose_mode = st.checkbox("🔍 Modo Verbose", value=True, help="Exibe detalhes estendidos durante o progresso")

        st.markdown("---")
        if st.button("🚀 Executar Manutenção e Limpeza em Segundo Plano", type="primary", key="btn_run_manutencao", width='stretch'):
            if not comp_manutencao.strip():
                st.warning("⚠️ Por favor, informe o Nome ou IP da máquina remota.")
            else:
                args = ["-ComputerName", comp_manutencao.strip()]
                if clean_win_old:
                    args.append("-CleanWindowsOld")
                if clean_do:
                    args.append("-CleanDeliveryOptimization")
                if clean_dumps:
                    args.append("-CleanCrashDumps")
                if verbose_mode:
                    args.append("-Verbose")

                job_id = f"manutencao_{comp_manutencao.strip()}_{int(time.time())}"
                start_background_ps_job(
                    job_id=job_id,
                    script_name="Manutenção e Limpeza Remota",
                    host=comp_manutencao.strip(),
                    script_path=PS_SCRIPT_MANUTENCAO,
                    args=args
                )
                st.toast(f"🚀 Limpeza remota da máquina {comp_manutencao.strip()} iniciada em segundo plano!", icon="🧹")
                st.rerun()

    # =========================================================================
    # TAB 3: Remoção de Perfis de Usuário
    # =========================================================================
    elif selected_subtab_title == "👥 Remoção de Perfis de Usuário":
        st.markdown("### 👥 Remoção Remota de Perfis de Usuários")
        st.caption("Remove com segurança contas locais e pastas de perfis inativos (C:\\Users) via Win32_UserProfile e StdRegProv.")

        comp_perfis = st.text_input("💻 Nome ou IP da Máquina Remota", key="input_perfis_host", placeholder="Ex: PGJ-58099 ou 10.x.x.x")
        users_input = st.text_area("👤 Logins dos Usuários a Remover", key="input_perfis_users", placeholder="Digite os logins separados por vírgula ou em linhas separadas.\nEx: pedrofonseca, amandabarbosa, lohanlima", height=100)

        st.markdown("---")
        if st.button("🚀 Remover Perfis em Segundo Plano", type="primary", key="btn_run_perfis", width='stretch'):
            if not comp_perfis.strip():
                st.warning("⚠️ Por favor, informe o Nome ou IP da máquina remota.")
            elif not users_input.strip():
                st.warning("⚠️ Por favor, informe ao menos um login de usuário para remoção.")
            else:
                raw_users = users_input.replace("\n", ",").replace(";", ",")
                users_list = [u.strip() for u in raw_users.split(",") if u.strip()]
                users_arg_str = ",".join(users_list)

                args = ["-ComputerName", comp_perfis.strip(), "-UsersToPurge", users_arg_str]

                job_id = f"perfis_{comp_perfis.strip()}_{int(time.time())}"
                start_background_ps_job(
                    job_id=job_id,
                    script_name="Remoção de Perfis de Usuários",
                    host=comp_perfis.strip(),
                    script_path=PS_SCRIPT_REMOVER_USUARIOS,
                    args=args
                )
                st.toast(f"🚀 Remoção de perfis na máquina {comp_perfis.strip()} iniciada em segundo plano!", icon="👥")
                st.rerun()
