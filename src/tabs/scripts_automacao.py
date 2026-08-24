import os
import sys
import re
import time
import shutil
import tempfile
import threading
import subprocess
import webbrowser
import keyring
from pathlib import Path
import streamlit as st
import streamlit.components.v1 as components
from src.components.status_banner import render_log_expander

from src.config import (
    PS_SCRIPTS_DIR,
    PS_SCRIPT_ANALISADOR,
    PS_SCRIPT_MANUTENCAO,
    PS_SCRIPT_REMOVER_USUARIOS,
    USER_HOME,
    DEBUG_DIR_SCRIPTS,
    setup_logging
)
from src.terminal import print_header, CYAN

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


def _get_powershell_exe(selected_option: str = None) -> str:
    """
    Detecta se o PowerShell Core 7+ (pwsh.exe) ou Windows PowerShell 5.1 (powershell.exe)
    deve ser utilizado com base na opção selecionada ou detecção automática.
    """
    if selected_option == "PowerShell 7+ (pwsh.exe)":
        return shutil.which("pwsh") or "pwsh.exe"
    elif selected_option == "Windows PowerShell 5.1 (powershell.exe)":
        return "powershell.exe"

    pwsh_path = shutil.which("pwsh")
    if pwsh_path:
        return pwsh_path
    return "powershell.exe"


def _clean_ansi(text: str) -> str:
    """Remove sequências de escape ANSI de uma string para exibição limpa no Streamlit."""
    return ANSI_ESCAPE.sub('', text)


def _read_file_safe_utf8(filepath: Path) -> str:
    """Lê um arquivo de texto testando encodings comuns para evitar UnicodeDecodeError."""
    if not filepath.exists():
        return ""
    for enc in ["utf-8", "utf-8-sig", "latin-1", "cp1252"]:
        try:
            return filepath.read_text(encoding=enc)
        except UnicodeDecodeError:
            continue
    return filepath.read_text(encoding="utf-8", errors="replace")


def _ensure_cred_admin_xml():
    """
    Verifica e regenera o cred_admin.xml na pasta interna src/scripts_powershell/ e nas subpastas dos scripts
    caso o arquivo não possa ser descriptografado pelo usuário atual do Windows (falha DPAPI).
    Utiliza as credenciais salvas do SCCM_ADMIN_USER no cofre do Windows (keyring).
    """
    admin_user = os.getenv("SCCM_ADMIN_USER", "")
    if not admin_user:
        return

    script_paths = [PS_SCRIPT_ANALISADOR, PS_SCRIPT_MANUTENCAO, PS_SCRIPT_REMOVER_USUARIOS]
    target_xmls = set()
    # Adiciona explicitamente o cred_admin.xml da pasta raiz interna de scripts
    if PS_SCRIPTS_DIR:
        PS_SCRIPTS_DIR.mkdir(parents=True, exist_ok=True)
        target_xmls.add(PS_SCRIPTS_DIR / "cred_admin.xml")

    # Adiciona o cred_admin.xml dentro de cada subpasta específica
    for sp in script_paths:
        if sp:
            sp.parent.mkdir(parents=True, exist_ok=True)
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
            ps_kwargs = {"capture_output": True}
            if sys.platform == "win32":
                ps_kwargs["creationflags"] = subprocess.CREATE_NO_WINDOW
            res = subprocess.run(
                [ps_exe, "-NonInteractive", "-NoProfile", "-Command", test_ps],
                **ps_kwargs
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

            gen_kwargs = {"capture_output": True, "text": True}
            if sys.platform == "win32":
                gen_kwargs["creationflags"] = subprocess.CREATE_NO_WINDOW
            gen_res = subprocess.run(
                [ps_exe, "-NonInteractive", "-NoProfile", "-Command", gen_ps],
                **gen_kwargs
            )

            if gen_res.returncode == 0:
                logger.info(f"✅ cred_admin.xml regenerado com sucesso em '{target_xml}'!")
            else:
                logger.error(f"❌ Falha ao gerar cred_admin.xml em '{target_xml}': {gen_res.stderr}")


def _get_latest_report_files(output_dir: Path) -> dict:
    """Busca os relatórios mais recentes (HTML, PDF, XLSX) gerados no diretório de saída."""
    res = {"html": None, "pdf": None, "xlsx": None}
    if not output_dir.exists():
        return res
    
    for ext in ["html", "pdf", "xlsx"]:
        files = list(output_dir.glob(f"*.{ext}"))
        if files:
            files.sort(key=lambda f: f.stat().st_mtime, reverse=True)
            res[ext] = files[0]
    return res


def start_background_ps_job(job_id: str, script_name: str, host: str, script_path: Path, args: list, out_folder: Path = None, ps_option: str = None):
    """
    Dispara o script PowerShell em uma thread/processo em segundo plano desacoplada da sessão do Streamlit,
    copiando apenas a subpasta específica da ferramenta para um diretório temporário exclusivo e limpando-o após a execução.
    """
    print_header("SCRIPTS DE AUTOMAÇÃO - POWERSHELL", color=CYAN)
    if not script_path or not script_path.exists():
        st.error(f"❌ Arquivo de script não encontrado no caminho: `{script_path}`")
        return None

    _ensure_cred_admin_xml()
    ps_exe = _get_powershell_exe(ps_option)

    # Identifica a subpasta dedicada do script (ex: src/scripts_powershell/analisador/)
    script_source_dir = script_path.parent
    if not script_source_dir.exists():
        script_source_dir = PS_SCRIPTS_DIR

    # Cria diretório temporário exclusivo e copia APENAS a subpasta específica da ferramenta
    try:
        temp_dir = Path(tempfile.mkdtemp(prefix=f"ps_{script_source_dir.name}_"))
        logger.info(f"📁 Criando diretório temporário exclusivo para [{script_name}]: {temp_dir}")
        logger.info(f"📥 Copiando módulo dedicado de `{script_source_dir}` para `{temp_dir}`...")
        shutil.copytree(script_source_dir, temp_dir, dirs_exist_ok=True)
        target_script_path = temp_dir / script_path.name
    except Exception as copy_err:
        logger.error(f"Erro ao copiar subpasta de scripts para diretório temporário: {copy_err}")
        temp_dir = script_source_dir
        target_script_path = script_path

    formatted_args = [f'"{a}"' if " " in str(a) else str(a) for a in args]
    args_str = " ".join(formatted_args)
    ps_cmd_str = (
        f"[Console]::OutputEncoding = [System.Text.Encoding]::UTF8; "
        f"$OutputEncoding = [System.Text.Encoding]::UTF8; "
        f"& '{target_script_path}' {args_str}"
    )

    cmd = [ps_exe, "-NoProfile", "-NonInteractive", "-ExecutionPolicy", "Bypass", "-Command", ps_cmd_str]

    job_data = {
        "job_id": job_id,
        "script_name": script_name,
        "host": host,
        "script_path": script_path,
        "out_folder": out_folder,
        "status": "running",
        "logs": [
            f"🚀 [{time.strftime('%H:%M:%S')}] Processo iniciado em segundo plano ({ps_exe})...",
            f"📁 [{time.strftime('%H:%M:%S')}] Diretório temporário exclusivo criado: `{temp_dir}`",
            f"📥 [{time.strftime('%H:%M:%S')}] Módulo e arquivos de suporte copiados de `{script_source_dir}`"
        ],
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
        logger.info(f"⚙️ Disparando processo no diretório de trabalho temporário (cwd): {temp_dir}")
            popen_kwargs = {
                "stdout": subprocess.PIPE,
                "stderr": subprocess.STDOUT,
                "text": True,
                "encoding": "utf-8",
                "errors": "replace",
                "bufsize": 1,
                "cwd": str(temp_dir),
            }
            if sys.platform == "win32":
                popen_kwargs["creationflags"] = subprocess.CREATE_NO_WINDOW
            process = subprocess.Popen(
                cmd,
                **popen_kwargs
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

                    # Auditoria de caminhos de arquivos gerados (HTML, PDF, XLSX e Diretório)
                    lower_line = cleaned_line.lower()
                    if (
                        "relatorio_html_path:" in lower_line
                        or "relatorio_pdf_path:" in lower_line
                        or "relatorio_excel_path:" in lower_line
                        or "relatório salvo em:" in lower_line
                        or "pdf gerado em:" in lower_line
                        or "arquivo excel gerado em:" in lower_line
                        or "diretório de saída" in lower_line
                    ):
                        logger.info(f"🎯 [AUDITORIA DE DESTINO] [{job_id}] {cleaned_line}")

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
        finally:
            if temp_dir and temp_dir != script_source_dir and temp_dir.exists():
                try:
                    logger.info(f"🧹 Limpando e removendo diretório temporário: {temp_dir}")
                    job_data["logs"].append(f"🧹 Diretório temporário e seu conteúdo foram removidos com segurança: `{temp_dir}`")
                    shutil.rmtree(temp_dir, ignore_errors=True)
                except Exception as clean_err:
                    logger.warning(f"⚠️ Erro ao remover diretório temporário {temp_dir}: {clean_err}")

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

            # Exibe botões dos relatórios gerados (HTML, PDF, XLSX) prontos para download na máquina do usuário
            if out_folder and Path(out_folder).exists():
                reports = _get_latest_report_files(Path(out_folder))
                latest_html = reports["html"]
                latest_pdf = reports["pdf"]
                latest_xlsx = reports["xlsx"]

                if latest_html or latest_pdf or latest_xlsx:
                    st.markdown("---")
                    st.markdown("#### 📥 Relatórios Gerados (Baixar no seu Computador)")
                    
                    dl_cols = st.columns(3)
                    if latest_html and latest_html.exists():
                        with dl_cols[0]:
                            html_bytes = latest_html.read_bytes()
                            st.download_button(
                                label="🌐 Baixar Relatório HTML",
                                data=html_bytes,
                                file_name=latest_html.name,
                                mime="text/html",
                                width='stretch',
                                key=f"bg_download_html_{job_id}"
                            )
                    if latest_pdf and latest_pdf.exists():
                        with dl_cols[1]:
                            pdf_bytes = latest_pdf.read_bytes()
                            st.download_button(
                                label="📄 Baixar Relatório PDF",
                                data=pdf_bytes,
                                file_name=latest_pdf.name,
                                mime="application/pdf",
                                width='stretch',
                                key=f"bg_download_pdf_{job_id}"
                            )
                    if latest_xlsx and latest_xlsx.exists():
                        with dl_cols[2]:
                            xlsx_bytes = latest_xlsx.read_bytes()
                            st.download_button(
                                label="📊 Baixar Relatório Excel",
                                data=xlsx_bytes,
                                file_name=latest_xlsx.name,
                                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                                width='stretch',
                                key=f"bg_download_xlsx_{job_id}"
                            )

                    if latest_html and latest_html.exists():
                        with st.expander("👁️ Pré-visualizar Relatório HTML no Painel", expanded=False):
                            html_content = _read_file_safe_utf8(latest_html)
                            st.components.v1.html(html_content, height=700, scrolling=True)



def render_scripts_automacao_page():
    """Renderiza a página principal de execução dos scripts de automação PowerShell."""
    col_hdr1, col_hdr2 = st.columns([2, 1])
    with col_hdr2:
        selected_ps_version = st.selectbox(
            "⚙️ Interpretador PowerShell",
            options=[
                "Detectar Automaticamente (Padrão)",
                "PowerShell 7+ (pwsh.exe)",
                "Windows PowerShell 5.1 (powershell.exe)"
            ],
            key="ps_version_selector"
        )
    with col_hdr1:
        ps_engine = _get_powershell_exe(selected_ps_version)
        engine_badge = "⚡ PowerShell 7+ (pwsh)" if "pwsh" in ps_engine.lower() else "💻 Windows PowerShell 5.1"

        st.markdown(f"""
            <div style="background: var(--metric-bg, #1e293b); padding: 15px 20px; border-radius: 12px; border-left: 6px solid #3b82f6; border-top: 1px solid var(--metric-border, #2d3139); border-right: 1px solid var(--metric-border, #2d3139); border-bottom: 1px solid var(--metric-border, #2d3139); margin-bottom: 10px; box-shadow: 0 2px 8px rgba(0,0,0,0.08);">
                <div style="display: flex; justify-content: space-between; align-items: center;">
                    <h2 style="color: var(--metric-value-color, #ffffff); margin: 0; font-size: 22px; font-weight: 700;">⚡ Scripts de Automação PowerShell</h2>
                    <span style="background-color: var(--secondary-background-color, rgba(56, 189, 248, 0.15)); color: var(--text-color, #38bdf8); font-size: 12px; font-weight: 600; padding: 4px 10px; border-radius: 20px; border: 1px solid #0284c7;">
                        {engine_badge}
                    </span>
                </div>
                <p style="color: var(--metric-title-color, #94a3b8); margin: 4px 0 0 0; font-size: 13px;">
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
        st.caption("Coleta inventário completo de hardware, drivers, serviços e gera relatórios em HTML, PDF e Excel para download direto na sua máquina.")

        col1, col2 = st.columns([2, 1])
        with col1:
            comp_analisador = st.text_input("💻 Nome ou IP da Máquina Remota", key="input_analisador_host", placeholder="Ex: PJCHA-54491 ou 10.x.x.x")
        with col2:
            user_profile = os.getenv("USERPROFILE") or str(Path.home())
            default_out = str(Path(user_profile) / "DeviceReports")
            out_folder = st.text_input(
                "📁 Pasta de Saída no Servidor",
                value=default_out,
                key="input_analisador_out",
                help="Os relatórios serão gerados nesta pasta no servidor e disponibilizados automaticamente nos botões de Download para você salvar onde quiser na sua máquina."
            )


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
                chosen_out = out_folder.strip()
                if not chosen_out or chosen_out == default_out or chosen_out == str(Path.home() / "DeviceReports"):
                    out_path_arg = r"$env:USERPROFILE\DeviceReports"
                    out_folder_obj = Path(default_out)
                else:
                    out_path_arg = chosen_out
                    out_folder_obj = Path(chosen_out)

                args = ["-ComputerName", comp_analisador.strip(), "-OutputFolder", out_path_arg, "-TimeoutSec", str(timeout_sec)]
                if skip_major:
                    args.append("-SkipMajorData")

                job_id = f"analisador_{comp_analisador.strip()}_{int(time.time())}"
                start_background_ps_job(
                    job_id=job_id,
                    script_name="Analisador de Dispositivos",
                    host=comp_analisador.strip(),
                    script_path=PS_SCRIPT_ANALISADOR,
                    args=args,
                    out_folder=out_folder_obj,
                    ps_option=selected_ps_version
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
                    args=args,
                    ps_option=selected_ps_version
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
                    args=args,
                    ps_option=selected_ps_version
                )
                st.toast(f"🚀 Remoção de perfis na máquina {comp_perfis.strip()} iniciada em segundo plano!", icon="👥")
                st.rerun()
