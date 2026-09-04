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
from urllib.parse import urlencode
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
    Detecta se o PowerShell Core 7+ (pwsh) ou Windows PowerShell 5.1 (powershell.exe)
    deve ser utilizado com base na opção selecionada ou detecção automática.
    """
    if selected_option == "PowerShell 7+ (pwsh.exe)":
        return shutil.which("pwsh") or shutil.which("pwsh.exe") or "pwsh"
    elif selected_option == "Windows PowerShell 5.1 (powershell.exe)":
        return (
            shutil.which("powershell.exe") or
            shutil.which("/mnt/c/Windows/System32/WindowsPowerShell/v1.0/powershell.exe") or
            shutil.which("/mnt/c/WINDOWS/System32/WindowsPowerShell/v1.0/powershell.exe") or
            "powershell.exe"
        )

    pwsh_path = shutil.which("pwsh") or shutil.which("pwsh.exe")
    if pwsh_path:
        return pwsh_path
    return (
        shutil.which("powershell.exe") or
        shutil.which("/mnt/c/Windows/System32/WindowsPowerShell/v1.0/powershell.exe") or
        shutil.which("/mnt/c/WINDOWS/System32/WindowsPowerShell/v1.0/powershell.exe") or
        "powershell.exe"
    )


def dispatch_bancada_uri(tool: str, host: str, extra_params: dict = None) -> str:
    r"""
    Aciona o Protocol Handler 'bancada://' no Windows do usuário através do navegador.
    O Windows abre o bancada-launcher.ps1 no PowerShell do usuário, executa o script na máquina alvo
    usando os privilégios/cmdlets locais do Windows, salva o relatório em %USERPROFILE%\DeviceReports
    e limpa os arquivos temporários automaticamente.
    """
    params = {
        "tool": tool,
        "host": host,
        "server": os.getenv("HOST_IP") or "localhost"
    }
    if extra_params:
        params.update(extra_params)

    query_str = urlencode(params)
    bancada_url = f"bancada://run?{query_str}"

    # Log detalhado no console do servidor (visível no terminal do Docker / 00-iniciar.sh)
    engine_log = params.get("ps_engine", "auto")
    print(f"\n{'='*70}", flush=True)
    print(f"🚀 [DISPATCH PROTOCOLO] Enviando comando para estação do usuário via bancada://", flush=True)
    print(f"   • Ferramenta   : {tool}", flush=True)
    print(f"   • Alvo Remoto  : {host}", flush=True)
    print(f"   • Interpretador: {engine_log}", flush=True)
    print(f"   • URL Gerada   : {bancada_url}", flush=True)
    print(f"{'='*70}\n", flush=True)

    # Injeta trigger JavaScript imediato usando iframe invisível para invocar o protocolo
    js_code = f"""
    <script>
        (function() {{
            try {{
                const iframe = document.createElement('iframe');
                iframe.style.display = 'none';
                iframe.src = '{bancada_url}';
                document.body.appendChild(iframe);
                setTimeout(() => {{ iframe.remove(); }}, 3000);
            }} catch (e) {{
                window.location.href = '{bancada_url}';
            }}
        }})();
    </script>
    """
    components.html(js_code, height=0, width=0)
    return bancada_url


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
            test_ps = f"try {{ $c = Import-Clixml -Path '{target_xml}'; if ($c.UserName) {{ exit 0 }} else {{ exit 1 }} }} catch {{ exit 1 }}"
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
                try:
                    from src.database.settings_db import get_setting
                    admin_pass = get_setting("SCCM_ADMIN_PASSWORD")
                except Exception:
                    pass
                if not admin_pass:
                    admin_pass = keyring.get_password("sccm_admin", admin_user) or keyring.get_password("sccm", admin_user) or os.getenv("SCCM_ADMIN_PASSWORD")

            if not admin_pass:
                logger.error("❌ Senha do SCCM_ADMIN_USER não encontrada no banco ou keyring. Não é possível gerar cred_admin.xml.")
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


def _resolve_local_path(path_str: str) -> Path:
    r"""Converte caminhos no estilo Windows (C:\... ou \\wsl$\...) para caminhos acessíveis no Linux/WSL/Docker se necessário."""
    if not path_str:
        return None
    p = str(path_str).strip()
    if re.match(r'^[a-zA-Z]:[\\/]', p):
        drive = p[0].lower()
        rest = p[2:].replace('\\', '/').lstrip('/')
        wsl_cand = Path(f"/mnt/{drive}/{rest}")
        if wsl_cand.exists():
            return wsl_cand
    direct_p = Path(p)
    if direct_p.exists():
        return direct_p
    return None


def _get_latest_report_files(output_dir: Path, generated_files: dict = None) -> dict:
    """Busca os relatórios mais recentes (HTML, PDF, XLSX) gerados no diretório de saída ou identificados no log."""
    res = {"html": None, "pdf": None, "xlsx": None}
    
    # 1. Primeiro verifica os caminhos capturados diretamente da saída do script
    if generated_files:
        for ext in ["html", "pdf", "xlsx"]:
            path_val = generated_files.get(ext)
            if path_val:
                resolved = _resolve_local_path(path_val)
                if resolved and resolved.exists():
                    res[ext] = resolved

    # 2. Se algum formato ainda não foi encontrado, busca no output_dir
    resolved_out = _resolve_local_path(str(output_dir)) if output_dir else None
    if resolved_out and resolved_out.exists():
        for ext in ["html", "pdf", "xlsx"]:
            if not res[ext]:
                files = list(resolved_out.glob(f"*.{ext}"))
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
        "generated_files": {"html": None, "pdf": None, "xlsx": None},
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
        try:
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

                    # Auditoria e captura de caminhos de arquivos gerados (HTML, PDF, XLSX e Diretório)
                    lower_line = cleaned_line.lower()
                    if "relatorio_html_path:" in lower_line:
                        path_part = cleaned_line.split(":", 1)[1].strip()
                        job_data["generated_files"]["html"] = path_part
                        logger.info(f"🎯 [AUDITORIA DE DESTINO HTML] [{job_id}] {path_part}")
                    elif "relatorio_pdf_path:" in lower_line:
                        path_part = cleaned_line.split(":", 1)[1].strip()
                        job_data["generated_files"]["pdf"] = path_part
                        logger.info(f"🎯 [AUDITORIA DE DESTINO PDF] [{job_id}] {path_part}")
                    elif "relatorio_excel_path:" in lower_line:
                        path_part = cleaned_line.split(":", 1)[1].strip()
                        job_data["generated_files"]["xlsx"] = path_part
                        logger.info(f"🎯 [AUDITORIA DE DESTINO EXCEL] [{job_id}] {path_part}")
                    elif (
                        "relatório salvo em:" in lower_line
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
                    Execute rotinas remotas com 1 clique direto no Windows cliente, sem fricção, com salvamento local dos relatórios e limpeza automática dos scripts.
                </p>
            </div>
        """, unsafe_allow_html=True)

    # Informações e instalador do disparador para técnicos (1 clique)
    launcher_installer = Path(__file__).parent.parent / "protocol_handler" / "instalar_disparador_windows.cmd"
    if launcher_installer.exists():
        with st.expander("🛠️ Configuração do Disparador Windows (Apenas 1ª vez por máquina)", expanded=False):
            st.info("Para que o navegador dispare os scripts nativamente no Windows sem precisar de intervenção manual, execute o instalador abaixo uma única vez na máquina do técnico:")
            st.download_button(
                label="📥 Baixar Instalador do Disparador Windows (.cmd)",
                data=launcher_installer.read_bytes(),
                file_name="instalar_disparador_windows.cmd",
                mime="application/octet-stream",
                key="btn_download_bancada_launcher_installer"
            )

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

                # Aciona diretamente o Protocol Handler no Windows da máquina cliente (sem executar no servidor Linux)
                engine_param = "pwsh" if "pwsh" in selected_ps_version.lower() else ("powershell" if "5.1" in selected_ps_version else "auto")
                b_url = dispatch_bancada_uri(
                    tool="analisador",
                    host=comp_analisador.strip(),
                    extra_params={
                        "skip_major": "true" if skip_major else "false",
                        "timeout": str(timeout_sec),
                        "ps_engine": engine_param
                    }
                )
                st.success(f"🚀 **Comando enviado para o seu Windows!** O PowerShell local executará a análise na estação **{comp_analisador.strip()}**.")
                st.info(f"📂 O relatório final será gravado diretamente no seu Windows em: `C:\\Users\\<Seu_Usuario>\\DeviceReports`.")
                st.markdown(f"""
                <div style="background-color: rgba(30, 144, 255, 0.1); border: 1px solid #1E90FF; border-radius: 6px; padding: 10px; margin-top: 8px;">
                    <span style="font-size: 0.9rem;">💡 <em>Se a janela preta do PowerShell não tiver aberto automaticamente, seu navegador pode ter bloqueado o disparo.</em></span><br>
                    👉 <a href="{b_url}" style="font-weight: bold; color: #1E90FF; text-decoration: underline;">Clique aqui para abrir manualmente no Windows</a>
                </div>
                """, unsafe_allow_html=True)

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

                # Aciona via Protocol Handler no cliente (sem executar no servidor Linux)
                engine_param = "pwsh" if "pwsh" in selected_ps_version.lower() else ("powershell" if "5.1" in selected_ps_version else "auto")
                b_url = dispatch_bancada_uri(
                    tool="manutencao",
                    host=comp_manutencao.strip(),
                    extra_params={
                        "ps_engine": engine_param
                    }
                )
                st.success(f"🧹 **Comando enviado para o seu Windows!** O PowerShell local executará a manutenção na estação **{comp_manutencao.strip()}**.")
                st.markdown(f"""
                <div style="background-color: rgba(30, 144, 255, 0.1); border: 1px solid #1E90FF; border-radius: 6px; padding: 10px; margin-top: 8px;">
                    <span style="font-size: 0.9rem;">💡 <em>Se a janela do PowerShell não tiver aberto automaticamente:</em></span><br>
                    👉 <a href="{b_url}" style="font-weight: bold; color: #1E90FF; text-decoration: underline;">Clique aqui para abrir manualmente no Windows</a>
                </div>
                """, unsafe_allow_html=True)

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

                # Aciona via Protocol Handler no cliente (sem executar no servidor Linux)
                engine_param = "pwsh" if "pwsh" in selected_ps_version.lower() else ("powershell" if "5.1" in selected_ps_version else "auto")
                b_url = dispatch_bancada_uri(
                    tool="perfis",
                    host=comp_perfis.strip(),
                    extra_params={
                        "users": users_arg_str,
                        "ps_engine": engine_param
                    }
                )
                st.success(f"👥 **Comando enviado para o seu Windows!** O PowerShell local executará a remoção de perfis na estação **{comp_perfis.strip()}**.")
                st.markdown(f"""
                <div style="background-color: rgba(30, 144, 255, 0.1); border: 1px solid #1E90FF; border-radius: 6px; padding: 10px; margin-top: 8px;">
                    <span style="font-size: 0.9rem;">💡 <em>Se a janela do PowerShell não tiver aberto automaticamente:</em></span><br>
                    👉 <a href="{b_url}" style="font-weight: bold; color: #1E90FF; text-decoration: underline;">Clique aqui para abrir manualmente no Windows</a>
                </div>
                """, unsafe_allow_html=True)
