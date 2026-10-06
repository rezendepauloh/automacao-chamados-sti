import re
import io
import pandas as pd
import streamlit as st
from typing import List, Dict, Any, Tuple
from datetime import datetime
from src.database.connection import get_connection, DB_TYPE
from src.database.sccm_db import normalize_hardware_model


def parse_patrimonios_input(raw_text: str) -> List[str]:
    """
    Extrai lista limpa e ordenada de números/códigos de patrimônio a partir de texto livre
    (copiado de colunas do Excel, Teams, separado por vírgula, ponto e vírgula, espaços, etc.).
    Preserva a ordem de inserção do usuário sem duplicatas.
    """
    if not raw_text:
        return []

    # Separa por quebra de linha, vírgula, ponto e vírgula, barra ou tabulação
    tokens = re.split(r'[\r\n,;\t]+', raw_text)
    seen = set()
    cleaned_list = []

    for token in tokens:
        item = str(token).strip()
        if not item:
            continue
        # Remove caracteres indesejados mantendo alfanuméricos e hífens
        item_clean = re.sub(r'[^a-zA-Z0-9\-_]', '', item)
        if item_clean and item_clean not in seen:
            seen.add(item_clean)
            cleaned_list.append(item_clean)

    return cleaned_list


def _clean_ou_path(dn_or_ou: str) -> str:
    """Extrai uma representação humana e legível de uma OU a partir de seu DN."""
    if not dn_or_ou or pd.isna(dn_or_ou):
        return ""
    raw = str(dn_or_ou).strip()
    if not raw or raw.lower() in ["none", "nan", "null"]:
        return ""

    # Extrai todas as OUs do DN (ex: OU=Notebooks,OU=PGJ,OU=MPE,DC=... -> Notebooks > PGJ > MPE)
    ou_matches = re.findall(r'OU=([^,]+)', raw, flags=re.IGNORECASE)
    if ou_matches:
        return " > ".join(ou_matches)
    return raw


def _clean_site_name(site: str) -> str:
    """Higieniza o nome do Site SCCM/AD."""
    if not site or pd.isna(site):
        return ""
    s = str(site).strip()
    return "" if s.lower() in ["none", "nan", "null", "-"] else s


def format_brazilian_datetime(val: Any) -> str:
    """
    Formata valores de data e hora para o padrão brasileiro: DD/MM/AAAA HH:MM:SS.
    Suporta:
      - WMI CIM DateTime (SCCM): 20260929080059.657000+*** ou 20260929080059
      - ISO-8601: 2026-09-29 08:00:59, 2026-09-29T08:00:59
      - Datetime/Timestamp do Python ou strings YYYY-MM-DD
    """
    if val is None or pd.isna(val):
        return "-"
    if isinstance(val, datetime) or (hasattr(pd, "Timestamp") and isinstance(val, getattr(pd, "Timestamp"))):
        return val.strftime("%d/%m/%Y %H:%M:%S")

    raw = str(val).strip()
    if not raw or raw.lower() in ["none", "nan", "null", "nat", "-", ""]:
        return "-"

    # 1. Padrão WMI / CIM DateTime do SCCM (ex: 20260929080059.657000+***)
    clean_digits = raw.split(".")[0].strip()
    if len(clean_digits) == 14 and clean_digits.isdigit():
        try:
            return datetime.strptime(clean_digits, "%Y%m%d%H%M%S").strftime("%d/%m/%Y %H:%M:%S")
        except Exception:
            pass

    # 2. Formatos comuns de data e hora
    for fmt in [
        "%Y-%m-%d %H:%M:%S",
        "%Y-%m-%dT%H:%M:%S",
        "%Y-%m-%d %H:%M:%S.%f",
        "%Y-%m-%dT%H:%M:%S.%f",
        "%d/%m/%Y %H:%M:%S",
        "%d/%m/%Y %H:%M"
    ]:
        try:
            return datetime.strptime(raw[:19], fmt).strftime("%d/%m/%Y %H:%M:%S")
        except Exception:
            continue

    # 3. Apenas data (YYYY-MM-DD ou DD/MM/AAAA)
    for fmt_d in ["%Y-%m-%d", "%d/%m/%Y"]:
        try:
            return datetime.strptime(raw[:10], fmt_d).strftime("%d/%m/%Y 00:00:00")
        except Exception:
            continue

    # 4. Fallback flexível via pd.to_datetime
    try:
        dt = pd.to_datetime(raw, errors="coerce")
        if pd.notna(dt):
            return dt.strftime("%d/%m/%Y %H:%M:%S")
    except Exception:
        pass

    return raw


def query_dmp_patrimonio_report(patrimonios: List[str]) -> Tuple[pd.DataFrame, Dict[str, Any]]:
    """
    Realiza o cruzamento multi-origem no banco de dados local:
      1. ad_cache_computers (Active Directory)
      2. sccm_cache_devices (SCCM Inventário)
      3. ad_cache_users (Dados completos de quem logou por último)
      4. equipamentos_doados (Histórico de Doação / Baixa / Garantia)

    Retorna o DataFrame final e um dicionário de métricas/resumo.
    """
    if not patrimonios:
        return pd.DataFrame(), {}

    conn = get_connection()
    cursor = conn.cursor()
    p_ph = "?" if DB_TYPE not in ["postgres", "postgresql"] else "%s"

    # 1. Carrega base de doações para cruzamento
    doacoes_dict: Dict[str, Dict[str, Any]] = {}
    try:
        cursor.execute("SELECT patrimonio, modelo, serial_number, tipo_movimentacao, data_movimentacao, chamado, motivo_baixa FROM equipamentos_doados")
        d_rows = cursor.fetchall()
        for r in d_rows:
            pat_d = str(r[0] or "").strip()
            if pat_d:
                # Mantém a ocorrência mais recente se houver múltiplas
                doacoes_dict[pat_d.lower()] = {
                    "modelo": str(r[1] or "").strip(),
                    "serial": str(r[2] or "").strip(),
                    "tipo_mov": str(r[3] or "").strip(),
                    "data_mov": str(r[4] or "").strip(),
                    "chamado": str(r[5] or "").strip(),
                    "motivo": str(r[6] or "").strip()
                }
    except Exception:
        pass

    # 2. Carrega todos os computadores do AD para busca flexível
    ad_comps_map: Dict[str, Dict[str, Any]] = {}
    try:
        cursor.execute("SELECT name, dns_hostname, operating_system, os_version, description, parent_ou_dn, is_active, last_logon, dn FROM ad_cache_computers")
        for row in cursor.fetchall():
            c_name = str(row[0] or "").strip()
            if c_name:
                ad_comps_map[c_name.lower()] = {
                    "name": c_name,
                    "dns_hostname": str(row[1] or "").strip(),
                    "operating_system": str(row[2] or "").strip(),
                    "os_version": str(row[3] or "").strip(),
                    "description": str(row[4] or "").strip(),
                    "parent_ou_dn": str(row[5] or "").strip(),
                    "is_active": row[6],
                    "last_logon": str(row[7] or "").strip(),
                    "dn": str(row[8] or "").strip()
                }
    except Exception:
        pass

    # 3. Carrega todos os dispositivos SCCM
    sccm_devs_map: Dict[str, Dict[str, Any]] = {}
    try:
        cursor.execute("SELECT name, last_logon_user, ip_addresses, manufacturer, model, operating_system, ad_site_name, client_active, last_active_time, raw_json FROM sccm_cache_devices")
        for row in cursor.fetchall():
            s_name = str(row[0] or "").strip()
            if s_name:
                sccm_devs_map[s_name.lower()] = {
                    "name": s_name,
                    "last_logon_user": str(row[1] or "").strip(),
                    "ip_addresses": str(row[2] or "").strip(),
                    "manufacturer": str(row[3] or "").strip(),
                    "model": normalize_hardware_model(str(row[4] or "").strip()),
                    "operating_system": str(row[5] or "").strip(),
                    "ad_site_name": _clean_site_name(str(row[6] or "").strip()),
                    "client_active": row[7],
                    "last_active_time": str(row[8] or "").strip(),
                    "raw_json": row[9]
                }
    except Exception:
        pass

    # 4. Cache dos usuários do AD para mapear login -> Nome Completo, Setor, Office
    ad_users_map: Dict[str, Dict[str, Any]] = {}
    try:
        cursor.execute("SELECT sam_account_name, display_name, department, office, company, title FROM ad_cache_users")
        for u_row in cursor.fetchall():
            sam = str(u_row[0] or "").strip().lower()
            if sam:
                ad_users_map[sam] = {
                    "display_name": str(u_row[1] or "").strip(),
                    "department": str(u_row[2] or "").strip(),
                    "office": str(u_row[3] or "").strip(),
                    "company": str(u_row[4] or "").strip(),
                    "title": str(u_row[5] or "").strip()
                }
    except Exception:
        pass

    cursor.close()
    conn.close()

    results = []

    for pat in patrimonios:
        pat_clean = pat.strip()
        pat_num = re.sub(r'^[a-zA-Z\-_]+', '', pat_clean).strip() or pat_clean
        pat_lower = pat_clean.lower()
        pat_num_lower = pat_num.lower()

        # Encontra máquina no AD
        matched_ad = None
        for k, v in ad_comps_map.items():
            if k == pat_lower or k == f"note-{pat_num_lower}" or k == f"mpe-{pat_num_lower}" or pat_num_lower in k:
                matched_ad = v
                break

        # Encontra máquina no SCCM
        matched_sccm = None
        for k, v in sccm_devs_map.items():
            if k == pat_lower or k == f"note-{pat_num_lower}" or k == f"mpe-{pat_num_lower}" or pat_num_lower in k:
                matched_sccm = v
                break

        # Encontra histórico em doações/baixas
        matched_doacao = doacoes_dict.get(pat_num_lower) or doacoes_dict.get(pat_lower)

        # Determina dados consolidados
        hostname = ""
        if matched_ad:
            hostname = matched_ad["name"]
        elif matched_sccm:
            hostname = matched_sccm["name"]
        else:
            hostname = f"NOTE-{pat_num}" if pat_num.isdigit() and len(pat_num) >= 5 else pat_clean

        # Modelo
        modelo = ""
        if matched_sccm and matched_sccm.get("model"):
            modelo = matched_sccm["model"]
        elif matched_doacao and matched_doacao.get("modelo"):
            modelo = matched_doacao["modelo"]
        else:
            modelo = "Não identificado"

        # Serial Number (se presente no histórico de doações ou SCCM)
        serial_number = ""
        if matched_doacao and matched_doacao.get("serial"):
            serial_number = matched_doacao["serial"]
        elif matched_sccm and matched_sccm.get("raw_json"):
            try:
                import json
                rj = json.loads(matched_sccm["raw_json"])
                serial_number = str(rj.get("SerialNumber") or rj.get("BIOSSerialNumber") or "").strip()
            except Exception:
                serial_number = ""

        # Usuário / Responsável
        login_user = ""
        nome_completo = ""
        setor_usuario = ""
        lotacao_usuario = ""

        if matched_sccm and matched_sccm.get("last_logon_user"):
            raw_user = matched_sccm["last_logon_user"]
            login_user = raw_user.split("\\")[-1].strip()
        elif matched_ad and matched_ad.get("description"):
            # Algumas máquinas no AD guardam o responsável no campo description
            desc_val = matched_ad["description"]
            if desc_val and not desc_val.lower().startswith("vm"):
                nome_completo = desc_val

        if login_user:
            u_info = ad_users_map.get(login_user.lower(), {})
            if u_info:
                nome_completo = u_info.get("display_name") or login_user
                setor_usuario = u_info.get("department") or ""
                lotacao_usuario = u_info.get("office") or u_info.get("company") or ""
            else:
                nome_completo = login_user

        # Localidade / Lotação Física
        localidade_resumo = ""
        ou_formatada = ""
        if matched_ad and matched_ad.get("parent_ou_dn"):
            ou_formatada = _clean_ou_path(matched_ad["parent_ou_dn"])

        site_sccm = matched_sccm.get("ad_site_name", "") if matched_sccm else ""

        if lotacao_usuario:
            localidade_resumo = lotacao_usuario
        elif site_sccm:
            localidade_resumo = f"Site SCCM: {site_sccm}"
        elif ou_formatada:
            localidade_resumo = ou_formatada
        else:
            localidade_resumo = "Não localizada"

        # IP
        ip_addr = ""
        if matched_sccm and matched_sccm.get("ip_addresses"):
            ips = [ip.strip() for ip in matched_sccm["ip_addresses"].split(",") if ip.strip() and not ip.strip().startswith("fe80:")]
            ip_addr = ips[0] if ips else matched_sccm["ip_addresses"].split(",")[0].strip()

        # Última Atividade / Logon (formatado no padrão brasileiro DD/MM/AAAA HH:MM:SS)
        ultima_atividade = ""
        if matched_sccm and matched_sccm.get("last_active_time"):
            ultima_atividade = format_brazilian_datetime(matched_sccm["last_active_time"])
        elif matched_ad and matched_ad.get("last_logon"):
            ultima_atividade = format_brazilian_datetime(matched_ad["last_logon"])
        elif matched_doacao and matched_doacao.get("data_mov"):
            dt_mov_fmt = format_brazilian_datetime(matched_doacao["data_mov"])
            ultima_atividade = f"Movimentado em {dt_mov_fmt}" if dt_mov_fmt != "-" else f"Movimentado em {matched_doacao['data_mov']}"

        # Status do Bem para Auditoria
        status_bem = ""
        status_cor = "#64748b" # cinza padrão
        
        if matched_doacao:
            tipo_m = matched_doacao.get("tipo_mov", "Baixa")
            status_bem = f"📦 {tipo_m}"
            status_cor = "#ef4444" if "baixa" in tipo_m.lower() else "#8b5cf6"
        elif matched_ad:
            is_active = matched_ad.get("is_active") == 1
            if not is_active:
                status_bem = "🔴 Desativado no AD"
                status_cor = "#ef4444"
            else:
                # Checa data de inatividade
                last_dt = None
                try:
                    if matched_ad.get("last_logon"):
                        last_dt = pd.to_datetime(matched_ad["last_logon"])
                except Exception:
                    pass

                if last_dt and (datetime.now() - last_dt.to_pydatetime()).days > 90:
                    status_bem = "🟡 Inativo (> 90 dias)"
                    status_cor = "#f59e0b"
                else:
                    status_bem = "🟢 Ativo na Rede"
                    status_cor = "#10b981"
        elif matched_sccm:
            status_bem = "🟢 Ativo (SCCM)"
            status_cor = "#10b981"
        else:
            status_bem = "❓ Não Localizado"
            status_cor = "#64748b"

        results.append({
            "patrimonio": pat_num,
            "hostname": hostname,
            "modelo": modelo,
            "serial_number": serial_number or "-",
            "usuario_login": login_user or "-",
            "usuario_nome": nome_completo or "-",
            "setor": setor_usuario or "-",
            "localidade": localidade_resumo,
            "ou_ad": ou_formatada or "-",
            "ip": ip_addr or "-",
            "site": site_sccm or "-",
            "ultima_atividade": ultima_atividade or "-",
            "status_bem": status_bem,
            "status_cor": status_cor,
            "encontrado": bool(matched_ad or matched_sccm or matched_doacao)
        })

    df_report = pd.DataFrame(results)

    # Resumo
    total_consultado = len(patrimonios)
    localizados = sum(1 for r in results if r["encontrado"])
    nao_localizados = total_consultado - localizados
    ativos = sum(1 for r in results if "ativo" in r["status_bem"].lower())
    baixados = sum(1 for r in results if "doação" in r["status_bem"].lower() or "baixa" in r["status_bem"].lower() or "redistribuição" in r["status_bem"].lower())

    stats = {
        "total": total_consultado,
        "localizados": localizados,
        "nao_localizados": nao_localizados,
        "ativos": ativos,
        "baixados": baixados
    }

    return df_report, stats


def build_dmp_report_html(df_report: pd.DataFrame) -> str:
    """
    Gera código HTML corporativo com design idêntico ao modelo da Redistribuição / Preparo de Chamados,
    pronto para copiar e colar no OTRS, e-mails ou Teams.
    """
    if df_report.empty:
        return "<p>Nenhum equipamento consultado.</p>"

    now_str = datetime.now().strftime("%d/%m/%Y às %H:%M")
    total_equip = len(df_report)
    equip_word = "equipamento" if total_equip == 1 else "equipamentos"

    html_parts = []
    html_parts.append("<div style='font-family: Arial, Helvetica, sans-serif; color: #000000; line-height: 1.5;'>")
    html_parts.append("<p>Prezados,</p>")
    html_parts.append(f"<p>Segue abaixo o levantamento unificado de localização, custódia e situação do lote de <strong>{total_equip} {equip_word}</strong> solicitado pelo Departamento de Material e Patrimônio (DMP):</p>")

    table_html = [
        "<table border='2' cellpadding='6' cellspacing='0' style='border-collapse: collapse; width: 100%; border: 2px solid #cccccc; font-family: Arial, Helvetica, sans-serif; font-size: 12px; color: #000000;'>"
    ]
    th_style = "background-color: #2f5597; border: 2px solid #cccccc; padding: 6px 8px; text-align: left; white-space: nowrap;"

    headers = [
        "#",
        "Patrimônio",
        "Hostname",
        "Modelo",
        "Serial Number",
        "Último Usuário / Responsável",
        "Setor / Lotação",
        "Localidade / Site",
        "Status do Bem",
        "Última Atividade"
    ]

    header_cols = [f"<th style='{th_style}'><span style=\"color:#ffffff\">{h}</span></th>" for h in headers]
    table_html.append(f"<tr>{''.join(header_cols)}</tr>")

    for idx, (_, row) in enumerate(df_report.iterrows()):
        item_num = idx + 1
        bg_style = "background-color: #d9e1f2;" if idx % 2 == 1 else ""
        
        pat = str(row.get("patrimonio", "")).strip()
        host = str(row.get("hostname", "")).strip()
        mod = str(row.get("modelo", "")).strip()
        ser = str(row.get("serial_number", "-")).strip()
        
        # Usuário formatado (Nome Completo + login entre parênteses)
        u_nome = str(row.get("usuario_nome", "")).strip()
        u_login = str(row.get("usuario_login", "")).strip()
        if u_nome and u_nome != "-" and u_login and u_login != "-" and u_login.lower() not in u_nome.lower():
            u_final = f"<strong>{u_nome}</strong> <span style='color:#555;'>({u_login})</span>"
        elif u_nome and u_nome != "-":
            u_final = f"<strong>{u_nome}</strong>"
        elif u_login and u_login != "-":
            u_final = f"<code>{u_login}</code>"
        else:
            u_final = "<span style='color:#888;'>Não identificado</span>"

        setor = str(row.get("setor", "-")).strip()
        loc = str(row.get("localidade", "-")).strip()
        status = str(row.get("status_bem", "-")).strip()
        ult = str(row.get("ultima_atividade", "-")).strip()

        row_cols = [
            f"<td style='border: 2px solid #cccccc; padding: 6px 8px; text-align: center; font-weight: bold; {bg_style}'>{item_num}</td>",
            f"<td style='border: 2px solid #cccccc; padding: 6px 8px; font-weight: bold; {bg_style}'>{pat}</td>",
            f"<td style='border: 2px solid #cccccc; padding: 6px 8px; {bg_style}'><code>{host}</code></td>",
            f"<td style='border: 2px solid #cccccc; padding: 6px 8px; {bg_style}'>{mod}</td>",
            f"<td style='border: 2px solid #cccccc; padding: 6px 8px; {bg_style}'>{ser}</td>",
            f"<td style='border: 2px solid #cccccc; padding: 6px 8px; {bg_style}'>{u_final}</td>",
            f"<td style='border: 2px solid #cccccc; padding: 6px 8px; {bg_style}'>{setor}</td>",
            f"<td style='border: 2px solid #cccccc; padding: 6px 8px; {bg_style}'>{loc}</td>",
            f"<td style='border: 2px solid #cccccc; padding: 6px 8px; font-weight: 600; {bg_style}'>{status}</td>",
            f"<td style='border: 2px solid #cccccc; padding: 6px 8px; {bg_style}'>{ult}</td>"
        ]
        table_html.append(f"<tr>{''.join(row_cols)}</tr>")

    # Rodapé de resumo
    footer_text = f"Total Levantado: {total_equip} {equip_word} | Relatório gerado em {now_str}"
    footer_html = f"<tr><td colspan='{len(headers)}' style='border: 2px solid #cccccc; padding: 8px 10px; background-color: #e9ecef; font-weight: bold; text-align: right;'>{footer_text}</td></tr>"
    table_html.append(footer_html)

    table_html.append("</table>")
    html_parts.append("".join(table_html))
    html_parts.append("<p style='margin-top: 15px; font-size: 11px; color: #666;'>Fonte dos dados: Active Directory (LDAP), Microsoft SCCM e Base Histórica da Bancada de TI.</p>")
    html_parts.append("</div>")

    return "\n".join(html_parts)


def generate_styled_dmp_excel(df_report: pd.DataFrame) -> io.BytesIO:
    """
    Gera um arquivo Excel (.xlsx) altamente estilizado e formatado para apresentação oficial ao DMP:
      - Cabeçalho corporativo com preenchimento azul marinho (#2F5597) e fonte branca em negrito;
      - Linhas zebradas com preenchimento suave (#F2F4F8);
      - Bordas finas completas em todas as células;
      - Auto-ajuste de largura de todas as colunas com margem para nenhum texto ficar cortado;
      - Alinhamento vertical e horizontal refinado;
      - Congelamento da primeira linha de cabeçalho para rolagem suave.
    """
    import openpyxl
    from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
    from openpyxl.utils import get_column_letter

    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Relatório DMP"
    ws.views.sheetView[0].showGridLines = True

    # Definição das colunas
    cols_map = [
        ("patrimonio", "Patrimônio"),
        ("hostname", "Hostname"),
        ("modelo", "Modelo"),
        ("serial_number", "Serial Number"),
        ("usuario_nome", "Nome Completo"),
        ("usuario_login", "Login"),
        ("setor", "Setor / Departamento"),
        ("localidade", "Localidade / Lotação"),
        ("status_bem", "Status do Bem"),
        ("ultima_atividade", "Última Atividade")
    ]

    headers = [col[1] for col in cols_map]
    ws.append(headers)

    # Estilos
    header_fill = PatternFill(start_color="2F5597", end_color="2F5597", fill_type="solid")
    header_font = Font(name="Calibri", size=11, bold=True, color="FFFFFF")
    zebra_fill = PatternFill(start_color="F2F5F9", end_color="F2F5F9", fill_type="solid")
    regular_font = Font(name="Calibri", size=10, color="000000")
    bold_font = Font(name="Calibri", size=10, bold=True, color="000000")

    thin_border_side = Side(border_style="thin", color="D9D9D9")
    cell_border = Border(
        left=thin_border_side,
        right=thin_border_side,
        top=thin_border_side,
        bottom=thin_border_side
    )

    header_border = Border(
        left=Side(border_style="thin", color="1B365D"),
        right=Side(border_style="thin", color="1B365D"),
        top=Side(border_style="medium", color="1B365D"),
        bottom=Side(border_style="medium", color="1B365D")
    )

    align_center = Alignment(horizontal="center", vertical="center", wrap_text=False)
    align_left = Alignment(horizontal="left", vertical="center", wrap_text=False)

    # Estiliza cabeçalho (Linha 1)
    ws.row_dimensions[1].height = 26
    for col_idx in range(1, len(headers) + 1):
        cell = ws.cell(row=1, column=col_idx)
        cell.fill = header_fill
        cell.font = header_font
        cell.border = header_border
        cell.alignment = align_center

    # Preenche linhas de dados
    for row_idx, (_, row_data) in enumerate(df_report.iterrows(), start=2):
        ws.row_dimensions[row_idx].height = 20
        is_even = (row_idx % 2 == 0)

        for col_idx, (col_key, col_title) in enumerate(cols_map, start=1):
            val = str(row_data.get(col_key, "-") or "-").strip()
            cell = ws.cell(row=row_idx, column=col_idx, value=val)
            
            # Fonte e Preenchimento
            if col_key == "patrimonio":
                cell.font = bold_font
            else:
                cell.font = regular_font

            if is_even:
                cell.fill = zebra_fill

            cell.border = cell_border

            # Alinhamento
            if col_key in ["patrimonio", "hostname", "serial_number", "usuario_login", "ultima_atividade"]:
                cell.alignment = align_center
            else:
                cell.alignment = align_left

    # Auto-ajuste de largura de cada coluna com folga de margem
    for col in ws.columns:
        col_letter = get_column_letter(col[0].column)
        max_len = 0
        for cell in col:
            val_str = str(cell.value or "")
            if len(val_str) > max_len:
                max_len = len(val_str)
        # Largura ajustada + margem de respiro
        adjusted_width = max(max_len + 4, 12)
        ws.column_dimensions[col_letter].width = adjusted_width

    # Congela painel na primeira linha para rolagem confortável
    ws.freeze_panes = "A2"

    out = io.BytesIO()
    wb.save(out)
    out.seek(0)
    return out


@st.dialog("📋 Relatório Unificado de Localização de Patrimônios (DMP)", width="large")
def modal_dmp_patrimonio_report():
    """
    Modal abrangente para consulta e geração do relatório de localização de patrimônios para o DMP.
    Permite entrada manual em lote (uma por linha ou separada por vírgula) ou upload de planilha Excel/CSV.
    """
    st.markdown("### 🏢 Dossiê e Localização de Bens para o DMP")
    st.caption("Consulte em lote a localização, último usuário, modelo, setor e situação de equipamentos no Active Directory, SCCM e histórico da Bancada.")

    tab_input, tab_upload = st.tabs(["✍️ Digitar / Colar Patrimônios", "📁 Importar de Arquivo (Excel / CSV)"])

    raw_input_text = ""
    with tab_input:
        st.write("Cole os números de patrimônio abaixo (como na mensagem do Teams/Excel, um por linha ou separados por vírgula):")
        exemplo_placeholder = "43188\n43299\n47675\n55477\n55522\n55589\n69288\n69305"
        raw_input_text = st.text_area(
            "Patrimônios:",
            value=st.session_state.get("dmp_patrimonios_text", ""),
            placeholder=exemplo_placeholder,
            height=160,
            key="dmp_input_textarea"
        )

    with tab_upload:
        st.write("Selecione uma planilha Excel (.xlsx, .xls) ou arquivo CSV contendo os patrimônios:")
        uploaded_file = st.file_uploader("Arquivo de Patrimônios:", type=["xlsx", "xls", "csv"], key="dmp_file_uploader")
        if uploaded_file is not None:
            try:
                if uploaded_file.name.endswith(".csv"):
                    df_up = pd.read_csv(uploaded_file, header=None)
                else:
                    df_up = pd.read_excel(uploaded_file, header=None)
                
                # Coleta todos os valores da primeira coluna ou busca coluna com 'patrimonio'
                extracted_pats = []
                for val in df_up.iloc[:, 0].dropna():
                    s_val = str(val).strip()
                    if s_val and not s_val.lower().startswith("patrim"):
                        extracted_pats.append(s_val)

                if extracted_pats:
                    raw_input_text = "\n".join(extracted_pats)
                    st.success(f"✅ {len(extracted_pats)} patrimônios identificados no arquivo!")
            except Exception as e_up:
                st.error(f"Erro ao ler arquivo: {e_up}")

    col_btn1, col_btn2 = st.columns([2, 1])
    with col_btn1:
        consultar = st.button("🔍 Gerar Relatório do DMP", type="primary", use_container_width=True)
    with col_btn2:
        if st.button("🗑️ Limpar", use_container_width=True):
            st.session_state["dmp_patrimonios_text"] = ""
            st.rerun()

    if consultar or st.session_state.get("dmp_generated", False):
        if raw_input_text.strip():
            st.session_state["dmp_patrimonios_text"] = raw_input_text
            st.session_state["dmp_generated"] = True
            
            pats_list = parse_patrimonios_input(raw_input_text)
            if not pats_list:
                st.warning("Nenhum patrimônio válido informado.")
                return

            with st.spinner(f"Consultando {len(pats_list)} equipamentos no Active Directory, SCCM e histórico..."):
                df_report, stats = query_dmp_patrimonio_report(pats_list)

            if df_report.empty:
                st.error("Não foi possível gerar os dados.")
                return

            st.markdown("---")

            # Resumo em métricas
            m_col1, m_col2, m_col3, m_col4, m_col5 = st.columns(5)
            m_col1.metric("Total Solicitado", stats["total"])
            m_col2.metric("🟢 Ativos na Rede", stats["ativos"])
            m_col3.metric("📦 Baixas / Doações", stats["baixados"])
            m_col4.metric("❓ Não Localizados", stats["nao_localizados"])
            m_col5.metric("Localizados", f"{stats['localizados']}/{stats['total']}")

            st.markdown("<br>", unsafe_allow_html=True)

            tab_grid, tab_html = st.tabs(["📊 Visualização em Tabela & Exportação", "📋 Formato de Texto / E-mail (HTML)"])

            with tab_grid:
                # Tabela Interativa
                cols_grid = [
                    "patrimonio",
                    "hostname",
                    "modelo",
                    "serial_number",
                    "usuario_nome",
                    "usuario_login",
                    "setor",
                    "localidade",
                    "status_bem",
                    "ultima_atividade"
                ]
                df_display = df_report[cols_grid].rename(columns={
                    "patrimonio": "Patrimônio",
                    "hostname": "Hostname",
                    "modelo": "Modelo",
                    "serial_number": "Serial Number",
                    "usuario_nome": "Nome Completo",
                    "usuario_login": "Login",
                    "setor": "Setor / Departamento",
                    "localidade": "Localidade / Lotação",
                    "status_bem": "Status do Bem",
                    "ultima_atividade": "Última Atividade"
                })

                st.dataframe(df_display, use_container_width=True, hide_index=True)

                st.markdown("<br>", unsafe_allow_html=True)
                exp_col1, exp_col2 = st.columns(2)
                with exp_col1:
                    # Exportação Excel Profissional Estilizada
                    try:
                        out_excel = generate_styled_dmp_excel(df_report)
                        st.download_button(
                            "📥 Baixar Planilha Excel (.xlsx)",
                            data=out_excel.getvalue(),
                            file_name=f"relatorio_patrimonio_dmp_{datetime.now().strftime('%Y%m%d_%H%M')}.xlsx",
                            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                            type="primary",
                            use_container_width=True
                        )
                    except Exception as e_exp:
                        st.caption(f"Aviso Excel: {e_exp}")

                with exp_col2:
                    csv_bytes = df_display.to_csv(index=False).encode("utf-8-sig")
                    st.download_button(
                        "📄 Baixar Arquivo CSV",
                        data=csv_bytes,
                        file_name=f"relatorio_patrimonio_dmp_{datetime.now().strftime('%Y%m%d_%H%M')}.csv",
                        mime="text/csv",
                        use_container_width=True
                    )

            with tab_html:
                html_result = build_dmp_report_html(df_report)
                
                st.write("💡 **Selecione o texto com o mouse, copie (Ctrl+C) e cole diretamente no chamado ou e-mail:**")
                st.markdown(
                    f'<div style="background-color: #ffffff; padding: 18px; border-radius: 6px; border: 1px solid #dddddd; max-height: 420px; overflow-y: auto;">{html_result}</div>',
                    unsafe_allow_html=True
                )

                st.markdown("<br>", unsafe_allow_html=True)
                with st.expander("💻 Ou copie o código-fonte HTML puro", expanded=False):
                    st.code(html_result, language="html")

        else:
            st.warning("⚠️ Informe ao menos um número de patrimônio para gerar o relatório.")
