import sys
import io
import pandas as pd
import streamlit as st
from datetime import datetime
from pathlib import Path

root_dir = Path(__file__).parent.parent.parent
sys.path.insert(0, str(root_dir))

from src.database import get_central_telefonica_df
from src.scrapers.oxe_scraper import check_oxe_sync_running, read_oxe_last_log_lines
from src.components.status_banner import render_log_expander
from src.components.pagination import (
    render_items_per_page_selector,
    paginate_items,
    render_pagination_controls
)
from src.components.metric_cards import render_metric_cards


@st.cache_data(ttl=300, show_spinner=False)
def _cached_get_central_telefonica_df() -> pd.DataFrame:
    return get_central_telefonica_df()


MISSING_EXTENSION_CATEGORY = "Não Cadastrado (Disponível)"
MISSING_EXTENSION_DISPLAY_NAME = "— Não cadastrado (Disponível) —"
MISSING_EXTENSION_STATUS = "Disponível para Criação"
MAX_CUSTOM_RANGE_SPAN = 10000

# Tabela oficial de ranges de numeração telefônica por unidade/prédio (Central OXE)
OFFICIAL_EXTENSION_RANGES: list[dict] = [
    {"unidade": "PGJ 1", "ddd_prefixo": "67 3312", "inicio": 4200, "fim": 4299, "obs": ""},
    {"unidade": "PGJ 1", "ddd_prefixo": "67 3312", "inicio": 4300, "fim": 4399, "obs": ""},
    {"unidade": "PGJ 1", "ddd_prefixo": "67 3312", "inicio": 5500, "fim": 5599, "obs": ""},
    {"unidade": "PGJ 2", "ddd_prefixo": "67 3318", "inicio": 2000, "fim": 2099, "obs": ""},
    {"unidade": "PGJ 2", "ddd_prefixo": "67 3318", "inicio": 2100, "fim": 2199, "obs": ""},
    {"unidade": "PGJ 2", "ddd_prefixo": "67 3318", "inicio": 3900, "fim": 3999, "obs": ""},
    {
        "unidade": "PGJ 3",
        "ddd_prefixo": "67 3318",
        "inicio": 8900,
        "fim": 8999,
        "obs": "Verificar na OXE (doc. original indica fim 8899)",
    },
    {"unidade": "PGJ 3", "ddd_prefixo": "67 2020", "inicio": 9300, "fim": 9399, "obs": ""},
    {"unidade": "PGJ 3", "ddd_prefixo": "67 2020", "inicio": 9400, "fim": 9499, "obs": ""},
    {"unidade": "PGJ 4", "ddd_prefixo": "67 2020", "inicio": 9000, "fim": 9099, "obs": ""},
    {"unidade": "PGJ 4", "ddd_prefixo": "67 2020", "inicio": 9100, "fim": 9199, "obs": ""},
    {"unidade": "PGJ 4", "ddd_prefixo": "67 2020", "inicio": 9200, "fim": 9299, "obs": ""},
    {"unidade": "PJCHA", "ddd_prefixo": "67 3357", "inicio": 2400, "fim": 2499, "obs": ""},
    {"unidade": "PJCHA", "ddd_prefixo": "67 3357", "inicio": 2500, "fim": 2599, "obs": ""},
    {"unidade": "PJCGR", "ddd_prefixo": "67 3316", "inicio": 2700, "fim": 2799, "obs": ""},
    {"unidade": "PJCGR", "ddd_prefixo": "67 3316", "inicio": 2800, "fim": 2899, "obs": ""},
    {"unidade": "PJCGR", "ddd_prefixo": "67 3316", "inicio": 2900, "fim": 2999, "obs": ""},
    {"unidade": "PJESP", "ddd_prefixo": "67 3316", "inicio": 4000, "fim": 4099, "obs": ""},
    {"unidade": "CORUMBÁ", "ddd_prefixo": "67 2020", "inicio": 8700, "fim": 8799, "obs": ""},
    {"unidade": "TRÊS LAGOAS", "ddd_prefixo": "67 2020", "inicio": 8800, "fim": 8899, "obs": ""},
]


def get_unit_for_extension(ramal_num: int) -> str:
    """Retorna a unidade oficial associada ao número de ramal informado, ou '-' se não mapeado."""
    for item in OFFICIAL_EXTENSION_RANGES:
        if item["inicio"] <= ramal_num <= item["fim"]:
            return item["unidade"]
    return "-"


def build_official_ranges_status_df(
    df: pd.DataFrame,
    ranges: list[dict] | None = None,
) -> pd.DataFrame:
    """
    Cruza a tabela oficial de ranges de ramais por unidade com os ramais cadastrados na OXE,
    calculando capacidade, cadastrados, disponíveis, ocupação (%), primeiro ramal livre e lacunas.
    """
    target_ranges = ranges if ranges is not None else OFFICIAL_EXTENSION_RANGES
    cols = [
        "Unidade",
        "Início",
        "Fim",
        "Faixa de Ramais",
        "Capacidade",
        "Cadastrados",
        "Disponíveis",
        "Ocupação (%)",
        "Primeiro Livre",
        "Números Disponíveis",
        "Observação",
        "_inicio_int",
        "_fim_int",
    ]
    if not target_ranges:
        return pd.DataFrame(columns=cols)

    existing_set: set[int] = set()
    if df is not None and not df.empty:
        ramal_col = _resolve_ramal_column(df)
        if ramal_col:
            for val in df[ramal_col].dropna().astype(str).str.strip().tolist():
                if val.isdigit():
                    existing_set.add(int(val))

    rows = []
    for item in target_ranges:
        inicio = int(item["inicio"])
        fim = int(item["fim"])
        if fim < inicio:
            inicio, fim = fim, inicio
        prefixo = str(item.get("ddd_prefixo", "")).strip()
        capacidade = fim - inicio + 1
        cadastrados = [n for n in range(inicio, fim + 1) if n in existing_set]
        disponiveis = [n for n in range(inicio, fim + 1) if n not in existing_set]
        ocupacao = round((len(cadastrados) / capacidade) * 100, 1) if capacidade > 0 else 0.0

        tel_inicio = f"{prefixo}-{inicio:04d}" if prefixo else f"{inicio:04d}"
        tel_fim = f"{prefixo}-{fim:04d}" if prefixo else f"{fim:04d}"

        rows.append({
            "Unidade": item["unidade"],
            "Início": tel_inicio,
            "Fim": tel_fim,
            "Faixa de Ramais": f"{inicio:04d}–{fim:04d}",
            "Capacidade": capacidade,
            "Cadastrados": len(cadastrados),
            "Disponíveis": len(disponiveis),
            "Ocupação (%)": ocupacao,
            "Primeiro Livre": f"{disponiveis[0]:04d}" if disponiveis else "Lotada",
            "Números Disponíveis": format_gap_ranges(disponiveis) if disponiveis else "Nenhum (faixa completa)",
            "Observação": item.get("obs", "") or "-",
            "_inicio_int": inicio,
            "_fim_int": fim,
        })

    return pd.DataFrame(rows, columns=cols)


def _resolve_ramal_column(df: pd.DataFrame) -> str | None:
    """Retorna o nome da coluna de ramal presente no DataFrame."""
    for candidate in ("ramal", "Ramal", "directory_number", "Directory_Number"):
        if candidate in df.columns:
            return candidate
    return None


def _numeric_sort_key(val) -> tuple[int, int, str]:
    """Chave de ordenação numérica natural para ramais (inteiros primeiro, depois alfanuméricos)."""
    text = str(val).strip()
    if text.isdigit():
        return (0, int(text), text)
    return (1, 0, text)


def format_gap_ranges(numbers: list[int]) -> str:
    """
    Agrupa uma lista de inteiros em intervalos legíveis compactos.
    Exemplo: [2010, 2011, 2012, 2013, 2014, 2020] -> '2010–2014, 2020'
    """
    if not numbers:
        return "-"
    ordered = sorted(set(numbers))
    ranges = []
    start = prev = ordered[0]
    for num in ordered[1:]:
        if num == prev + 1:
            prev = num
        else:
            ranges.append(f"{start}–{prev}" if start != prev else str(start))
            start = prev = num
    ranges.append(f"{start}–{prev}" if start != prev else str(start))
    return ", ".join(ranges)


def find_missing_extensions(
    df: pd.DataFrame,
    block_size: int = 100,
    custom_range: tuple[int, int] | None = None,
) -> pd.DataFrame:
    """
    Identifica números de ramais não cadastrados (lacunas na sequência numérica).

    - Modo automático (`custom_range=None`): agrupa os ramais numéricos existentes por blocos
      (padrão: centenas, ex: 2000–2099, 2200–2299) e retorna todos os números ausentes entre
      o menor e o maior ramal cadastrado de cada bloco, evitando lacunas artificiais entre
      faixas distintas.
    - Modo personalizado (`custom_range=(inicio, fim)`): retorna todos os números ausentes
      dentro do intervalo fechado [inicio, fim] (limitado a MAX_CUSTOM_RANGE_SPAN números).
    """
    cols = [
        "ramal",
        "nome_exibido",
        "nome_titular",
        "grupo_captura",
        "cat_rede_publica",
        "tipo_estacao",
        "subtipo",
        "ip_address",
        "mac_address",
        "endereco_equipamento",
        "categoria_dispositivo",
        "funcao_role",
        "login_externo",
        "email",
        "status_equipamento",
        "faixa",
        "ramal_anterior",
        "ramal_proximo",
        "is_missing_extension",
    ]
    if df is None or df.empty:
        if custom_range is None:
            return pd.DataFrame(columns=cols)
        existing_set: set[int] = set()
        sorted_existing: list[int] = []
        width_by_block: dict[int, int] = {}
    else:
        ramal_col = _resolve_ramal_column(df)
        if not ramal_col:
            return pd.DataFrame(columns=cols)

        raw_values = df[ramal_col].dropna().astype(str).str.strip().tolist()
        existing_set = set()
        width_by_block = {}
        for val in raw_values:
            if val.isdigit():
                num = int(val)
                existing_set.add(num)
                blk = num // block_size
                if blk not in width_by_block or len(val) > width_by_block[blk]:
                    width_by_block[blk] = len(val)
        sorted_existing = sorted(existing_set)

    def _fmt_num(num: int, min_width: int = 0) -> str:
        s = str(num)
        return s.zfill(min_width) if min_width > len(s) else s

    def _find_neighbors(num: int) -> tuple[str, str]:
        if not sorted_existing:
            return ("-", "-")
        import bisect
        idx = bisect.bisect_left(sorted_existing, num)
        prev_val = str(sorted_existing[idx - 1]) if idx > 0 else "-"
        next_val = str(sorted_existing[idx]) if idx < len(sorted_existing) else "-"
        return (prev_val, next_val)

    missing_rows = []

    if custom_range is not None:
        start_num, end_num = int(custom_range[0]), int(custom_range[1])
        if end_num < start_num:
            start_num, end_num = end_num, start_num
        if (end_num - start_num + 1) > MAX_CUSTOM_RANGE_SPAN:
            end_num = start_num + MAX_CUSTOM_RANGE_SPAN - 1

        faixa_label = f"{start_num}–{end_num}"
        min_w = max(len(str(start_num)), len(str(end_num)))
        for num in range(start_num, end_num + 1):
            if num not in existing_set:
                prev_r, next_r = _find_neighbors(num)
                missing_rows.append({
                    "ramal": _fmt_num(num, min_w),
                    "nome_exibido": MISSING_EXTENSION_DISPLAY_NAME,
                    "nome_titular": "-",
                    "grupo_captura": "-",
                    "cat_rede_publica": "-",
                    "tipo_estacao": "-",
                    "subtipo": "-",
                    "ip_address": "-",
                    "mac_address": "-",
                    "endereco_equipamento": "-",
                    "categoria_dispositivo": MISSING_EXTENSION_CATEGORY,
                    "funcao_role": "-",
                    "login_externo": "-",
                    "email": "-",
                    "status_equipamento": MISSING_EXTENSION_STATUS,
                    "faixa": faixa_label,
                    "ramal_anterior": prev_r,
                    "ramal_proximo": next_r,
                    "is_missing_extension": True,
                })
    else:
        if not sorted_existing:
            return pd.DataFrame(columns=cols)

        by_block: dict[int, list[int]] = {}
        for num in sorted_existing:
            by_block.setdefault(num // block_size, []).append(num)

        for blk in sorted(by_block):
            nums_in_blk = by_block[blk]
            if len(nums_in_blk) < 2:
                continue
            lo, hi = nums_in_blk[0], nums_in_blk[-1]
            if hi - lo + 1 == len(nums_in_blk):
                continue
            blk_start = blk * block_size
            blk_end = blk_start + block_size - 1
            min_w = width_by_block.get(blk, len(str(hi)))
            faixa_label = f"{_fmt_num(blk_start, min_w)}–{_fmt_num(blk_end, min_w)}"
            blk_set = set(nums_in_blk)

            for num in range(lo, hi + 1):
                if num not in blk_set:
                    prev_r, next_r = _find_neighbors(num)
                    missing_rows.append({
                        "ramal": _fmt_num(num, min_w),
                        "nome_exibido": MISSING_EXTENSION_DISPLAY_NAME,
                        "nome_titular": "-",
                        "grupo_captura": "-",
                        "cat_rede_publica": "-",
                        "tipo_estacao": "-",
                        "subtipo": "-",
                        "ip_address": "-",
                        "mac_address": "-",
                        "endereco_equipamento": "-",
                        "categoria_dispositivo": MISSING_EXTENSION_CATEGORY,
                        "funcao_role": "-",
                        "login_externo": "-",
                        "email": "-",
                        "status_equipamento": MISSING_EXTENSION_STATUS,
                        "faixa": faixa_label,
                        "ramal_anterior": prev_r,
                        "ramal_proximo": next_r,
                        "is_missing_extension": True,
                    })

    if not missing_rows:
        return pd.DataFrame(columns=cols)
    return pd.DataFrame(missing_rows, columns=cols)


def summarize_missing_by_range(
    df: pd.DataFrame,
    missing_df: pd.DataFrame,
    block_size: int = 100,
) -> pd.DataFrame:
    """
    Gera tabela resumo por faixa (centena) contendo quantidade de ramais cadastrados,
    quantidade de livres na sequência, primeiro ramal livre e lacunas agrupadas.
    """
    summary_cols = [
        "Faixa",
        "Unidade Oficial",
        "Intervalo Ativo",
        "Cadastrados",
        "Disponíveis (Sem Cadastro)",
        "Primeiro Livre",
        "Números Disponíveis (Lacunas)",
    ]
    if missing_df is None or missing_df.empty:
        return pd.DataFrame(columns=summary_cols)

    ramal_col = _resolve_ramal_column(df) if df is not None and not df.empty else None
    registered_by_block: dict[int, list[int]] = {}
    if ramal_col:
        for val in df[ramal_col].dropna().astype(str).str.strip().tolist():
            if val.isdigit():
                num = int(val)
                registered_by_block.setdefault(num // block_size, []).append(num)

    rows = []
    for faixa_label, grp in missing_df.groupby("faixa", sort=False):
        missing_nums = sorted(int(r) for r in grp["ramal"].astype(str) if r.isdigit())
        if not missing_nums:
            continue
        blk = missing_nums[0] // block_size
        reg_nums = sorted(set(registered_by_block.get(blk, [])))
        intervalo = f"{reg_nums[0]}–{reg_nums[-1]}" if reg_nums else "-"
        rows.append({
            "Faixa": faixa_label,
            "Unidade Oficial": get_unit_for_extension(missing_nums[0]),
            "Intervalo Ativo": intervalo,
            "Cadastrados": len(reg_nums),
            "Disponíveis (Sem Cadastro)": len(missing_nums),
            "Primeiro Livre": str(missing_nums[0]),
            "Números Disponíveis (Lacunas)": format_gap_ranges(missing_nums),
        })

    return pd.DataFrame(rows, columns=summary_cols)


def merge_registered_and_missing(
    df: pd.DataFrame,
    missing_df: pd.DataFrame,
) -> pd.DataFrame:
    """
    Combina ramais cadastrados e ramais não cadastrados mantendo ordenação numérica natural.
    """
    if df is None or df.empty:
        return missing_df.copy() if missing_df is not None else pd.DataFrame()

    base = df.copy()
    if "is_missing_extension" not in base.columns:
        base["is_missing_extension"] = False

    if missing_df is None or missing_df.empty:
        return base

    combined = pd.concat([base, missing_df], ignore_index=True, sort=False)
    combined["is_missing_extension"] = combined["is_missing_extension"].fillna(False).astype(bool)

    ramal_col = _resolve_ramal_column(combined)
    if ramal_col:
        sort_keys = combined[ramal_col].map(_numeric_sort_key)
        combined = combined.iloc[sort_keys.argsort(kind="mergesort")].reset_index(drop=True)

    return combined


def get_val(row, *keys, default="-"):
    """Busca insensível a maiúsculas/minúsculas e formato snake_case em dicionário/Series."""
    row_dict = {str(k).lower().replace(" ", "_").replace("/", "_").replace(".", "").replace("-", "_"): v for k, v in row.items()}
    for k in keys:
        norm_k = str(k).lower().replace(" ", "_").replace("/", "_").replace(".", "").replace("-", "_")
        if norm_k in row_dict:
            val = str(row_dict[norm_k]).strip()
            if val and val.lower() not in ["none", "nan", "null", "255"]:
                return val
    return default


def render_central_telefonica_page():
    """
    Renderiza a página principal de gestão e consulta da Central Telefônica (OXE).
    """
    st.markdown("""
        <style>
            .badge-ip {

                background-color: #065f46;
                color: #34d399;
                padding: 2px 8px;
                border-radius: 4px;
                font-weight: bold;
                font-size: 0.8rem;
            }
            .badge-analog {
                background-color: #7c2d12;
                color: #fb923c;
                padding: 2px 8px;
                border-radius: 4px;
                font-weight: bold;
                font-size: 0.8rem;
            }
        </style>
    """, unsafe_allow_html=True)

    st.title("📞 Central Telefônica (OXE)")
    st.caption("Consulta unificada de ramais, utilizadores, endereços IP e MAC Addresses de hardware.")

    oxe_ativo = check_oxe_sync_running()

    # Controle de término de sincronização para invalidação de cache
    was_oxe_syncing = st.session_state.get("was_oxe_syncing", False)
    if was_oxe_syncing and not oxe_ativo:
        st.cache_data.clear()
        st.session_state["was_oxe_syncing"] = False
        st.toast("🎉 Sincronização do OXE concluída com sucesso!", icon="📞")
        st.rerun()

    if oxe_ativo:
        st.session_state["was_oxe_syncing"] = True

    render_log_expander(
        "🤖 Robô do OXE Rodando em Segundo Plano – Acompanhar Progresso",
        oxe_ativo,
        read_oxe_last_log_lines,
        check_oxe_sync_running,
        "O robô está conectando à central Alcatel e pré-processando os dados neste momento. O painel permanece livre para uso!"
    )

    # Carrega dados do banco de dados SQLite / Tratados (com Cache)
    df = _cached_get_central_telefonica_df()

    # -----------------------------------------------------------------------------
    # FILTROS LATERAIS (SIDEBAR) & AÇÕES DE COLETA
    # -----------------------------------------------------------------------------
    with st.sidebar:
        from src.auth import is_admin
        if is_admin():
            st.markdown("## ⚙️ Ações e Coleta")
            if oxe_ativo:
                st.button("🤖 Sincronizando OXE...", width='stretch', disabled=True)
            else:
                if st.button("🔄 Sincronizar Ramais (OXE)", type="primary", width='stretch', help="Executa o scraper e o pré-processamento em segundo plano."):
                    import subprocess, time
                    popen_kwargs = {"creationflags": subprocess.CREATE_NO_WINDOW} if sys.platform == "win32" else {}
                    subprocess.Popen([sys.executable, "src/scrapers/oxe_scraper.py"], **popen_kwargs)
                    time.sleep(1.0)
                    st.toast("🚀 Scraper do OXE iniciado em segundo plano!", icon="🤖")
                    st.rerun()

            from src.config import OXE_URL
            oxe_web_url = OXE_URL if OXE_URL else ""
            st.link_button(
                "🌐 Acessar Central OXE ↗",
                url=oxe_web_url,
                width='stretch',
                help="Abre o painel de gerenciamento da Central Telefônica OXE em uma nova aba."
            )

            st.markdown("---")
        st.header("🔍 Filtros de Busca")

    total_ramais = len(df)
    df_missing_auto = find_missing_extensions(df)
    total_ramais_disponiveis = len(df_missing_auto)
    
    def is_valid_series(series):
        return ~series.astype(str).str.strip().str.lower().isin(["", "-", "none", "nan", "null", "255"])

    col_mac = 'mac_address' if 'mac_address' in df.columns else None
    col_ip = 'ip_address' if 'ip_address' in df.columns else ('endereco_ip' if 'endereco_ip' in df.columns else None)
    col_tipo = 'tipo_estacao' if 'tipo_estacao' in df.columns else ('tipo_de_estacao' if 'tipo_de_estacao' in df.columns else None)
    col_grupo = 'grupo_captura' if 'grupo_captura' in df.columns else ('pickup_group_name' if 'pickup_group_name' in df.columns else None)
    col_cat_pub = 'cat_rede_publica' if 'cat_rede_publica' in df.columns else ('public_network_category_id' if 'public_network_category_id' in df.columns else None)
    col_role = 'funcao_role' if 'funcao_role' in df.columns else ('set_role' if 'set_role' in df.columns else None)

    ramais_com_mac = len(df[is_valid_series(df[col_mac])]) if col_mac and col_mac in df.columns else 0
    ramais_com_ip = len(df[is_valid_series(df[col_ip])]) if col_ip and col_ip in df.columns else 0
    ramais_analogicos = len(df[df[col_tipo].astype(str).str.upper().str.contains('ANALOG', na=False)]) if col_tipo and col_tipo in df.columns else 0

    render_metric_cards([
        {
            "title": "TOTAL DE RAMAIS",
            "value": total_ramais,
            "border_color": "#3b82f6",
        },
        {
            "title": "TELEFONES IP COM MAC",
            "value": ramais_com_mac,
            "border_color": "#10b981",
        },
        {
            "title": "RAMAIS COM ENDEREÇO IP",
            "value": ramais_com_ip,
            "border_color": "#3b82f6",
        },
        {
            "title": "RAMAIS ANALÓGICOS / OUTROS",
            "value": ramais_analogicos,
            "border_color": "#f59e0b",
        },
        {
            "title": "LIVRES NA SEQUÊNCIA",
            "value": total_ramais_disponiveis,
            "border_color": "#8b5cf6",
        },
    ])

    st.markdown("---")

    # -----------------------------------------------------------------------------
    # FILTROS LATERAIS (SIDEBAR)
    # -----------------------------------------------------------------------------
    with st.sidebar:
        search_query = st.text_input(
            "Buscar por Ramal, Nome, IP, MAC ou Login:",
            value="",
            placeholder="Ex: 2153, Jean, 10.111..., 48:7A..."
        ).strip().lower()


        col_cat = 'categoria_dispositivo' if 'categoria_dispositivo' in df.columns else None
        categorias_disponiveis = sorted(df[col_cat].dropna().unique().tolist()) if col_cat else []
        
        sel_categorias = st.multiselect(
            "Categoria do Dispositivo:",
            options=categorias_disponiveis,
            default=[]
        )

        tipos_disponiveis = sorted(df[col_tipo].dropna().unique().tolist()) if col_tipo else []
        
        sel_tipos = st.multiselect(
            "Tipo de Estação (OXE):",
            options=tipos_disponiveis,
            default=[]
        )

        grupos_disponiveis = sorted(df[col_grupo].dropna().astype(str).unique().tolist()) if col_grupo else []
        sel_grupos = st.multiselect(
            "👥 Grupo de Captura:",
            options=[g for g in grupos_disponiveis if g.strip() and g.strip() not in ["-", "None", "nan"]],
            default=[]
        )

        cat_pub_disponiveis = sorted(df[col_cat_pub].dropna().astype(str).unique().tolist()) if col_cat_pub else []
        sel_cat_pub = st.multiselect(
            "🌐 Categoria Rede Pública:",
            options=[c for c in cat_pub_disponiveis if c.strip() and c.strip() not in ["-", "None", "nan"]],
            default=[]
        )

        roles_disponiveis = sorted(df[col_role].dropna().astype(str).unique().tolist()) if col_role else []
        sel_roles = st.multiselect(
            "🎭 Função / Role:",
            options=[r for r in roles_disponiveis if r.strip() and r.strip() not in ["-", "None", "nan"]],
            default=[]
        )

        only_with_mac = st.checkbox("Exibir apenas ramais com MAC Address", value=False)
        only_with_ip = st.checkbox("Exibir apenas ramais com Endereço IP", value=False)
        only_without_mac = st.checkbox("Exibir apenas ramais sem MAC Address", value=False)
        only_without_ip = st.checkbox("Exibir apenas ramais sem Endereço IP", value=False)

        st.markdown("---")
        st.markdown("**🆕 Números sem Ramal (Sequência)**")
        include_missing = st.checkbox(
            "Incluir números não cadastrados na lista",
            value=False,
            help="Insere os números livres (lacunas da sequência) diretamente na tabela na ordem numérica."
        )
        only_missing = st.checkbox(
            "Exibir apenas números não cadastrados",
            value=False,
            help="Mostra somente os números vagos dentro das sequências ativas, prontos para criação."
        )

        items_per_page = render_items_per_page_selector(
            key_prefix="central_oxe",
            options=[10, 25, 50, 100, 200, 500, "Todos"],
            default_index=2,
            label="📄 Ramais por página:"
        )


    # -----------------------------------------------------------------------------
    # APLICAÇÃO DOS FILTROS
    # -----------------------------------------------------------------------------
    if only_missing:
        df_base = df_missing_auto.copy()
    elif include_missing:
        df_base = merge_registered_and_missing(df, df_missing_auto)
    else:
        df_base = df.copy()

    df_filtered = df_base.copy()

    if search_query:
        cols_to_search = [c for c in df_filtered.columns if c not in ['id', 'data_atualizacao', 'is_missing_extension']]
        mask = pd.Series(False, index=df_filtered.index)
        for col in cols_to_search:
            mask |= df_filtered[col].astype(str).str.lower().str.contains(search_query, na=False)
        df_filtered = df_filtered[mask]

    if sel_categorias and col_cat and col_cat in df_filtered.columns:
        df_filtered = df_filtered[df_filtered[col_cat].isin(sel_categorias)]

    if sel_tipos and col_tipo and col_tipo in df_filtered.columns:
        df_filtered = df_filtered[df_filtered[col_tipo].isin(sel_tipos)]

    if sel_grupos and col_grupo and col_grupo in df_filtered.columns:
        df_filtered = df_filtered[df_filtered[col_grupo].astype(str).isin(sel_grupos)]

    if sel_cat_pub and col_cat_pub and col_cat_pub in df_filtered.columns:
        df_filtered = df_filtered[df_filtered[col_cat_pub].astype(str).isin(sel_cat_pub)]

    if sel_roles and col_role and col_role in df_filtered.columns:
        df_filtered = df_filtered[df_filtered[col_role].astype(str).isin(sel_roles)]

    if only_with_mac and col_mac and col_mac in df_filtered.columns:
        df_filtered = df_filtered[is_valid_series(df_filtered[col_mac])]

    if only_with_ip and col_ip and col_ip in df_filtered.columns:
        df_filtered = df_filtered[is_valid_series(df_filtered[col_ip])]

    if only_without_mac and col_mac and col_mac in df_filtered.columns:
        df_filtered = df_filtered[~is_valid_series(df_filtered[col_mac])]

    if only_without_ip and col_ip and col_ip in df_filtered.columns:
        df_filtered = df_filtered[~is_valid_series(df_filtered[col_ip])]

    @st.dialog("📞 Detalhes Completos do Ramal", width="large")
    def show_ramal_details(row):
        ramal_num = get_val(row, "ramal", "directory_number")
        is_missing = bool(row.get("is_missing_extension", False)) if hasattr(row, "get") else False

        if is_missing:
            st.markdown(f"### 🆕 Ramal **{ramal_num}** — Não Cadastrado (Disponível)")
            st.success(
                f"O número **{ramal_num}** não está cadastrado na Central Telefônica OXE e está livre na sequência para criação de um novo ramal."
            )
            col1, col2, col3 = st.columns(3)
            with col1:
                st.metric("Faixa (Centena)", get_val(row, "faixa"))
            with col2:
                st.metric("Ramal Anterior Cadastrado", get_val(row, "ramal_anterior"))
            with col3:
                st.metric("Próximo Ramal Cadastrado", get_val(row, "ramal_proximo"))

            from src.config import OXE_URL
            if OXE_URL:
                st.link_button(
                    "🌐 Abrir Central OXE para Criar Ramal ↗",
                    url=OXE_URL,
                    width='stretch',
                )
            if st.button("Fechar", key=f"close_ramal_modal_{ramal_num}"):
                st.rerun()
            return

        nome = get_val(row, "nome_exibido", "nome_titular", "annu_name", "display_name")
        
        st.markdown(f"### 📞 Ramal **{ramal_num}** — {nome}")
        st.caption("Ficha técnica completa extraída da Central Telefônica Alcatel-Lucent OmniPCX Enterprise (OXE).")

        # 1. Identificação & Titular
        with st.expander("👤 Identificação do Assinante & Titular", expanded=True):
            col1, col2 = st.columns(2)
            with col1:
                st.markdown(f"**Ramal:** `{ramal_num}`")
                st.markdown(f"**Nome Titular:** {get_val(row, 'nome_titular', 'annu_name')}")
                st.markdown(f"**Complemento:** {get_val(row, 'complemento', 'annu_first_name', 'utf8_phone_book_first_name')}")
                st.markdown(f"**Nome Exibido:** {get_val(row, 'nome_exibido', 'phone_book_name', 'display_name')}")
            with col2:
                st.markdown(f"**Login Externo:** `{get_val(row, 'login_externo', 'external_login')}`")
                st.markdown(f"**E-mail:** {get_val(row, 'email', 'mail_address')}")
                st.markdown(f"**Idioma (Language_Id):** `{get_val(row, 'language_id')}`")
                st.markdown(f"**Qtd. Usuários:** `{get_val(row, 'number_of_users')}`")

        # 2. Rede, IP & Hardware
        with st.expander("🌐 Rede, IP & Hardware", expanded=True):
            col1, col2 = st.columns(2)
            with col1:
                st.markdown(f"**Endereço IP:** `{get_val(row, 'endereco_ip', 'ip_address')}`")
                st.markdown(f"**MAC Address:** `{get_val(row, 'mac_address', 'ethernet_address')}`")
                st.markdown(f"**Categoria Dispositivo:** `{get_val(row, 'categoria_dispositivo')}`")
            with col2:
                st.markdown(f"**Endereço PABX:** `{get_val(row, 'endereco_equipamento')}` (Rack `{get_val(row, 'rack', 'equipment_address_rack')}` / Placa `{get_val(row, 'placa', 'equipment_address_board')}` / Terminal `{get_val(row, 'terminal', 'equipment_address_terminal')}`)")
                st.markdown(f"**Criptografia Nativa:** `{get_val(row, 'nativeencryption')}`")
                st.markdown(f"**Suporte a Vídeo:** `{get_val(row, 'videosupportprofile')}`")

        # 3. Telefonia, Grupo de Captura & Regras de Acesso
        with st.expander("📞 Telefonia, Grupo de Captura & Regras de Acesso", expanded=True):
            col1, col2 = st.columns(2)
            with col1:
                st.markdown(f"**Grupo de Captura (Pickup Group):** `{get_val(row, 'grupo_captura', 'pickup_group_name')}`")
                st.markdown(f"**Cat. Rede Pública (Public Network):** `{get_val(row, 'cat_rede_publica', 'public_network_category_id')}`")
                st.markdown(f"**Cat. Encaminhamento Externo:** `{get_val(row, 'external_forwarding_category_id')}`")
                st.markdown(f"**Cat. Medição / Tarifação:** `{get_val(row, 'metering_category')}`")
            with col2:
                st.markdown(f"**Centro de Custo:** `{get_val(row, 'centro_de_custo', 'cost_center_name', 'cost_center_id')}`")
                st.markdown(f"**Função Adm:** `{get_val(row, 'funcao_adm', 'function')}`")
                st.markdown(f"**Cat. Chamador (Caller Category):** `{get_val(row, 'caller_category')}`")
                st.markdown(f"**Cat. Conexão / Dados:** `{get_val(row, 'connection_category_id')}` / `{get_val(row, 'data_connection_category_id')}`")

        # 4. Tipo de Estação, Perfil & Teclado
        with st.expander("⚙️ Tipo de Estação, Perfil & Teclado", expanded=False):
            col1, col2 = st.columns(2)
            with col1:
                st.markdown(f"**Tipo de Estação:** `{get_val(row, 'tipo_de_estacao', 'tipo_estacao', 'station_type')}`")
                st.markdown(f"**Subtipo de Estação:** `{get_val(row, 'subtipo', 'station_sub_type')}`")
                st.markdown(f"**Função / Role:** `{get_val(row, 'funcao_role', 'set_role')}`")
                st.markdown(f"**Perfil DM:** `{get_val(row, 'dm_profile')}` | **Tipo:** `{get_val(row, 'profile_type')}`")
            with col2:
                st.markdown(f"**Módulos de Extensão:** `{get_val(row, 'add_on_module_1')}` / `{get_val(row, 'add_on_module_2')}`")
                st.markdown(f"**Teclado Interno:** `{get_val(row, 'internal_keyboard')}`")
                st.markdown(f"**Entidade / Domínio:** `{get_val(row, 'entity_number')}` / `{get_val(row, 'domain_identifier')}`")
                st.markdown(f"**Licença Opex:** `{get_val(row, 'opex_license')}`")

        # 5. Correio de Voz & Recursos Especiais
        with st.expander("🎙️ Correio de Voz & Recursos Especiais", expanded=False):
            col1, col2 = st.columns(2)
            with col1:
                st.markdown(f"**Tipo Correio de Voz:** `{get_val(row, 'voice_mail_type')}`")
                st.markdown(f"**Ramal Correio de Voz:** `{get_val(row, 'voice_mail_directory_number')}`")
                st.markdown(f"**Grupo ACD:** `{get_val(row, 'acd_group_directory_number')}` ({get_val(row, 'acd_type')})")
            with col2:
                st.markdown(f"**Controle PIN:** `{get_val(row, 'pin_group_control')}` (Grupo `{get_val(row, 'user_pin_group')}`)")
                st.markdown(f"**Atendimento VIP:** `{get_val(row, 'vip')}`")
                st.markdown(f"**Chamada Urgente:** `{get_val(row, 'urgent_call')}`")

        # 6. Todos os Dados Brutos da API
        with st.expander("🔍 Dicionário Completo de Dados Brutos (API OXE)", expanded=False):
            raw_items = row.to_dict() if hasattr(row, "to_dict") else dict(row)
            clean_dict = {k: v for k, v in raw_items.items() if str(v).strip() and str(v).strip().lower() not in ["none", "nan", "null"]}
            st.json(clean_dict)

        if st.button("Fechar", key=f"close_ramal_modal_{ramal_num}"):
            st.rerun()

    # -----------------------------------------------------------------------------
    # ABAS PRINCIPAIS: LISTA DE RAMAIS, NÚMEROS DISPONÍVEIS & RANGES OFICIAIS
    # -----------------------------------------------------------------------------
    df_official_ranges = build_official_ranges_status_df(df)
    tab_lista, tab_disponiveis, tab_ranges = st.tabs([
        f"📋 Lista de Ramais ({len(df_filtered)})",
        f"🆕 Números Disponíveis / Sem Cadastro ({total_ramais_disponiveis})",
        f"📐 Ranges Oficiais por Unidade ({len(df_official_ranges)})",
    ])

    with tab_lista:
        st.subheader(f"📋 Lista de Ramais ({len(df_filtered)} de {len(df_base)})")
        st.caption("💡 Clique na **caixa de seleção (checkbox)** no início de qualquer linha para abrir a Ficha Técnica Completa do Ramal.")

        # Mapeamento de colunas amigáveis para exibição
        column_rename_map = {
            "ramal": "Ramal",
            "nome_titular": "Nome / Titular",
            "complemento": "Complemento",
            "nome_exibido": "Nome Exibido",
            "tipo_estacao": "Tipo de Estação",
            "subtipo": "Subtipo",
            "funcao_role": "Função / Role",
            "grupo_captura": "Grupo de Captura",
            "cat_rede_publica": "Cat. Rede Pública",
            "login_externo": "Login Externo",
            "email": "E-mail",
            "rack": "Rack",
            "placa": "Placa",
            "terminal": "Terminal",
            "endereco_equipamento": "Endereço Equipamento",
            "ip_address": "Endereço IP",
            "mac_address": "MAC Address",
            "categoria_dispositivo": "Categoria",
            "status_equipamento": "Status",
            "data_atualizacao": "Última Atualização"
        }

        df_display = df_filtered.rename(columns=column_rename_map)

        cols_order = [
            "Ramal", "Nome Exibido", "Grupo de Captura", "Cat. Rede Pública",
            "Tipo de Estação", "Subtipo", "Endereço IP", "MAC Address",
            "Endereço Equipamento", "Categoria", "Função / Role", "Login Externo", "E-mail"
        ]
        existing_cols = [c for c in cols_order if c in df_display.columns]

        # Paginação
        df_page, current_page, total_pages, total_items = paginate_items(
            df_display[existing_cols],
            page_key="central_oxe",
            items_per_page=items_per_page
        )

        if "last_selected_ramal" not in st.session_state:
            st.session_state["last_selected_ramal"] = None

        selection_event = st.dataframe(
            df_page,
            width='stretch',
            hide_index=True,
            on_select="rerun",
            selection_mode="single-row",
            key="tabela_oxe_datagrid"
        )

        selected_rows = selection_event.selection.rows if hasattr(selection_event, "selection") else []

        if selected_rows:
            current_selected = selected_rows[0]
            if st.session_state["last_selected_ramal"] != current_selected:
                st.session_state["last_selected_ramal"] = current_selected
                page_offset = 0 if items_per_page == "Todos" else (current_page - 1) * int(items_per_page)
                row_data = df_filtered.iloc[page_offset + current_selected]
                show_ramal_details(row_data)
        else:
            st.session_state["last_selected_ramal"] = None

        # Régua de controles de paginação
        render_pagination_controls(
            page_key="central_oxe",
            current_page=current_page,
            total_pages=total_pages,
            total_items=total_items,
            items_per_page=items_per_page
        )

        # -----------------------------------------------------------------------------
        # BOTÃO DE EXPORTAÇÃO
        # -----------------------------------------------------------------------------
        st.markdown("---")
        output = io.BytesIO()
        with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
            df_display[existing_cols].to_excel(writer, sheet_name="Central Telefônica", index=False)
        
        st.download_button(
            label="📥 Baixar Tabela Filtrada em Excel (.xlsx)",
            data=output.getvalue(),
            file_name=f"Central_Telefonica_Filtrada_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            width='stretch'
        )

    with tab_disponiveis:
        st.subheader("🆕 Números Disponíveis (Não Cadastrados na Sequência)")
        st.caption(
            "Identifica automaticamente os números ausentes (lacunas) dentro de cada faixa de centena ativa na Central OXE, "
            "ou permite consultar qualquer intervalo personalizado para encontrar rapidamente um ramal livre."
        )

        modo_consulta = st.radio(
            "Modo de busca de números livres:",
            options=["Lacunas automáticas por faixa (centena)", "Consultar faixa personalizada (início e fim)"],
            horizontal=True,
            key="oxe_missing_mode",
        )

        if modo_consulta == "Lacunas automáticas por faixa (centena)":
            df_summary = summarize_missing_by_range(df, df_missing_auto)

            if df_summary.empty:
                st.info("Nenhuma lacuna numérica encontrada nas sequências ativas de ramais.")
            else:
                with st.expander(f"📊 Resumo de Lacunas por Faixa ({len(df_summary)} faixas com números livres)", expanded=True):
                    st.dataframe(df_summary, width='stretch', hide_index=True)

                faixas_opcoes = df_summary["Faixa"].tolist()
                col_f1, col_f2 = st.columns([2, 1])
                with col_f1:
                    sel_faixas = st.multiselect(
                        "Filtrar por Faixa (Centena):",
                        options=faixas_opcoes,
                        default=[],
                        placeholder="Todas as faixas com números disponíveis...",
                        key="oxe_missing_faixas",
                    )
                with col_f2:
                    busca_livre = st.text_input(
                        "Buscar número específico:",
                        value="",
                        placeholder="Ex: 2007, 24...",
                        key="oxe_missing_search",
                    ).strip()

                df_livres_view = df_missing_auto.copy()
                if sel_faixas:
                    df_livres_view = df_livres_view[df_livres_view["faixa"].isin(sel_faixas)]
                if busca_livre:
                    df_livres_view = df_livres_view[
                        df_livres_view["ramal"].astype(str).str.contains(busca_livre, case=False, na=False)
                        | df_livres_view["faixa"].astype(str).str.contains(busca_livre, case=False, na=False)
                    ]
        else:
            col_r1, col_r2 = st.columns(2)
            with col_r1:
                range_start = st.number_input(
                    "Ramal inicial:",
                    min_value=0,
                    max_value=999999,
                    value=2000,
                    step=1,
                    key="oxe_custom_range_start",
                )
            with col_r2:
                range_end = st.number_input(
                    "Ramal final:",
                    min_value=0,
                    max_value=999999,
                    value=2099,
                    step=1,
                    key="oxe_custom_range_end",
                )

            if range_end < range_start:
                st.warning("O ramal final é menor que o inicial; os limites foram ajustados automaticamente.")
            elif (range_end - range_start + 1) > MAX_CUSTOM_RANGE_SPAN:
                st.warning(f"Intervalo limitado aos primeiros {MAX_CUSTOM_RANGE_SPAN:,} números para manter a resposta instantânea.")

            df_livres_view = find_missing_extensions(df, custom_range=(int(range_start), int(range_end)))
            if not df_livres_view.empty:
                livres_ints = [int(r) for r in df_livres_view["ramal"].astype(str) if r.isdigit()]
                st.info(
                    f"**{len(df_livres_view)}** número(s) não cadastrado(s) entre **{min(range_start, range_end)}** e **{max(range_start, range_end)}**: "
                    f"`{format_gap_ranges(livres_ints)}`"
                )
            else:
                st.success(
                    f"Todos os números entre **{min(range_start, range_end)}** e **{max(range_start, range_end)}** já estão cadastrados na Central OXE."
                )

        if not df_livres_view.empty:
            df_livres_display = df_livres_view[[
                "ramal", "faixa", "ramal_anterior", "ramal_proximo", "status_equipamento"
            ]].rename(columns={
                "ramal": "Número Disponível",
                "faixa": "Faixa",
                "ramal_anterior": "Ramal Anterior Cadastrado",
                "ramal_proximo": "Próximo Ramal Cadastrado",
                "status_equipamento": "Status",
            })

            st.markdown(f"**📋 Números disponíveis encontrados: {len(df_livres_display)}**")
            df_livres_page, cur_p_m, tot_p_m, tot_i_m = paginate_items(
                df_livres_display,
                page_key="central_oxe_missing",
                items_per_page=items_per_page,
            )
            st.dataframe(df_livres_page, width='stretch', hide_index=True, key="tabela_oxe_missing_datagrid")
            render_pagination_controls(
                page_key="central_oxe_missing",
                current_page=cur_p_m,
                total_pages=tot_p_m,
                total_items=tot_i_m,
                items_per_page=items_per_page,
            )

            st.markdown("---")
            out_missing = io.BytesIO()
            with pd.ExcelWriter(out_missing, engine='xlsxwriter') as writer:
                df_livres_display.to_excel(writer, sheet_name="Ramais Disponiveis", index=False)

            st.download_button(
                label="📥 Baixar Números Disponíveis em Excel (.xlsx)",
                data=out_missing.getvalue(),
                file_name=f"Central_Telefonica_Ramais_Disponiveis_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                width='stretch',
                key="download_oxe_missing_xlsx",
            )

    with tab_ranges:
        st.subheader("📐 Ranges Oficiais de Telefones para Cadastro (OXE)")
        st.caption(
            "Faixas oficiais de numeração telefônica por unidade/prédio cruzadas em tempo real com os ramais "
            "cadastrados na Central OXE para facilitar o cadastro correto de novos ramais."
        )
        st.warning(
            "⚠️ **Atenção:** Cadastrar um ramal fora da faixa oficial correspondente à unidade pode impedir "
            "o funcionamento correto das ligações. Consulte a tabela abaixo antes de criar um novo ramal na OXE."
        )

        unidades_oficiais = list(dict.fromkeys(df_official_ranges["Unidade"].tolist())) if not df_official_ranges.empty else []
        col_u1, col_u2 = st.columns([2, 1])
        with col_u1:
            sel_unidades_range = st.multiselect(
                "Filtrar por Unidade / Prédio:",
                options=unidades_oficiais,
                default=[],
                placeholder="Todas as unidades oficiais...",
                key="oxe_ranges_unidades",
            )
        with col_u2:
            only_ranges_with_free = st.checkbox(
                "Apenas faixas com números disponíveis",
                value=False,
                key="oxe_ranges_only_free",
            )

        df_ranges_view = df_official_ranges.copy()
        if sel_unidades_range:
            df_ranges_view = df_ranges_view[df_ranges_view["Unidade"].isin(sel_unidades_range)]
        if only_ranges_with_free:
            df_ranges_view = df_ranges_view[df_ranges_view["Disponíveis"] > 0]

        ranges_display_cols = [
            "Unidade",
            "Início",
            "Fim",
            "Faixa de Ramais",
            "Cadastrados",
            "Disponíveis",
            "Ocupação (%)",
            "Primeiro Livre",
            "Números Disponíveis",
            "Observação",
        ]
        df_ranges_display = df_ranges_view[ranges_display_cols]

        if df_ranges_display.empty:
            st.info("Nenhuma faixa oficial encontrada para os filtros selecionados.")
        else:
            total_cap = int(df_ranges_view["Capacidade"].sum())
            total_cad = int(df_ranges_view["Cadastrados"].sum())
            total_liv = int(df_ranges_view["Disponíveis"].sum())
            c_m1, c_m2, c_m3 = st.columns(3)
            c_m1.metric("Capacidade das Faixas Filtradas", total_cap)
            c_m2.metric("Ramais Cadastrados", total_cad)
            c_m3.metric("Números Disponíveis para Cadastro", total_liv)

            st.dataframe(
                df_ranges_display,
                width='stretch',
                hide_index=True,
                key="tabela_oxe_ranges_oficiais",
                column_config={
                    "Ocupação (%)": st.column_config.ProgressColumn(
                        "Ocupação (%)",
                        min_value=0.0,
                        max_value=100.0,
                        format="%.1f%%",
                    ),
                },
            )

            # Detalhamento rápido dos números livres de uma faixa específica
            faixas_com_livres = df_ranges_view[df_ranges_view["Disponíveis"] > 0]
            if not faixas_com_livres.empty:
                opcoes_faixa = [
                    f"{r['Unidade']} — {r['Início']} a {r['Fim']} ({r['Disponíveis']} livre(s))"
                    for _, r in faixas_com_livres.iterrows()
                ]
                with st.expander("🔍 Ver todos os números livres de uma faixa oficial específica", expanded=False):
                    escolha_idx = st.selectbox(
                        "Selecione a faixa oficial:",
                        options=list(range(len(opcoes_faixa))),
                        format_func=lambda i: opcoes_faixa[i],
                        key="oxe_ranges_detail_select",
                    )
                    chosen_row = faixas_com_livres.iloc[escolha_idx]
                    df_range_free = find_missing_extensions(
                        df,
                        custom_range=(int(chosen_row["_inicio_int"]), int(chosen_row["_fim_int"])),
                    )
                    st.caption(
                        f"**{chosen_row['Unidade']}** (`{chosen_row['Início']}` a `{chosen_row['Fim']}`): "
                        f"Primeiro livre sugerido: **`{chosen_row['Primeiro Livre']}`**"
                    )
                    if not df_range_free.empty:
                        st.dataframe(
                            df_range_free[["ramal", "ramal_anterior", "ramal_proximo", "status_equipamento"]].rename(columns={
                                "ramal": "Ramal Livre",
                                "ramal_anterior": "Ramal Anterior Cadastrado",
                                "ramal_proximo": "Próximo Ramal Cadastrado",
                                "status_equipamento": "Status",
                            }),
                            width='stretch',
                            hide_index=True,
                            key="tabela_oxe_range_detail_free",
                        )

            st.markdown("---")
            out_ranges = io.BytesIO()
            with pd.ExcelWriter(out_ranges, engine='xlsxwriter') as writer:
                df_ranges_display.to_excel(writer, sheet_name="Ranges Oficiais OXE", index=False)

            st.download_button(
                label="📥 Baixar Tabela de Ranges Oficiais em Excel (.xlsx)",
                data=out_ranges.getvalue(),
                file_name=f"Central_Telefonica_Ranges_Oficiais_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                width='stretch',
                key="download_oxe_ranges_xlsx",
            )


