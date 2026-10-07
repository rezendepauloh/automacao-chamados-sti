import pandas as pd
from datetime import datetime
from .connection import get_connection

def setup_ramais_table():
    """Cria a tabela ramais_mpms se não existir."""
    conn = get_connection()
    cursor = conn.cursor()
    cursor.execute("""
    CREATE TABLE IF NOT EXISTS ramais_mpms (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        localidade TEXT,
        setor_nome TEXT,
        telefone_ramal TEXT,
        tipo TEXT,
        data_atualizacao TEXT
    )
    """)
    conn.commit()
    conn.close()

def save_ramais_to_db(df: pd.DataFrame):
    """Limpa a tabela ramais_mpms e insere os dados do DataFrame recebido."""
    setup_ramais_table()
    conn = get_connection()
    cursor = conn.cursor()
    cursor.execute("DELETE FROM ramais_mpms")
    conn.commit()
    
    if not df.empty:
        df_to_save = df.copy()
        if "data_atualizacao" not in df_to_save.columns:
            df_to_save["data_atualizacao"] = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
            
        cols = ["localidade", "setor_nome", "telefone_ramal", "tipo", "data_atualizacao"]
        cols_present = [c for c in cols if c in df_to_save.columns]
        df_to_save[cols_present].to_sql("ramais_mpms", conn, if_exists="append", index=False)
        conn.commit()
    conn.close()

SIGLAS_INSTITUCIONAIS = {
    'MPMS', 'PJ', 'PGJ', 'GAECO', 'CAO', 'CAOMA', 'STI', 'STIC', 'DTI', 'SECOM', 
    'NAEP', 'GACEP', 'CEAF', 'CGMP', 'CPJ', 'CSMP', 'EIA', 'RIMA', 'TAC',
    'AD', 'SCCM', 'PXE', 'TI', 'RH', 'I', 'II', 'III', 'IV', 'V', 'VI', 'VII', 'VIII', 'IX', 'X'
}
PREPOSICOES_PORTUGUES = {'de', 'da', 'do', 'das', 'dos', 'em', 'no', 'na', 'nos', 'nas', 'a', 'o', 'e', 'ao', 'aos', 'com', 'por', 'para'}

def clean_dots_and_page_numbers(text: str) -> str:
    """Remove sequências de pontilhados de sumário e números de página (ex: '... 33')."""
    import re
    t = re.sub(r'\.{2,}\s*\d*', '', str(text))
    return re.sub(r'\s+', ' ', t).strip()

def smart_title(text: str) -> str:
    """Converte texto para Title Case preservando siglas conhecidas e preposições."""
    import re
    if not text:
        return ""
    t = clean_dots_and_page_numbers(text)
    words = t.split(" ")
    result = []
    for i, w in enumerate(words):
        w_clean = re.sub(r'^[^\w]+|[^\w]+$', '', w)
        w_upper = w_clean.upper()
        if w_upper in SIGLAS_INSTITUCIONAIS:
            result.append(w.replace(w_clean, w_upper))
        elif w_clean.lower() in PREPOSICOES_PORTUGUES and 0 < i < len(words) - 1:
            result.append(w.replace(w_clean, w_clean.lower()))
        else:
            result.append(w.replace(w_clean, w_clean.capitalize()))
    return " ".join(result)

def format_ramal_num(num_str: str) -> str:
    """Formata sequências de números de ramal separados por espaço usando separador visual elegante."""
    import re
    s = re.sub(r'\s+', ' ', str(num_str)).strip()
    tokens = s.split(' ')
    if len(tokens) > 1 and all(len(t) == 4 and t.isdigit() for t in tokens):
        return ' • '.join(tokens)
    return s

def clean_ramais_dataframe(df_raw: pd.DataFrame) -> pd.DataFrame:
    """Higieniza os dados de ramais, propagando a localidade correta e limpando cabeçalhos."""
    import re
    if df_raw.empty:
        return df_raw
    df = df_raw.copy()
    cleaned_rows = []
    last_valid_loc = "MPMS - Geral"
    header_regex = re.compile(r'PJ MEMBRO|GABINETE ASSESSORIA|GABIN ASSESSORIA|Chefe do Departamento|REMOTA|RECEPÇÃO', re.IGNORECASE)

    for _, row in df.iterrows():
        loc_raw = str(row.get('localidade', '')).strip()
        setor_raw = str(row.get('setor_nome', '')).strip()
        ramal_raw = str(row.get('telefone_ramal', '')).strip()

        is_loc_header = bool(header_regex.search(loc_raw))
        if not is_loc_header and not re.search(r'^\.{3,}', loc_raw) and loc_raw not in ['Geral', 'INTERIOR DO ESTADO']:
            last_valid_loc = loc_raw
        elif loc_raw in ['INTERIOR DO ESTADO']:
            last_valid_loc = loc_raw

        final_loc = last_valid_loc if is_loc_header else loc_raw

        # Normaliza setor_nome se for repetição da localidade ou cabeçalho residual
        if header_regex.search(setor_raw):
            final_setor = "Recepção / Apoio Administrativo"
        elif setor_raw == loc_raw or setor_raw == final_loc:
            final_setor = "Atendimento Geral / Recepção"
        else:
            final_setor = setor_raw

        final_loc = smart_title(final_loc)
        final_setor = smart_title(final_setor)
        final_ramal = format_ramal_num(ramal_raw)

        r_dict = row.to_dict()
        r_dict['localidade'] = final_loc
        r_dict['setor_nome'] = final_setor
        r_dict['telefone_ramal'] = final_ramal
        cleaned_rows.append(r_dict)

    return pd.DataFrame(cleaned_rows)

def get_ramais_df(clean: bool = True) -> pd.DataFrame:
    """Retorna os dados da tabela ramais_mpms em um DataFrame, com higienização inteligente por padrão."""
    setup_ramais_table()
    conn = get_connection()
    try:
        df = pd.read_sql_query("SELECT id, localidade, setor_nome, telefone_ramal, tipo, data_atualizacao FROM ramais_mpms", conn)
    except Exception:
        df = pd.DataFrame(columns=["id", "localidade", "setor_nome", "telefone_ramal", "tipo", "data_atualizacao"])
    conn.close()
    
    if clean and not df.empty:
        return clean_ramais_dataframe(df)
    return df


def setup_ramais_config_table():
    """Cria a tabela ramais_config se não existir."""
    conn = get_connection()
    cursor = conn.cursor()
    cursor.execute("""
    CREATE TABLE IF NOT EXISTS ramais_config (
        chave TEXT PRIMARY KEY,
        valor TEXT
    )
    """)
    conn.commit()
    conn.close()


def get_ramais_config() -> dict:
    """Retorna os links configurados para os PDFs de ramais."""
    setup_ramais_config_table()
    conn = get_connection()
    cursor = conn.cursor()
    cursor.execute("SELECT chave, valor FROM ramais_config")
    rows = cursor.fetchall()
    conn.close()
    config = {k: v for k, v in rows}
    return {
        "Interior": config.get("url_interior", "https://www.mpms.mp.br/anexo/MTMzMDYxNDI3NTAwODYzMjkwNDNmYmI5MGYwYjU2ZGE5ZWI5M2ZmN2EwMTQxLTA0MQ"),
        "Capital / PGJ": config.get("url_capital", "https://www.mpms.mp.br/anexo/MTMzMDYxNDE2ODMwOGI3MjcxZWQ2YzhkYjYyODkwOGFlMDRjNTUzYWFmY2ZhLTA0MQ")
    }


def save_ramais_config(url_interior: str, url_capital: str):
    """Salva os links atualizados dos PDFs de ramais."""
    setup_ramais_config_table()
    conn = get_connection()
    cursor = conn.cursor()
    cursor.execute("INSERT OR REPLACE INTO ramais_config (chave, valor) VALUES ('url_interior', ?)", (url_interior.strip(),))
    cursor.execute("INSERT OR REPLACE INTO ramais_config (chave, valor) VALUES ('url_capital', ?)", (url_capital.strip(),))
    conn.commit()
    conn.close()
