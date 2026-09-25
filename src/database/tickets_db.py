import logging
from typing import Any
import pandas as pd
from datetime import datetime
from .connection import DB_PATH, get_connection

logger = logging.getLogger(__name__)

def setup_database():
    """Cria a tabela de chamados se não existir e garante as colunas."""
    conn = get_connection()
    cursor = conn.cursor()
    
    # Criando a tabela com as colunas básicas se não existir
    cursor.execute("""
    CREATE TABLE IF NOT EXISTS chamados (
        id TEXT PRIMARY KEY,
        data_criacao TEXT,
        titulo TEXT,
        cidade_predio TEXT,
        unidade TEXT,
        localidade_fisica TEXT,
        usuario TEXT,
        id_cliente TEXT,
        descricao TEXT,
        tag TEXT,
        ip_origem TEXT,
        status TEXT DEFAULT 'Aberto',
        data_atualizacao TEXT,
        base TEXT
    )
    """)
    
    # Criando a tabela de comentarios se não existir
    cursor.execute("""
    CREATE TABLE IF NOT EXISTS comentarios (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        chamado_id TEXT,
        data TEXT,
        autor TEXT,
        texto TEXT,
        FOREIGN KEY(chamado_id) REFERENCES chamados(id) ON DELETE CASCADE
    )
    """)
    
    # -------------------------------------------------------------
    # MIGRAÇÃO E LIMPEZA DE IDs DUPLICADOS COM SUFIXO DECIMAL (.0)
    # -------------------------------------------------------------
    cursor.execute("""
    DELETE FROM chamados 
    WHERE id LIKE '%.0' 
      AND substr(id, 1, length(id) - 2) IN (SELECT id FROM chamados)
    """)
    cursor.execute("""
    DELETE FROM comentarios 
    WHERE chamado_id LIKE '%.0' 
      AND substr(chamado_id, 1, length(chamado_id) - 2) IN (SELECT chamado_id FROM chamados)
    """)
    
    cursor.execute("""
    UPDATE chamados 
    SET id = substr(id, 1, length(id) - 2) 
    WHERE id LIKE '%.0'
    """)
    cursor.execute("""
    UPDATE comentarios 
    SET chamado_id = substr(chamado_id, 1, length(chamado_id) - 2) 
    WHERE chamado_id LIKE '%.0'
    """)
    # -------------------------------------------------------------

    # Garante as colunas no banco de dados (Migração automática)
    cursor.execute("PRAGMA table_info(chamados)")
    columns = [col[1] for col in cursor.fetchall()]
    if 'andamento' not in columns:
        cursor.execute("ALTER TABLE chamados ADD COLUMN andamento TEXT")
    if 'base' not in columns:
        cursor.execute("ALTER TABLE chamados ADD COLUMN base TEXT")
    if 'link' not in columns:
        cursor.execute("ALTER TABLE chamados ADD COLUMN link TEXT")
    if 'hostname' not in columns:
        cursor.execute("ALTER TABLE chamados ADD COLUMN hostname TEXT")
    if 'tag_manual' not in columns:
        cursor.execute("ALTER TABLE chamados ADD COLUMN tag_manual INTEGER DEFAULT 0")
    if 'dados_manuais' not in columns:
        cursor.execute("ALTER TABLE chamados ADD COLUMN dados_manuais INTEGER DEFAULT 0")

    conn.commit()
    conn.close()

def save_tickets_to_db(df: pd.DataFrame):
    """
    Insere novos chamados ou atualiza os existentes.
    Não altera o status para 'Aberto' se o chamado já estiver 'Fechado' manualmente.
    """
    setup_database()
    conn = get_connection()
    cursor = conn.cursor()
    
    now = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    
    def _clean_str(val: Any) -> str:
        if val is None or pd.isna(val):
            return ""
        s = str(val).strip()
        if s.lower() in ["nan", "none", "null", "<na>"]:
            return ""
        return s

    from .sccm_db import get_device_by_user

    for _, row in df.iterrows():
        cid = str(row.get('Chamado#', '')).strip()
        if cid.endswith('.0'):
            cid = cid[:-2]
            
        if not cid:
            continue
            
        titulo_val = _clean_str(row.get('Título', ''))
        usuario_val = _clean_str(row.get('Nome do Usuário', ''))
        id_cliente_val = _clean_str(row.get('ID do Cliente', ''))
        desc_val = _clean_str(row.get('Descrição', ''))
        ip_val = _clean_str(row.get('IP_Origem', ''))
        hostname_val = _clean_str(row.get('Hostname', ''))
        base_val = _clean_str(row.get('Base', ''))
        link_val = _clean_str(row.get('Link', ''))
        data_criacao_val = _clean_str(row.get('Data Criação', ''))
        tag_val = _clean_str(row.get('TAG', ''))
        cidade_predio_val = _clean_str(row.get('Cidade - Prédio', ''))
        unidade_val = _clean_str(row.get('Unidade', ''))
        localidade_val = _clean_str(row.get('Localidade física', ''))

        if "não encontrad" in unidade_val.lower() or "nao encontrad" in unidade_val.lower():
            unidade_val = ""
        if localidade_val.lower() in ["nan", "none", "null", "não identificada", ""]:
            localidade_val = "Não identificada"
        elif "nan -" in localidade_val.lower() or "- nan" in localidade_val.lower():
            localidade_val = "Não identificada"

        # Enriquecimento com cache do SCCM caso IP ou Hostname estejam vazios
        if (not ip_val or not hostname_val) and id_cliente_val:
            dev = get_device_by_user(id_cliente_val)
            if dev:
                if not ip_val and dev.get("ip"):
                    ip_val = dev["ip"]
                if not hostname_val and dev.get("hostname"):
                    hostname_val = dev["hostname"]

        cursor.execute("SELECT status, tag_manual, dados_manuais, ip_origem, hostname, titulo FROM chamados WHERE id = ?", (cid,))
        result = cursor.fetchone()
        
        if result:
            current_status = result[0]
            tag_manual = result[1] if len(result) > 1 and result[1] is not None else 0
            dados_manuais = result[2] if len(result) > 2 and result[2] is not None else 0
            curr_ip = _clean_str(result[3]) if len(result) > 3 else ""
            curr_host = _clean_str(result[4]) if len(result) > 4 else ""
            curr_title = _clean_str(result[5]) if len(result) > 5 else ""
            
            # Preserva IP, Hostname e Título existentes no banco se o novo for vazio
            if not ip_val and curr_ip:
                ip_val = curr_ip
            if not hostname_val and curr_host:
                hostname_val = curr_host
            if not titulo_val and curr_title:
                titulo_val = curr_title

            if current_status == 'Fechado':
                cursor.execute("UPDATE chamados SET status = 'Aberto' WHERE id = ?", (cid,))
                
            update_fields = []
            update_params = []
            
            update_fields.extend([
                "titulo = ?", "usuario = ?", "id_cliente = ?", "descricao = ?",
                "ip_origem = ?", "data_atualizacao = ?", "base = ?", "link = ?", "hostname = ?", "data_criacao = ?"
            ])
            update_params.extend([
                titulo_val, usuario_val, id_cliente_val, desc_val,
                ip_val, now, base_val, link_val, hostname_val, data_criacao_val
            ])
            
            if tag_manual != 1:
                update_fields.append("tag = ?")
                update_params.append(tag_val)
                
            if dados_manuais != 1 and current_status != 'Fechado':
                update_fields.extend(["cidade_predio = ?", "unidade = ?", "localidade_fisica = ?"])
                update_params.extend([cidade_predio_val, unidade_val, localidade_val])
                
            query = f"UPDATE chamados SET {', '.join(update_fields)} WHERE id = ?"
            update_params.append(cid)
            cursor.execute(query, tuple(update_params))

        else:
            cursor.execute("""
            INSERT INTO chamados (
                id, data_criacao, titulo, cidade_predio, unidade, localidade_fisica,
                usuario, id_cliente, descricao, tag, ip_origem, status, data_atualizacao, base, link, hostname
            ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, 'Aberto', ?, ?, ?, ?)
            """, (
                cid, data_criacao_val, titulo_val,
                cidade_predio_val, unidade_val, localidade_val,
                usuario_val, id_cliente_val, desc_val,
                tag_val, ip_val, now, base_val, link_val,
                hostname_val
            ))
            
        comments_val = row.get('Comentários', '[]')
        if pd.notna(comments_val) and str(comments_val).strip() and str(comments_val).strip() != '[]':
            from config import clean_otrs_comments
            try:
                comments_list = clean_otrs_comments(comments_val)
                cursor.execute("DELETE FROM comentarios WHERE chamado_id = ?", (cid,))
                for comment in comments_list:
                    cursor.execute("""
                    INSERT INTO comentarios (chamado_id, data, autor, texto)
                    VALUES (?, ?, ?, ?)
                    """, (cid, comment.get('data', ''), comment.get('autor', ''), comment.get('texto', '')))
            except Exception:
                pass
            
    conn.commit()
    conn.close()

def close_missing_tickets(active_ids: list):
    """
    Identifica chamados abertos no SQLite que não estão na lista de ativos recebida e marca status = 'Fechado'.
    """
    if not active_ids or len(active_ids) < 3:
        logger.warning(f"⚠️ Lista de IDs ativos muito pequena ({len(active_ids) if active_ids else 0}). Nenhum chamado foi fechado para evitar falsos positivos.")
        return 0
        
    active_set = {str(cid).strip()[:-2] if str(cid).strip().endswith('.0') else str(cid).strip() for cid in active_ids if cid}
    
    conn = get_connection()
    cursor = conn.cursor()
    now = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    
    cursor.execute("SELECT id FROM chamados WHERE status IS NULL OR status != 'Fechado'")
    db_open_rows = cursor.fetchall()
    
    closed_count = 0
    for (db_id,) in db_open_rows:
        db_id_clean = str(db_id).strip()[:-2] if str(db_id).strip().endswith('.0') else str(db_id).strip()
        if db_id_clean not in active_set:
            cursor.execute("UPDATE chamados SET status = 'Fechado', data_atualizacao = ? WHERE id = ?", (now, db_id))
            closed_count += 1
            
    conn.commit()
    conn.close()
    
    logger.info(f"✅ Fechamento automático (Geral): {closed_count} chamados atualizados para 'Fechado'.")
    return closed_count

def close_missing_tickets_by_base(active_ids: list, base: str):
    """
    Identifica chamados abertos no SQLite para uma base específica ('OTRS' ou 'CitSmart') 
    que não estão na lista de ativos recebida e marca status = 'Fechado'.
    """
    if not active_ids or len(active_ids) < 3:
        logger.warning(f"⚠️ Lista de IDs ativos para {base} muito pequena ({len(active_ids) if active_ids else 0}). Nenhum chamado de {base} foi fechado para evitar falsos positivos.")
        return 0
        
    active_set = {str(cid).strip()[:-2] if str(cid).strip().endswith('.0') else str(cid).strip() for cid in active_ids if cid}
    
    conn = get_connection()
    cursor = conn.cursor()
    now = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    
    cursor.execute("SELECT id FROM chamados WHERE base = ? AND (status IS NULL OR status != 'Fechado')", (base,))
    db_open_rows = cursor.fetchall()
    
    closed_count = 0
    for (db_id,) in db_open_rows:
        db_id_clean = str(db_id).strip()[:-2] if str(db_id).strip().endswith('.0') else str(db_id).strip()
        if db_id_clean not in active_set:
            cursor.execute("UPDATE chamados SET status = 'Fechado', data_atualizacao = ? WHERE id = ?", (now, db_id))
            closed_count += 1
            
    conn.commit()
    conn.close()
    
    logger.info(f"✅ Fechamento automático ({base}): {closed_count} chamados atualizados para 'Fechado'.")
    return closed_count

def update_ticket_status(cid: str, new_status: str):
    """Atualiza o status de um chamado específico (usado pelo Streamlit)."""
    conn = get_connection()
    cursor = conn.cursor()
    now = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    
    cid_clean = str(cid)[:-2] if str(cid).endswith('.0') else str(cid)
    
    cursor.execute("""
    UPDATE chamados 
    SET status = ?, data_atualizacao = ?
    WHERE id = ?
    """, (new_status, now, cid_clean))
    
    conn.commit()
    conn.close()

def update_ticket_andamento(cid: str, text: str):
    """Atualiza o andamento/notas rápidas de um chamado específico."""
    conn = get_connection()
    cursor = conn.cursor()
    now = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    
    cid_clean = str(cid)[:-2] if str(cid).endswith('.0') else str(cid)
    
    cursor.execute("""
    UPDATE chamados 
    SET andamento = ?, data_atualizacao = ?
    WHERE id = ?
    """, (text, now, cid_clean))
    
    conn.commit()
    conn.close()

def update_ticket_tag(cid: str, new_tag: str):
    """Atualiza a TAG de um chamado específico e marca como manual (tag_manual = 1)."""
    conn = get_connection()
    cursor = conn.cursor()
    now = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    
    cid_clean = str(cid)[:-2] if str(cid).endswith('.0') else str(cid)
    
    cursor.execute("""
    UPDATE chamados 
    SET tag = ?, tag_manual = 1, data_atualizacao = ?
    WHERE id = ?
    """, (new_tag, now, cid_clean))
    
    conn.commit()
    conn.close()

def update_ticket_location_details(cid: str, localidade: str, cidade_predio: str, unidade: str):
    """Atualiza localidade_fisica, cidade_predio, unidade de um chamado e marca dados_manuais = 1."""
    conn = get_connection()
    cursor = conn.cursor()
    now = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    
    cid_clean = str(cid)[:-2] if str(cid).endswith('.0') else str(cid)
    
    cursor.execute("""
    UPDATE chamados 
    SET localidade_fisica = ?, cidade_predio = ?, unidade = ?, dados_manuais = 1, data_atualizacao = ?
    WHERE id = ?
    """, (localidade, cidade_predio, unidade, now, cid_clean))
    
    conn.commit()
    conn.close()

def update_ticket_title(cid: str, new_title: str):
    """Atualiza o título de um chamado específico no SQLite."""
    conn = get_connection()
    cursor = conn.cursor()
    now = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    
    cid_clean = str(cid)[:-2] if str(cid).endswith('.0') else str(cid)
    
    cursor.execute("""
    UPDATE chamados 
    SET titulo = ?, data_atualizacao = ?
    WHERE id = ?
    """, (new_title.strip(), now, cid_clean))
    
    conn.commit()
    conn.close()

def update_ticket_device_info(cid: str, ip_origem: str, hostname: str):
    """Atualiza o IP de origem e Hostname de um chamado específico."""
    conn = get_connection()
    cursor = conn.cursor()
    now = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    
    cid_clean = str(cid)[:-2] if str(cid).endswith('.0') else str(cid)
    
    clean_ip = str(ip_origem).strip() if ip_origem and str(ip_origem).lower() not in ["none", "nan", "null"] else ""
    clean_host = str(hostname).strip() if hostname and str(hostname).lower() not in ["none", "nan", "null"] else ""
    
    cursor.execute("""
    UPDATE chamados 
    SET ip_origem = ?, hostname = ?, data_atualizacao = ?
    WHERE id = ?
    """, (clean_ip, clean_host, now, cid_clean))
    
    conn.commit()
    conn.close()

def save_comments_to_db(chamado_id: str, comments: list):
    """
    Salva a lista de comentários de um chamado no banco de dados.
    """
    setup_database()
    conn = get_connection()
    cursor = conn.cursor()
    
    cursor.execute("DELETE FROM comentarios WHERE chamado_id = ?", (chamado_id,))
    
    from config import clean_otrs_comments
    cleaned_comments = clean_otrs_comments(comments)
    for comment in cleaned_comments:
        cursor.execute("""
        INSERT INTO comentarios (chamado_id, data, autor, texto)
        VALUES (?, ?, ?, ?)
        """, (chamado_id, comment.get('data', ''), comment.get('autor', ''), comment.get('texto', '')))
        
    conn.commit()
    conn.close()

def get_comments_by_ticket(chamado_id: str) -> list:
    """Retorna todos os comentários de um chamado ordenados por id (filtrando robôs)."""
    setup_database()
    conn = get_connection()
    cursor = conn.cursor()
    cursor.execute("""
    SELECT data, autor, texto FROM comentarios 
    WHERE chamado_id = ? 
    ORDER BY id ASC
    """, (chamado_id,))
    rows = cursor.fetchall()
    conn.close()
    
    raw_comments = [{'data': r[0], 'autor': r[1], 'texto': r[2]} for r in rows]
    from config import clean_otrs_comments
    return clean_otrs_comments(raw_comments)

def sync_closed_tickets_to_train_dataset():
    """
    Coleta todos os chamados com status 'Fechado' do banco SQLite e os envia/sincroniza 
    com o arquivo de dataset de treino (Chamados_Treino.xlsx).
    """
    from config import TREINO_PATH
    
    conn = get_connection()
    cursor = conn.cursor()
    
    cursor.execute("""
    SELECT id, usuario, data_criacao, tag, cidade_predio, unidade, andamento, descricao, base
    FROM chamados
    WHERE status = 'Fechado'
    """)
    rows = cursor.fetchall()
    conn.close()
    
    if not rows:
        return
        
    df_closed = pd.DataFrame(rows, columns=[
        'Chamado#', 'Nome do Usuário', 'Data Criação', 'TAG', 
        'Cidade - Prédio', 'Unidade', 'Andamento', 'Descrição', 'Base'
    ])
    
    df_closed['Ramal'] = ""
    
    try:
        if TREINO_PATH.exists():
            df_treino_atual = pd.read_excel(TREINO_PATH)
            df_treino_atual['Chamado#'] = df_treino_atual['Chamado#'].astype(str).str.strip().str.replace(r'\.0$', '', regex=True)
            df_closed['Chamado#'] = df_closed['Chamado#'].astype(str).str.strip().str.replace(r'\.0$', '', regex=True)
            
            df_treino_novo = pd.concat([df_treino_atual, df_closed], ignore_index=True)
        else:
            df_treino_novo = df_closed
            
        df_treino_novo = df_treino_novo.drop_duplicates(subset=['Chamado#'], keep='last')
        df_treino_novo = df_treino_novo.fillna("")
        
        cols_order = ['Chamado#', 'Nome do Usuário', 'Data Criação', 'TAG', 'Cidade - Prédio', 'Unidade', 'Ramal', 'Andamento', 'Descrição', 'Base']
        for col in cols_order:
            if col not in df_treino_novo.columns:
                df_treino_novo[col] = ""
        df_treino_novo = df_treino_novo[cols_order]
        
        TREINO_PATH.parent.mkdir(parents=True, exist_ok=True)
        df_treino_novo.to_excel(TREINO_PATH, index=False)
        
    except Exception as e:
        print(f"Erro ao sincronizar chamados fechados com o dataset de treino: {e}")

def load_data():
    """Carrega todos os chamados da tabela SQLite em um DataFrame pandas."""
    if not DB_PATH.exists():
        return pd.DataFrame()
    conn = get_connection()
    df = pd.read_sql_query("SELECT * FROM chamados", conn)
    conn.close()
    
    if 'localidade_fisica' in df.columns:
        import re
        df['localidade_fisica'] = df['localidade_fisica'].apply(
            lambda x: re.sub(r'\s*-\s*Sede\b', '', str(x), flags=re.IGNORECASE).strip() if pd.notna(x) else x
        )

    # Sanitização preventiva para evitar que 'nan', 'none' ou 'null' literais cheguem à interface
    for col in ['ip_origem', 'hostname', 'titulo', 'id_cliente', 'andamento', 'cidade_predio', 'unidade']:
        if col in df.columns:
            df[col] = df[col].fillna("").astype(str).replace(r'^(?i:nan|none|null|<na>)$', '', regex=True).str.strip()

    if 'unidade' in df.columns:
        df['unidade'] = df['unidade'].apply(
            lambda x: "" if any(term in str(x).lower() for term in ["não encontrad", "nao encontrad"]) else str(x).strip()
        )

    if 'localidade_fisica' in df.columns:
        def _clean_loc_series(val):
            if pd.isna(val):
                return "Não identificada"
            s = str(val).strip()
            if s.lower() in ["nan", "none", "null", "<na>", "", "não identificada"]:
                return "Não identificada"
            if "nan -" in s.lower() or "- nan" in s.lower():
                return "Não identificada"
            return s
        df['localidade_fisica'] = df['localidade_fisica'].apply(_clean_loc_series)

    return df
