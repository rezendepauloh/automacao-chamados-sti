import sqlite3
import pandas as pd
from pathlib import Path
from datetime import datetime

DB_PATH = Path("chamados.db")

def get_connection():
    """Retorna uma conexão com o banco de dados SQLite."""
    return sqlite3.connect(DB_PATH)

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

def get_ramais_df() -> pd.DataFrame:
    """Retorna os dados da tabela ramais_mpms em um DataFrame."""
    setup_ramais_table()
    conn = get_connection()
    try:
        df = pd.read_sql_query("SELECT id, localidade, setor_nome, telefone_ramal, tipo, data_atualizacao FROM ramais_mpms", conn)
    except Exception:
        df = pd.DataFrame(columns=["id", "localidade", "setor_nome", "telefone_ramal", "tipo", "data_atualizacao"])
    conn.close()
    return df

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
    # 1. Remove duplicatas terminadas em .0 no banco caso a versão inteira correspondente já exista
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
    
    # 2. Converte os restantes que restaram com .0 para inteiro
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
    
    for _, row in df.iterrows():
        cid = str(row.get('Chamado#', '')).strip()
        # Normaliza IDs removendo o .0 se for interpretado como float pelo Pandas
        if cid.endswith('.0'):
            cid = cid[:-2]
            
        if not cid:
            continue
            
        # Verifica se o chamado já existe, qual o status e se a tag foi alterada manualmente
        cursor.execute("SELECT status, tag_manual, dados_manuais FROM chamados WHERE id = ?", (cid,))
        result = cursor.fetchone()
        
        if result:
            current_status = result[0]
            tag_manual = result[1] if len(result) > 1 and result[1] is not None else 0
            dados_manuais = result[2] if len(result) > 2 and result[2] is not None else 0
            
            # Se estava marcado como Fechado mas foi re-coletado como ativo pelo robô, reabre no banco!
            if current_status == 'Fechado':
                cursor.execute("UPDATE chamados SET status = 'Aberto' WHERE id = ?", (cid,))
                
            # Monta query dinâmica de update para evitar sobrescrever dados manuais
            update_fields = []
            update_params = []
            
            # Campos comuns
            update_fields.extend([
                "titulo = ?", "usuario = ?", "id_cliente = ?", "descricao = ?",
                "ip_origem = ?", "data_atualizacao = ?", "base = ?", "link = ?", "hostname = ?", "data_criacao = ?"
            ])
            update_params.extend([
                row.get('Título', ''), row.get('Nome do Usuário', ''), row.get('ID do Cliente', ''), row.get('Descrição', ''),
                row.get('IP_Origem', ''), now, row.get('Base', ''), row.get('Link', ''), row.get('Hostname', ''), row.get('Data Criação', '')
            ])
            
            # Se a tag NÃO for manual, atualiza
            if tag_manual != 1:
                update_fields.append("tag = ?")
                update_params.append(row.get('TAG', ''))
                
            # Se a localidade/prédio/unidade NÃO for manual e o chamado NÃO estiver Fechado, atualiza
            if dados_manuais != 1 and current_status != 'Fechado':
                update_fields.extend(["cidade_predio = ?", "unidade = ?", "localidade_fisica = ?"])
                update_params.extend([row.get('Cidade - Prédio', ''), row.get('Unidade', ''), row.get('Localidade física', '')])
                
            query = f"UPDATE chamados SET {', '.join(update_fields)} WHERE id = ?"
            update_params.append(cid)
            cursor.execute(query, tuple(update_params))

        else:
            # Se não existe, insere como Aberto
            cursor.execute("""
            INSERT INTO chamados (
                id, data_criacao, titulo, cidade_predio, unidade, localidade_fisica,
                usuario, id_cliente, descricao, tag, ip_origem, status, data_atualizacao, base, link, hostname
            ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, 'Aberto', ?, ?, ?, ?)
            """, (
                cid, row.get('Data Criação', ''), row.get('Título', ''),
                row.get('Cidade - Prédio', ''), row.get('Unidade', ''), row.get('Localidade física', ''),
                row.get('Nome do Usuário', ''), row.get('ID do Cliente', ''), row.get('Descrição', ''),
                row.get('TAG', ''), row.get('IP_Origem', ''), now, row.get('Base', ''), row.get('Link', ''),
                row.get('Hostname', '')
            ))
            
        # Salva os comentários se a coluna 'Comentários' estiver presente (Atômico na mesma transação)
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
    Marca como 'Fechado' os chamados que estão no banco como 'Aberto'
    mas não estão na lista de IDs ativos (que vieram da última coleta).
    """
    if not active_ids:
        return
        
    # Normaliza IDs na entrada limpando .0
    active_ids_clean = [str(cid)[:-2] if str(cid).endswith('.0') else str(cid) for cid in active_ids]
        
    conn = get_connection()
    cursor = conn.cursor()
    
    now = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    
    placeholders = ",".join("?" for _ in active_ids_clean)
    
    cursor.execute(f"""
    UPDATE chamados 
    SET status = 'Fechado', data_atualizacao = ?
    WHERE status = 'Aberto' AND id NOT IN ({placeholders})
    """, [now] + [str(cid) for cid in active_ids_clean])
    
    conn.commit()
    conn.close()

def close_missing_tickets_by_base(active_ids: list, base: str):
    """
    Marca como 'Fechado' os chamados de uma base específica (OTRS ou CitSmart) que estão 
    no banco como 'Aberto' mas não estão na lista de IDs ativos da última coleta.
    Garante que falhas de scrapers ou coletas vazias de uma base não fechem chamados da outra.
    """
    if not active_ids or not base:
        return
        
    # Normaliza IDs na entrada limpando .0
    active_ids_clean = [str(cid)[:-2] if str(cid).endswith('.0') else str(cid) for cid in active_ids]
        
    conn = get_connection()
    cursor = conn.cursor()
    
    now = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    
    placeholders = ",".join("?" for _ in active_ids_clean)
    
    cursor.execute(f"""
    UPDATE chamados 
    SET status = 'Fechado', data_atualizacao = ?
    WHERE status = 'Aberto' AND base = ? AND id NOT IN ({placeholders})
    """, [now, base] + [str(cid) for cid in active_ids_clean])
    
    conn.commit()
    conn.close()

def update_ticket_status(cid: str, new_status: str):
    """Atualiza o status de um chamado específico (usado pelo Streamlit)."""
    conn = get_connection()
    cursor = conn.cursor()
    now = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    
    # Normaliza o ID na entrada limpando .0
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
    
    # Normaliza o ID na entrada limpando .0
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
    
    # Normaliza o ID na entrada limpando .0
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
    
    # Normaliza o ID na entrada limpando .0
    cid_clean = str(cid)[:-2] if str(cid).endswith('.0') else str(cid)
    
    cursor.execute("""
    UPDATE chamados 
    SET localidade_fisica = ?, cidade_predio = ?, unidade = ?, dados_manuais = 1, data_atualizacao = ?
    WHERE id = ?
    """, (localidade, cidade_predio, unidade, now, cid_clean))
    
    conn.commit()
    conn.close()

def save_comments_to_db(chamado_id: str, comments: list):
    """
    Salva a lista de comentários de um chamado no banco de dados.
    Cada comentário na lista deve ser um dicionário: {'data': '...', 'autor': '...', 'texto': '...'}
    Remove os comentários antigos daquele chamado antes de inserir os novos para evitar duplicados.
    """
    setup_database()
    conn = get_connection()
    cursor = conn.cursor()
    
    # Remove antigos para evitar duplicidade
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
    com o arquivo de dataset de treino (Chamados_Treino.xlsx) para que a IA aprenda 
    com as tags corretas e manuais ajustadas pelo Streamlit.
    """
    from config import TREINO_PATH
    
    conn = get_connection()
    cursor = conn.cursor()
    
    # Seleciona as colunas correspondentes ao dataset de treino
    cursor.execute("""
    SELECT id, usuario, data_criacao, tag, cidade_predio, unidade, andamento, descricao, base
    FROM chamados
    WHERE status = 'Fechado'
    """)
    rows = cursor.fetchall()
    conn.close()
    
    if not rows:
        return
        
    # Transforma em DataFrame com os nomes de colunas originais do Excel de Treino
    df_closed = pd.DataFrame(rows, columns=[
        'Chamado#', 'Nome do Usuário', 'Data Criação', 'TAG', 
        'Cidade - Prédio', 'Unidade', 'Andamento', 'Descrição', 'Base'
    ])
    
    # Adiciona colunas extras que podem existir no dataset de treino original
    df_closed['Ramal'] = ""
    
    try:
        if TREINO_PATH.exists():
            df_treino_atual = pd.read_excel(TREINO_PATH)
            # Garante que Chamado# é string para bater certo e não duplicar
            df_treino_atual['Chamado#'] = df_treino_atual['Chamado#'].astype(str).str.strip().str.replace(r'\.0$', '', regex=True)
            df_closed['Chamado#'] = df_closed['Chamado#'].astype(str).str.strip().str.replace(r'\.0$', '', regex=True)
            
            df_treino_novo = pd.concat([df_treino_atual, df_closed], ignore_index=True)
        else:
            df_treino_novo = df_closed
            
        # Remove duplicatas mantendo a última versão (que contém a tag manual ou andamento atualizados)
        df_treino_novo = df_treino_novo.drop_duplicates(subset=['Chamado#'], keep='last')
        df_treino_novo = df_treino_novo.fillna("")
        
        # Garante a ordem correta das colunas
        cols_order = ['Chamado#', 'Nome do Usuário', 'Data Criação', 'TAG', 'Cidade - Prédio', 'Unidade', 'Ramal', 'Andamento', 'Descrição', 'Base']
        # Se alguma coluna não existir, adiciona
        for col in cols_order:
            if col not in df_treino_novo.columns:
                df_treino_novo[col] = ""
        df_treino_novo = df_treino_novo[cols_order]
        
        TREINO_PATH.parent.mkdir(parents=True, exist_ok=True)
        df_treino_novo.to_excel(TREINO_PATH, index=False)
        
    except Exception as e:
        print(f"Erro ao sincronizar chamados fechados com o dataset de treino: {e}")


def setup_map_tables():
    """Cria as tabelas do mapa se não existirem."""
    conn = get_connection()
    cursor = conn.cursor()
    
    # Tabela para guardar a configuração estrutural de prédios/pavimentos em JSON
    cursor.execute("""
    CREATE TABLE IF NOT EXISTS mapa_config (
        id TEXT PRIMARY KEY,
        config_json TEXT
    )
    """)
    
    # Tabela para indexar os pins de busca rápida
    cursor.execute("""
    CREATE TABLE IF NOT EXISTS mapa_pins (
        id TEXT PRIMARY KEY,
        predio_id TEXT,
        pavimento_id INTEGER,
        sala TEXT,
        x INTEGER,
        y INTEGER,
        descricao TEXT
    )
    """)
    conn.commit()
    conn.close()


def save_map_config(config_data: dict):
    """Salva as configurações de prédios e os pins no banco de dados SQLite."""
    import json
    setup_map_tables()
    conn = get_connection()
    cursor = conn.cursor()
    
    # Salva a estrutura de prédios e pavimentos (preservando pins aninhados se houver)
    predios = config_data.get("predios", [])
    config_json_str = json.dumps({"predios": predios})
    cursor.execute("INSERT OR REPLACE INTO mapa_config (id, config_json) VALUES ('config_atual', ?)", (config_json_str,))
    
    # Limpa pins antigos e insere os novos
    cursor.execute("DELETE FROM mapa_pins")
    
    # 1. Insere pins do novo formato aninhado dentro de cada prédio
    for predio in predios:
        p_id = predio.get("id")
        p_pins = predio.get("pins", [])
        for pin in p_pins:
            cursor.execute("""
            INSERT OR REPLACE INTO mapa_pins (id, predio_id, pavimento_id, sala, x, y, descricao)
            VALUES (?, ?, ?, ?, ?, ?, ?)
            """, (
                str(pin.get("id")),
                str(pin.get("predio_id", p_id)),
                int(pin.get("pavimento_id")),
                str(pin.get("sala")),
                int(pin.get("x")),
                int(pin.get("y")),
                str(pin.get("descricao", ""))
            ))
            
    # 2. Insere pins do formato antigo plano (nível raiz) para retrocompatibilidade
    pins = config_data.get("pins", [])
    for pin in pins:
        cursor.execute("""
        INSERT OR REPLACE INTO mapa_pins (id, predio_id, pavimento_id, sala, x, y, descricao)
        VALUES (?, ?, ?, ?, ?, ?, ?)
        """, (
            str(pin.get("id")),
            str(pin.get("predio_id")),
            int(pin.get("pavimento_id")),
            str(pin.get("sala")),
            int(pin.get("x")),
            int(pin.get("y")),
            str(pin.get("descricao", ""))
        ))
        
    conn.commit()
    conn.close()


def get_map_config() -> dict:
    """Retorna a configuração atual de prédios e pavimentos."""
    import json
    setup_map_tables()
    conn = get_connection()
    cursor = conn.cursor()
    cursor.execute("SELECT config_json FROM mapa_config WHERE id = 'config_atual'")
    row = cursor.fetchone()
    conn.close()
    if row:
        return json.loads(row[0])
    return {"predios": []}


def get_map_pins(predio_id=None, pavimento_id=None) -> list:
    """Retorna os pins cadastrados, podendo filtrar por prédio e pavimento."""
    setup_map_tables()
    conn = get_connection()
    cursor = conn.cursor()
    
    if predio_id is not None and pavimento_id is not None:
        cursor.execute("""
        SELECT id, predio_id, pavimento_id, sala, x, y, descricao 
        FROM mapa_pins 
        WHERE predio_id = ? AND pavimento_id = ?
        """, (predio_id, pavimento_id))
    elif predio_id is not None:
        cursor.execute("""
        SELECT id, predio_id, pavimento_id, sala, x, y, descricao 
        FROM mapa_pins 
        WHERE predio_id = ?
        """, (predio_id,))
    elif pavimento_id is not None:
        cursor.execute("""
        SELECT id, predio_id, pavimento_id, sala, x, y, descricao 
        FROM mapa_pins 
        WHERE pavimento_id = ?
        """, (pavimento_id,))
    else:
        cursor.execute("""
        SELECT id, predio_id, pavimento_id, sala, x, y, descricao 
        FROM mapa_pins
        """)
        
    rows = cursor.fetchall()
    conn.close()
    
    return [
        {
            "id": r[0],
            "predio_id": r[1],
            "pavimento_id": r[2],
            "sala": r[3],
            "x": r[4],
            "y": r[5],
            "descricao": r[6]
        }
        for r in rows
    ]


def setup_donations_table():
    """Cria a tabela de equipamentos doados se não existir."""
    conn = get_connection()
    cursor = conn.cursor()
    cursor.execute("""
    CREATE TABLE IF NOT EXISTS equipamentos_doados (
        patrimonio TEXT,
        modelo TEXT,
        serial_number TEXT,
        equipamento TEXT,
        tipo_movimentacao TEXT,
        data_movimentacao TEXT,
        chamado TEXT,
        ssd TEXT,
        motivo_baixa TEXT
    )
    """)
    conn.commit()
    conn.close()


def sync_donations_from_excel(file_path: str):
    """Lê os dados da planilha de equipamentos doados e salva no SQLite."""
    import pandas as pd
    from datetime import datetime
    from src.config import setup_logging, DEBUG_DIR_DONATIONS
    
    logger = setup_logging(DEBUG_DIR_DONATIONS / "donations.log", "donations_sync")
    logger.info(f"Iniciando sincronização da planilha: {file_path}")
    
    setup_donations_table()
    
    try:
        # Lê a aba correta da planilha
        df = pd.read_excel(file_path, sheet_name="Equipamentos doados")
        logger.info(f"Planilha lida com sucesso. Total de linhas encontradas: {len(df)}")
    except Exception as e:
        logger.error(f"Erro ao ler a planilha Excel: {e}")
        raise e
    
    # Limpa dados nulos/NaN e garante tipos de dados corretos
    df = df.fillna("")
    
    conn = get_connection()
    cursor = conn.cursor()
    
    # Limpa a tabela existente para sincronização completa
    cursor.execute("DELETE FROM equipamentos_doados")
    
    added_count = 0
    for _, row in df.iterrows():
        # Formatação de datas a partir do pandas/excel (verificando NaT)
        dt_val = row.get('Data da doação', '')
        if isinstance(dt_val, (datetime, pd.Timestamp)) and not pd.isna(dt_val):
            dt_str = dt_val.strftime("%Y-%m-%d")
        else:
            dt_str = str(dt_val).strip()
            # tenta formatar se for no formato YYYY-MM-DD HH:MM:SS
            if len(dt_str) > 10:
                dt_str = dt_str[:10]
            if dt_str.lower() in ["nat", "nan", "null", ""]:
                dt_str = ""
        
        patrimonio = str(row.get('Patrimônio', '')).strip().replace(".0", "")
        chamado = str(row.get('Tem chamado?', '')).strip().replace(".0", "")
        
        cursor.execute("""
        INSERT INTO equipamentos_doados (
            patrimonio, modelo, serial_number, equipamento, 
            tipo_movimentacao, data_movimentacao, chamado, ssd, motivo_baixa
        ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?)
        """, (
            patrimonio,
            str(row.get('Modelo', '')).strip(),
            str(row.get('Serial Number PC', '')).strip(),
            str(row.get('Equipamento', '')).strip(),
            str(row.get('Doação ou Baixa', '')).strip(),
            dt_str,
            chamado,
            str(row.get('SSD', '')).strip(),
            str(row.get('Motivo baixa', '')).strip()
        ))
        added_count += 1
        
    conn.commit()
    conn.close()
    logger.info(f"Sincronização concluída. {added_count} registros importados para a tabela equipamentos_doados.")

        
    conn.commit()
    conn.close()


def get_donations_data() -> pd.DataFrame:
    """Retorna todos os equipamentos doados/movimentados da tabela SQLite."""
    setup_donations_table()
    conn = get_connection()
    df = pd.read_sql_query("SELECT * FROM equipamentos_doados", conn)
    conn.close()
    return df


def load_data():
    """Carrega todos os chamados da tabela SQLite em um DataFrame pandas."""
    if not DB_PATH.exists():
        return pd.DataFrame()
    conn = get_connection()
    df = pd.read_sql_query("SELECT * FROM chamados", conn)
    conn.close()
    
    # Limpa " - Sede" de forma inteligente na coluna de exibição da Localidade Física
    if 'localidade_fisica' in df.columns:
        import re
        df['localidade_fisica'] = df['localidade_fisica'].apply(
            lambda x: re.sub(r'\s*-\s*Sede\b', '', str(x), flags=re.IGNORECASE).strip() if pd.notna(x) else x
        )
    return df


def setup_plantoes_tables():
    """Cria as tabelas de plantão matutino e plantão semanal se não existirem."""
    conn = get_connection()
    cursor = conn.cursor()
    
    cursor.execute("""
    CREATE TABLE IF NOT EXISTS plantoes_matutino (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        ano INTEGER,
        data_iso TEXT,
        dia_semana TEXT,
        servidor TEXT,
        telefone TEXT,
        UNIQUE(ano, data_iso, servidor) ON CONFLICT REPLACE
    )
    """)
    
    cursor.execute("""
    CREATE TABLE IF NOT EXISTS plantoes_semanal (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        ano INTEGER,
        mes TEXT,
        periodo_str TEXT,
        data_inicio TEXT,
        data_fim TEXT,
        service_desk TEXT,
        manutencao TEXT,
        infraestrutura TEXT,
        desenvolvimento TEXT,
        UNIQUE(ano, data_inicio, manutencao) ON CONFLICT REPLACE
    )
    """)
    
    conn.commit()
    conn.close()


def save_plantoes_matutino(records: list[dict]):
    """Salva ou atualiza registros de plantão matutino no SQLite."""
    setup_plantoes_tables()
    if not records:
        return
    conn = get_connection()
    cursor = conn.cursor()
    for r in records:
        cursor.execute("""
        INSERT OR REPLACE INTO plantoes_matutino (ano, data_iso, dia_semana, servidor, telefone)
        VALUES (?, ?, ?, ?, ?)
        """, (
            r.get("ano"),
            r.get("data_iso"),
            r.get("dia_semana"),
            r.get("servidor"),
            r.get("telefone")
        ))
    conn.commit()
    conn.close()


def save_plantoes_semanal(records: list[dict]):
    """Salva ou atualiza registros de plantão semanal SIMP no SQLite."""
    setup_plantoes_tables()
    if not records:
        return
    conn = get_connection()
    cursor = conn.cursor()
    for r in records:
        cursor.execute("""
        INSERT OR REPLACE INTO plantoes_semanal 
        (ano, mes, periodo_str, data_inicio, data_fim, service_desk, manutencao, infraestrutura, desenvolvimento)
        VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?)
        """, (
            r.get("ano"),
            r.get("mes"),
            r.get("periodo_str"),
            r.get("data_inicio"),
            r.get("data_fim"),
            r.get("service_desk"),
            r.get("manutencao"),
            r.get("infraestrutura"),
            r.get("desenvolvimento")
        ))
    conn.commit()
    conn.close()


def get_plantoes_matutino(ano: int | None = None) -> pd.DataFrame:
    """Retorna DataFrame pandas dos plantões matutinos cadastrados."""
    setup_plantoes_tables()
    conn = get_connection()
    if ano:
        df = pd.read_sql_query("SELECT * FROM plantoes_matutino WHERE ano = ? ORDER BY data_iso ASC", conn, params=(ano,))
    else:
        df = pd.read_sql_query("SELECT * FROM plantoes_matutino ORDER BY data_iso ASC", conn)
    conn.close()
    return df


def get_plantoes_semanal(ano: int | None = None) -> pd.DataFrame:
    """Retorna DataFrame pandas dos plantões semanais SIMP cadastrados."""
    setup_plantoes_tables()
    conn = get_connection()
    if ano:
        df = pd.read_sql_query("SELECT * FROM plantoes_semanal WHERE ano = ? ORDER BY data_inicio ASC", conn, params=(ano,))
    else:
        df = pd.read_sql_query("SELECT * FROM plantoes_semanal ORDER BY data_inicio ASC", conn)
    conn.close()
    return df


# -----------------------------------------------------------------------------
# TABELA E FUNÇÕES DE NOTIFICAÇÃO
# -----------------------------------------------------------------------------

def setup_notifications_table():
    """Cria a tabela de notificações se não existir."""
    conn = get_connection()
    cursor = conn.cursor()
    cursor.execute("""
    CREATE TABLE IF NOT EXISTS notificacoes (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        tipo TEXT NOT NULL,
        titulo TEXT NOT NULL,
        mensagem TEXT NOT NULL,
        data_evento TEXT,
        link_pagina TEXT,
        lida INTEGER DEFAULT 0,
        data_criacao DATETIME DEFAULT CURRENT_TIMESTAMP,
        UNIQUE(tipo, titulo, data_evento) ON CONFLICT IGNORE
    )
    """)
    conn.commit()
    conn.close()


def add_notification(tipo: str, titulo: str, mensagem: str, data_evento: str = "", link_pagina: str = "") -> bool:
    """Adiciona uma notificação no banco de dados. Retorna True se foi inserida uma nova notificação."""
    setup_notifications_table()
    conn = get_connection()
    cursor = conn.cursor()
    try:
        cursor.execute("""
        INSERT INTO notificacoes (tipo, titulo, mensagem, data_evento, link_pagina, lida)
        VALUES (?, ?, ?, ?, ?, 0)
        """, (tipo, titulo, mensagem, data_evento, link_pagina))
        inserted = cursor.rowcount > 0
        conn.commit()
        return inserted
    except Exception:
        return False
    finally:
        conn.close()


def get_notifications(only_unread: bool = False, limit: int = 100) -> pd.DataFrame:
    """Retorna um DataFrame pandas com as notificações do banco."""
    setup_notifications_table()
    conn = get_connection()
    query = "SELECT * FROM notificacoes"
    if only_unread:
        query += " WHERE lida = 0"
    query += " ORDER BY id DESC LIMIT ?"
    df = pd.read_sql_query(query, conn, params=(limit,))
    conn.close()
    return df


def get_unread_notifications_count() -> int:
    """Retorna a quantidade de notificações não lidas."""
    setup_notifications_table()
    conn = get_connection()
    cursor = conn.cursor()
    cursor.execute("SELECT COUNT(*) FROM notificacoes WHERE lida = 0")
    count = cursor.fetchone()[0]
    conn.close()
    return count


def mark_notification_as_read(notif_id: int):
    """Marca uma notificação específica como lida."""
    setup_notifications_table()
    conn = get_connection()
    cursor = conn.cursor()
    cursor.execute("UPDATE notificacoes SET lida = 1 WHERE id = ?", (notif_id,))
    conn.commit()
    conn.close()


def mark_all_notifications_as_read():
    """Marca todas as notificações como lidas."""
    setup_notifications_table()
    conn = get_connection()
    cursor = conn.cursor()
    cursor.execute("UPDATE notificacoes SET lida = 1 WHERE lida = 0")
    conn.commit()
    conn.close()


def setup_impressoras_table():
    """Cria a tabela de impressoras/dispositivos do PaperCut se não existir."""
    conn = get_connection()
    cursor = conn.cursor()
    cursor.execute("""
    CREATE TABLE IF NOT EXISTS impressoras (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        nome TEXT UNIQUE,
        servidor TEXT,
        tipo TEXT,
        modelo TEXT,
        localizacao TEXT,
        ip_host TEXT,
        status TEXT,
        total_paginas INTEGER DEFAULT 0,
        filas_relacionadas TEXT,
        detalhes_extra TEXT,
        data_atualizacao TEXT
    )
    """)
    conn.commit()
    conn.close()


def save_impressoras_to_db(df: pd.DataFrame):
    """
    Salva ou atualiza os dados de impressoras/dispositivos do PaperCut no banco de dados.
    """
    setup_impressoras_table()
    if df is None or df.empty:
        return

    conn = get_connection()
    cursor = conn.cursor()
    now_str = datetime.now().strftime("%Y-%m-%d %H:%M:%S")

    # Limpa registros antigos antes de reescrever a base unificada
    cursor.execute("DELETE FROM impressoras")

    for _, row in df.iterrows():

        nome = str(row.get('nome', '')).strip()
        if not nome or nome.lower() in ['nan', 'none', '']:
            continue

        servidor = str(row.get('servidor', '')).strip()
        tipo = str(row.get('tipo', 'Impressora')).strip()
        modelo = str(row.get('modelo', '')).strip()
        localizacao = str(row.get('localizacao', '')).strip()
        ip_host = str(row.get('ip_host', '')).strip()
        status = str(row.get('status', 'OK')).strip()
        
        try:
            total_paginas = int(row.get('total_paginas', 0))
        except (ValueError, TypeError):
            total_paginas = 0

        filas_relacionadas = str(row.get('filas_relacionadas', '')).strip()
        detalhes_extra = str(row.get('detalhes_extra', '')).strip()

        cursor.execute("""
        INSERT INTO impressoras (
            nome, servidor, tipo, modelo, localizacao, ip_host, status, 
            total_paginas, filas_relacionadas, detalhes_extra, data_atualizacao
        ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
        ON CONFLICT(nome) DO UPDATE SET
            servidor=excluded.servidor,
            tipo=excluded.tipo,
            modelo=excluded.modelo,
            localizacao=excluded.localizacao,
            ip_host=excluded.ip_host,
            status=excluded.status,
            total_paginas=excluded.total_paginas,
            filas_relacionadas=excluded.filas_relacionadas,
            detalhes_extra=excluded.detalhes_extra,
            data_atualizacao=excluded.data_atualizacao
        """, (
            nome, servidor, tipo, modelo, localizacao, ip_host, status,
            total_paginas, filas_relacionadas, detalhes_extra, now_str
        ))

    conn.commit()
    conn.close()


def get_impressoras_df() -> pd.DataFrame:
    """
    Retorna o DataFrame contendo todas as impressoras/dispositivos do PaperCut cadastrados no banco.
    """
    setup_impressoras_table()
    conn = get_connection()
    df = pd.read_sql_query("SELECT * FROM impressoras ORDER BY nome ASC", conn)
    conn.close()
    return df


def get_central_telefonica_df() -> pd.DataFrame:
    """
    Retorna o DataFrame contendo todos os ramais da Central Telefônica (OXE) cadastrados no banco SQLite.
    Se a tabela estiver vazia, tenta executar o pré-processamento para popular o banco.
    """
    conn = get_connection()
    try:
        df = pd.read_sql_query("SELECT * FROM central_telefonica ORDER BY CAST(ramal AS INTEGER) ASC", conn)
        conn.close()
        if not df.empty:
            return df
    except Exception:
        conn.close()

    # Fallback: executa pré-processamento dos dados brutos se existirem
    try:
        from src.preprocess_oxe import preprocess_oxe
        if preprocess_oxe():
            conn = get_connection()
            df = pd.read_sql_query("SELECT * FROM central_telefonica ORDER BY CAST(ramal AS INTEGER) ASC", conn)
            conn.close()
            return df
    except Exception:
        pass

    # Fallback secundário: leitura direta de 02 - Dados tratados
    from config import OUTPUT_DIR_TRATADOS
    files = sorted(OUTPUT_DIR_TRATADOS.glob("Central_Telefonica_OXE_Tratados_*.xlsx"), key=lambda f: f.stat().st_mtime)
    if files:
        df = pd.read_excel(files[-1], dtype=str)
        df.fillna("", inplace=True)
        return df

    return pd.DataFrame()


def setup_garantia_tables():
    """Cria as tabelas de contratos e chamados de garantia no SQLite se não existirem."""
    conn = get_connection()
    cursor = conn.cursor()
    cursor.execute("""
    CREATE TABLE IF NOT EXISTS garantia_contratos (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        contrato TEXT,
        pu_saj TEXT,
        item TEXT,
        contratacao_por TEXT,
        fornecedor TEXT,
        termo_referencia TEXT,
        termo_recebimento TEXT,
        nota_fiscal TEXT,
        garantia_inicio TEXT,
        garantia_fim TEXT,
        status_garantia TEXT,
        dias_restantes INTEGER,
        link_suporte TEXT,
        data_atualizacao TEXT
    )
    """)
    cursor.execute("""
    CREATE TABLE IF NOT EXISTS garantia_chamados (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        item TEXT,
        status TEXT,
        numero_serie TEXT,
        patrimonio TEXT,
        chamado_mpm TEXT,
        chamado_externo TEXT,
        defeito TEXT,
        solucao TEXT,
        nota_no_chamado TEXT,
        chamado_dmp TEXT,
        data_atualizacao TEXT
    )
    """)
    conn.commit()
    conn.close()


def sync_garantia_from_excel(excel_path: str = None) -> bool:
    """
    Sincroniza os dados da planilha de Garantia (Contratos para garantia.xlsx)
    para as tabelas SQLite 'garantia_contratos' e 'garantia_chamados'.
    """
    setup_garantia_tables()
    if not excel_path:
        from src.config import WARRANTY_FILE_PATH
        excel_path = str(WARRANTY_FILE_PATH)

    p = Path(excel_path)
    if not p.exists():
        return False

    try:
        xls = pd.ExcelFile(excel_path)
        conn = get_connection()
        cursor = conn.cursor()
        now_str = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        today_date = datetime.now().date()

        # 1. Processa aba 'Contratos'
        if 'Contratos' in xls.sheet_names:
            df_raw_c = pd.read_excel(xls, sheet_name='Contratos', header=None, dtype=str)
            header_idx_c = 0
            for r_idx, r in df_raw_c.iterrows():
                r_str = " ".join([str(val).lower() for val in r if pd.notnull(val)])
                if 'item' in r_str or 'contrato' in r_str or 'fornecedor' in r_str or 'saj' in r_str:
                    header_idx_c = r_idx
                    break

            df_contratos = pd.read_excel(xls, sheet_name='Contratos', header=header_idx_c, dtype=str)
            df_contratos.fillna("", inplace=True)
            cursor.execute("DELETE FROM garantia_contratos")

            for _, row in df_contratos.iterrows():
                contrato = ""
                for k in row.index:
                    if 'contrato' in str(k).lower():
                        contrato = str(row[k]).strip()
                        break

                pu_saj = ""
                for k in row.index:
                    if 'saj' in str(k).lower() or 'pu' in str(k).lower():
                        pu_saj = str(row[k]).strip()
                        break

                item = ""
                for k in row.index:
                    if 'item' in str(k).lower() or 'equipamento' in str(k).lower():
                        item = str(row[k]).strip()
                        break

                contratacao_por = ""
                for k in row.index:
                    if 'contrata' in str(k).lower() or 'preg' in str(k).lower() or 'ata' in str(k).lower():
                        contratacao_por = str(row[k]).strip()
                        break

                fornecedor = ""
                for k in row.index:
                    if 'fornecedor' in str(k).lower() or 'empresa' in str(k).lower():
                        fornecedor = str(row[k]).strip()
                        break

                termo_ref = ""
                for k in row.index:
                    if 'refer' in str(k).lower():
                        termo_ref = str(row[k]).strip()
                        break

                termo_rec = ""
                for k in row.index:
                    if 'receb' in str(k).lower() or 'definitiv' in str(k).lower():
                        termo_rec = str(row[k]).strip()
                        break

                nota_fiscal = ""
                for k in row.index:
                    if 'nota' in str(k).lower() or 'fiscal' in str(k).lower():
                        nota_fiscal = str(row[k]).strip()
                        break

                g_inicio = ""
                for k in row.index:
                    if 'começ' in str(k).lower() or 'iníc' in str(k).lower() or 'inicio' in str(k).lower():
                        g_inicio = str(row[k]).strip()
                        break

                g_fim = ""
                for k in row.index:
                    if 'fim' in str(k).lower() or 'térm' in str(k).lower() or 'termino' in str(k).lower():
                        g_fim = str(row[k]).strip()
                        break

                link_sup = ""
                for k in row.index:
                    if 'link' in str(k).lower() or 'chamado' in str(k).lower() or 'site' in str(k).lower() or 'abertura' in str(k).lower():
                        link_sup = str(row[k]).strip()
                        break

                if not item and not contrato and not fornecedor and not pu_saj:
                    continue

                dias_restantes = None
                status_garantia = "Não Informada"
                if g_fim:
                    try:
                        dt_fim_obj = pd.to_datetime(g_fim, dayfirst=True, errors='coerce')
                        if pd.notnull(dt_fim_obj):
                            dias_restantes = (dt_fim_obj.date() - today_date).days
                            if dias_restantes < 0:
                                status_garantia = "Garantia Vencida"
                            elif dias_restantes <= 30:
                                status_garantia = "A Vencer (≤ 30 dias)"
                            else:
                                status_garantia = "Garantia Ativa"
                            g_fim = dt_fim_obj.strftime("%Y-%m-%d")
                    except Exception:
                        pass

                cursor.execute("""
                INSERT INTO garantia_contratos (
                    contrato, pu_saj, item, contratacao_por, fornecedor,
                    termo_referencia, termo_recebimento, nota_fiscal,
                    garantia_inicio, garantia_fim, status_garantia,
                    dias_restantes, link_suporte, data_atualizacao
                ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
                """, (
                    contrato, pu_saj, item, contratacao_por, fornecedor,
                    termo_ref, termo_rec, nota_fiscal,
                    g_inicio, g_fim, status_garantia,
                    dias_restantes, link_sup, now_str
                ))

        # 2. Processa aba 'Chamados'
        if 'Chamados' in xls.sheet_names:
            df_raw_ch = pd.read_excel(xls, sheet_name='Chamados', header=None, dtype=str)
            header_idx_ch = 1
            for r_idx, r in df_raw_ch.iterrows():
                r_str = " ".join([str(val).lower() for val in r if pd.notnull(val)])
                if 'item' in r_str and ('status' in r_str or 'série' in r_str or 'serie' in r_str or 'patrimô' in r_str or 'defeito' in r_str):
                    header_idx_ch = r_idx
                    break

            df_chamados = pd.read_excel(xls, sheet_name='Chamados', header=header_idx_ch, dtype=str)
            df_chamados.fillna("", inplace=True)
            cursor.execute("DELETE FROM garantia_chamados")


            for _, row in df_chamados.iterrows():
                item = ""
                for k in row.index:
                    if 'item' in str(k).lower() or 'equipamento' in str(k).lower():
                        item = str(row[k]).strip()
                        break

                status = ""
                for k in row.index:
                    if 'status' in str(k).lower() or 'situac' in str(k).lower():
                        status = str(row[k]).strip()
                        break

                n_serie = ""
                for k in row.index:
                    if 'série' in str(k).lower() or 'serie' in str(k).lower() or 'serial' in str(k).lower():
                        n_serie = str(row[k]).strip()
                        break

                patrimonio = ""
                for k in row.index:
                    if 'patrimô' in str(k).lower() or 'patrimo' in str(k).lower() or 'tombo' in str(k).lower():
                        patrimonio = str(row[k]).strip()
                        break

                c_mpm = ""
                for k in row.index:
                    if 'mpm' in str(k).lower() or 'otrs' in str(k).lower() or 'citsmart' in str(k).lower():
                        c_mpm = str(row[k]).strip()
                        break

                c_ext = ""
                for k in row.index:
                    if 'externo' in str(k).lower() or 'fornecedor' in str(k).lower():
                        c_ext = str(row[k]).strip()
                        break

                defeito = ""
                for k in row.index:
                    if 'defeito' in str(k).lower() or 'problema' in str(k).lower():
                        defeito = str(row[k]).strip()
                        break

                solucao = ""
                for k in row.index:
                    if 'soluç' in str(k).lower() or 'soluc' in str(k).lower() or 'acao' in str(k).lower():
                        solucao = str(row[k]).strip()
                        break

                nota_chamado = ""
                for k in row.index:
                    if 'nota' in str(k).lower():
                        nota_chamado = str(row[k]).strip()
                        break

                c_dmp = ""
                for k in row.index:
                    if 'dmp' in str(k).lower():
                        c_dmp = str(row[k]).strip()
                        break

                if not item and not patrimonio and not n_serie and not c_mpm and not c_ext:
                    continue

                cursor.execute("""
                INSERT INTO garantia_chamados (
                    item, status, numero_serie, patrimonio, chamado_mpm,
                    chamado_externo, defeito, solucao, nota_no_chamado,
                    chamado_dmp, data_atualizacao
                ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
                """, (
                    item, status, n_serie, patrimonio, c_mpm,
                    c_ext, defeito, solucao, nota_chamado,
                    c_dmp, now_str
                ))

        conn.commit()
        conn.close()
        return True
    except Exception as e:
        print(f"Erro ao sincronizar garantia da planilha: {e}")
        return False


def get_garantia_contratos_df() -> pd.DataFrame:
    """Retorna o DataFrame de Contratos de Garantia do SQLite."""
    setup_garantia_tables()
    conn = get_connection()
    try:
        df = pd.read_sql_query("SELECT * FROM garantia_contratos ORDER BY id ASC", conn)
        conn.close()
        return df
    except Exception:
        conn.close()
        return pd.DataFrame()


def get_garantia_chamados_df() -> pd.DataFrame:
    """Retorna o DataFrame de Chamados de Garantia do SQLite."""
    setup_garantia_tables()
    conn = get_connection()
    try:
        df = pd.read_sql_query("SELECT * FROM garantia_chamados ORDER BY id ASC", conn)
        conn.close()
        return df
    except Exception:
        conn.close()
        return pd.DataFrame()


def setup_eventos_manuais_table():
    """Cria a tabela de eventos manuais no banco de dados SQLite se não existir."""
    conn = get_connection()
    cursor = conn.cursor()
    cursor.execute("""
    CREATE TABLE IF NOT EXISTS eventos_manuais (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        titulo TEXT NOT NULL,
        data_inicio TEXT NOT NULL,
        data_fim TEXT,
        descricao TEXT,
        autor TEXT,
        data_criacao TEXT
    )
    """)
    conn.commit()
    conn.close()


def save_evento_manual(titulo: str, data_inicio: str, data_fim: str = "", descricao: str = "", autor: str = "Bancada STI"):
    """Insere um novo evento manual na tabela eventos_manuais."""
    setup_eventos_manuais_table()
    conn = get_connection()
    cursor = conn.cursor()
    now_str = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    cursor.execute("""
    INSERT INTO eventos_manuais (titulo, data_inicio, data_fim, descricao, autor, data_criacao)
    VALUES (?, ?, ?, ?, ?, ?)
    """, (titulo, data_inicio, data_fim, descricao, autor, now_str))
    conn.commit()
    conn.close()


def get_eventos_manuais() -> pd.DataFrame:
    """Retorna um DataFrame com todos os eventos manuais cadastrados no SQLite."""
    setup_eventos_manuais_table()
    conn = get_connection()
    try:
        df = pd.read_sql_query("SELECT * FROM eventos_manuais ORDER BY id DESC", conn)
        conn.close()
        return df
    except Exception:
        conn.close()
        return pd.DataFrame()


def setup_unidades_tables():
    """Cria as tabelas de unidades unificadas e unidades manuais no SQLite se não existirem."""
    conn = get_connection()
    cursor = conn.cursor()
    cursor.execute("""
    CREATE TABLE IF NOT EXISTS unidades_manuais (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        cidade TEXT,
        tipo TEXT,
        setor TEXT,
        sigla TEXT,
        titular TEXT,
        unidade_predio TEXT,
        telefone TEXT,
        url TEXT,
        data_atualizacao TEXT
    )
    """)
    cursor.execute("""
    CREATE TABLE IF NOT EXISTS unidades (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        cidade TEXT,
        tipo TEXT,
        setor TEXT,
        sigla TEXT,
        titular TEXT,
        unidade_predio TEXT,
        telefone TEXT,
        url TEXT,
        origem TEXT DEFAULT 'web',
        data_atualizacao TEXT
    )
    """)
    conn.commit()
    conn.close()


def setup_unidades_manuais_table():
    """Cria a tabela de unidades manuais no SQLite se não existir."""
    setup_unidades_tables()


def get_unidades_manuais() -> pd.DataFrame:
    """Retorna todas as unidades manuais cadastradas no SQLite."""
    setup_unidades_tables()
    conn = get_connection()
    try:
        df = pd.read_sql_query("SELECT * FROM unidades_manuais ORDER BY id ASC", conn)
        conn.close()
        return df
    except Exception:
        conn.close()
        return pd.DataFrame()


def save_unidades_to_db(df: pd.DataFrame):
    """Salva a lista completa unificada de unidades (Web + Manuais) na tabela 'unidades' do SQLite."""
    setup_unidades_tables()
    if df is None or df.empty:
        return

    conn = get_connection()
    cursor = conn.cursor()
    now_str = datetime.now().strftime("%Y-%m-%d %H:%M:%S")

    cursor.execute("DELETE FROM unidades")

    for _, row in df.iterrows():
        cidade = str(row.get('Cidade', row.get('cidade', ''))).strip()
        tipo = str(row.get('Tipo', row.get('tipo', ''))).strip()
        setor = str(row.get('Setor', row.get('setor', ''))).strip()
        sigla = str(row.get('Sigla', row.get('sigla', ''))).strip()
        titular = str(row.get('Titular', row.get('titular', ''))).strip()
        u_predio = str(row.get('Unidade (Prédio)', row.get('unidade_predio', ''))).strip()
        telefone = str(row.get('Telefone', row.get('telefone', ''))).strip()
        url = str(row.get('URL', row.get('url', ''))).strip()

        if not setor and not cidade:
            continue

        cursor.execute("""
        INSERT INTO unidades (cidade, tipo, setor, sigla, titular, unidade_predio, telefone, url, data_atualizacao)
        VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?)
        """, (cidade, tipo, setor, sigla, titular, u_predio, telefone, url, now_str))

    conn.commit()
    conn.close()


def get_unidades_df() -> pd.DataFrame:
    """Retorna o DataFrame de todas as unidades (Promotorias, Procuradorias e Setores Manuais) salvas no SQLite com identificador de origem."""
    setup_unidades_tables()
    conn = get_connection()
    try:
        query = """
        SELECT 
            u.id,
            u.cidade AS Cidade, 
            u.tipo AS Tipo, 
            u.setor AS Setor, 
            u.sigla AS Sigla, 
            u.titular AS Titular, 
            u.unidade_predio AS 'Unidade (Prédio)', 
            u.telefone AS Telefone, 
            u.url AS URL,
            um.id AS manual_id,
            CASE WHEN um.id IS NOT NULL THEN '📌 Manual' ELSE '🌐 Portal Web' END AS Origem
        FROM unidades u
        LEFT JOIN unidades_manuais um ON u.cidade = um.cidade AND u.tipo = um.tipo AND u.setor = um.setor
        ORDER BY u.cidade ASC, u.setor ASC
        """
        df = pd.read_sql_query(query, conn)
        conn.close()
        return df
    except Exception:
        conn.close()
        return pd.DataFrame()


def add_unidade_manual(cidade: str, tipo: str, setor: str, sigla: str = "", titular: str = "", unidade_predio: str = "", telefone: str = "", url: str = ""):
    """Insere ou atualiza uma unidade manual no SQLite e sincroniza a tabela unificada 'unidades'."""
    setup_unidades_tables()
    conn = get_connection()
    cursor = conn.cursor()
    now_str = datetime.now().strftime("%Y-%m-%d %H:%M:%S")

    # 1. Atualiza tabela unidades_manuais
    cursor.execute("SELECT id FROM unidades_manuais WHERE cidade = ? AND tipo = ? AND setor = ?", (cidade, tipo, setor))
    existing = cursor.fetchone()

    if existing:
        cursor.execute("""
        UPDATE unidades_manuais 
        SET sigla = ?, titular = ?, unidade_predio = ?, telefone = ?, url = ?, data_atualizacao = ?
        WHERE id = ?
        """, (sigla, titular, unidade_predio, telefone, url, now_str, existing[0]))
    else:
        cursor.execute("""
        INSERT INTO unidades_manuais (cidade, tipo, setor, sigla, titular, unidade_predio, telefone, url, data_atualizacao)
        VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?)
        """, (cidade, tipo, setor, sigla, titular, unidade_predio, telefone, url, now_str))

    # 2. Insere/Atualiza também na tabela unificada 'unidades'
    cursor.execute("SELECT id FROM unidades WHERE cidade = ? AND tipo = ? AND setor = ?", (cidade, tipo, setor))
    existing_uni = cursor.fetchone()
    if existing_uni:
        cursor.execute("""
        UPDATE unidades 
        SET sigla = ?, titular = ?, unidade_predio = ?, telefone = ?, url = ?, data_atualizacao = ?
        WHERE id = ?
        """, (sigla, titular, unidade_predio, telefone, url, now_str, existing_uni[0]))
    else:
        cursor.execute("""
        INSERT INTO unidades (cidade, tipo, setor, sigla, titular, unidade_predio, telefone, url, data_atualizacao)
        VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?)
        """, (cidade, tipo, setor, sigla, titular, unidade_predio, telefone, url, now_str))

    conn.commit()
    conn.close()


def update_unidade_manual_by_id(manual_id: int, cidade: str, tipo: str, setor: str, sigla: str = "", titular: str = "", unidade_predio: str = "", telefone: str = "", url: str = ""):
    """Atualiza uma unidade manual pelo seu manual_id no SQLite nas duas tabelas."""
    setup_unidades_tables()
    conn = get_connection()
    cursor = conn.cursor()
    now_str = datetime.now().strftime("%Y-%m-%d %H:%M:%S")

    cursor.execute("SELECT cidade, tipo, setor FROM unidades_manuais WHERE id = ?", (manual_id,))
    old_row = cursor.fetchone()
    if old_row:
        old_cidade, old_tipo, old_setor = old_row[0], old_row[1], old_row[2]
        cursor.execute("""
        UPDATE unidades
        SET cidade = ?, tipo = ?, setor = ?, sigla = ?, titular = ?, unidade_predio = ?, telefone = ?, url = ?, data_atualizacao = ?
        WHERE cidade = ? AND tipo = ? AND setor = ?
        """, (cidade, tipo, setor, sigla, titular, unidade_predio, telefone, url, now_str, old_cidade, old_tipo, old_setor))

    cursor.execute("""
    UPDATE unidades_manuais 
    SET cidade = ?, tipo = ?, setor = ?, sigla = ?, titular = ?, unidade_predio = ?, telefone = ?, url = ?, data_atualizacao = ?
    WHERE id = ?
    """, (cidade, tipo, setor, sigla, titular, unidade_predio, telefone, url, now_str, manual_id))

    conn.commit()
    conn.close()


def delete_unidade_manual(unit_id: int):
    """Deleta uma unidade manual do SQLite (das duas tabelas) pelo ID."""
    setup_unidades_tables()
    conn = get_connection()
    cursor = conn.cursor()

    cursor.execute("SELECT cidade, tipo, setor FROM unidades_manuais WHERE id = ?", (unit_id,))
    row = cursor.fetchone()
    if row:
        cidade, tipo, setor = row[0], row[1], row[2]
        cursor.execute("DELETE FROM unidades WHERE cidade = ? AND tipo = ? AND setor = ?", (cidade, tipo, setor))

    cursor.execute("DELETE FROM unidades_manuais WHERE id = ?", (unit_id,))
    conn.commit()
    conn.close()




