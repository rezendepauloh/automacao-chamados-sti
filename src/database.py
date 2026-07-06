import sqlite3
import pandas as pd
from pathlib import Path
from datetime import datetime

DB_PATH = Path("chamados.db")

def get_connection():
    """Retorna uma conexão com o banco de dados SQLite."""
    return sqlite3.connect(DB_PATH)

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
                "ip_origem = ?", "data_atualizacao = ?", "base = ?", "link = ?", "hostname = ?"
            ])
            update_params.extend([
                row.get('Título', ''), row.get('Nome do Usuário', ''), row.get('ID do Cliente', ''), row.get('Descrição', ''),
                row.get('IP_Origem', ''), now, row.get('Base', ''), row.get('Link', ''), row.get('Hostname', '')
            ])
            
            # Se a tag NÃO for manual, atualiza
            if tag_manual != 1:
                update_fields.append("tag = ?")
                update_params.append(row.get('TAG', ''))
                
            # Se a localidade/prédio/unidade NÃO for manual, atualiza
            if dados_manuais != 1:
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



