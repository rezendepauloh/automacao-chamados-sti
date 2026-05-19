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
    
    # Verifica se a coluna 'base' existe (Migração automática)
    cursor.execute("PRAGMA table_info(chamados)")
    columns = [col[1] for col in cursor.fetchall()]
    if 'base' not in columns:
        cursor.execute("ALTER TABLE chamados ADD COLUMN base TEXT")
        
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
        if not cid:
            continue
            
        # Verifica se o chamado já existe e qual o status
        cursor.execute("SELECT status FROM chamados WHERE id = ?", (cid,))
        result = cursor.fetchone()
        
        if result:
            current_status = result[0]
            # Se já estiver fechado, não fazemos nada (respeita a decisão manual ou automática anterior)
            if current_status == 'Fechado':
                continue
                
            # Se existe e está aberto, atualizamos os dados (exceto status)
            cursor.execute("""
            UPDATE chamados SET
                titulo = ?, cidade_predio = ?, unidade = ?, localidade_fisica = ?,
                usuario = ?, id_cliente = ?, descricao = ?, tag = ?, ip_origem = ?,
                data_atualizacao = ?, base = ?
            WHERE id = ?
            """, (
                row.get('Título', ''), row.get('Cidade - Prédio', ''), row.get('Unidade', ''),
                row.get('Localidade física', ''), row.get('Nome do Usuário', ''), row.get('ID do Cliente', ''),
                row.get('Descrição', ''), row.get('TAG', ''), row.get('IP_Origem', ''),
                now, row.get('Base', ''), cid
            ))
        else:
            # Se não existe, insere como Aberto
            cursor.execute("""
            INSERT INTO chamados (
                id, data_criacao, titulo, cidade_predio, unidade, localidade_fisica,
                usuario, id_cliente, descricao, tag, ip_origem, status, data_atualizacao, base
            ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, 'Aberto', ?, ?)
            """, (
                cid, row.get('Data Criação', ''), row.get('Título', ''),
                row.get('Cidade - Prédio', ''), row.get('Unidade', ''), row.get('Localidade física', ''),
                row.get('Nome do Usuário', ''), row.get('ID do Cliente', ''), row.get('Descrição', ''),
                row.get('TAG', ''), row.get('IP_Origem', ''), now, row.get('Base', '')
            ))
            
    conn.commit()
    conn.close()

def close_missing_tickets(active_ids: list):
    """
    Marca como 'Fechado' os chamados que estão no banco como 'Aberto'
    mas não estão na lista de IDs ativos (que vieram da última coleta).
    """
    if not active_ids:
        return
        
    conn = get_connection()
    cursor = conn.cursor()
    
    now = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    
    placeholders = ",".join("?" for _ in active_ids)
    
    cursor.execute(f"""
    UPDATE chamados 
    SET status = 'Fechado', data_atualizacao = ?
    WHERE status = 'Aberto' AND id NOT IN ({placeholders})
    """, [now] + [str(cid) for cid in active_ids])
    
    conn.commit()
    conn.close()

def update_ticket_status(cid: str, new_status: str):
    """Atualiza o status de um chamado específico (usado pelo Streamlit)."""
    conn = get_connection()
    cursor = conn.cursor()
    now = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    
    cursor.execute("""
    UPDATE chamados 
    SET status = ?, data_atualizacao = ?
    WHERE id = ?
    """, (new_status, now, cid))
    
    conn.commit()
    conn.close()
