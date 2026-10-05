# -*- coding: utf-8 -*-
"""
Módulo de Persistência Relacional — Gerenciador de Senhas / Cofre da Bancada
Gerencia o armazenamento criptografado (Fernet/AES) das credenciais de sistemas,
redes, switches, servidores, bancos de dados e ferramentas de suporte.
"""

import sqlite3
from datetime import datetime
from typing import List, Dict, Any, Optional
import pandas as pd

from src.database.connection import get_connection
from src.crypto_utils import encrypt_value, decrypt_value


CATEGORIAS_PADRAO = [
    "Sistemas Web",
    "Rede & Switches",
    "Servidores",
    "Bancos de Dados",
    "Ferramentas Internas",
    "Outros"
]


def setup_senhas_table():
    """
    Cria a tabela senhas_cofre se ela não existir.
    Garante índices para consultas ágeis por categoria e título.
    """
    conn = get_connection()
    cursor = conn.cursor()
    cursor.execute("""
    CREATE TABLE IF NOT EXISTS senhas_cofre (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        titulo TEXT NOT NULL,
        categoria TEXT NOT NULL,
        url_sistema TEXT DEFAULT '',
        usuario TEXT NOT NULL,
        senha_cifrada TEXT NOT NULL,
        observacoes TEXT DEFAULT '',
        data_criacao TIMESTAMP DEFAULT CURRENT_TIMESTAMP,
        data_atualizacao TIMESTAMP DEFAULT CURRENT_TIMESTAMP
    );
    """)
    cursor.execute("CREATE INDEX IF NOT EXISTS idx_senhas_categoria ON senhas_cofre(categoria);")
    cursor.execute("CREATE INDEX IF NOT EXISTS idx_senhas_titulo ON senhas_cofre(titulo);")
    conn.commit()
    conn.close()


def salvar_senha(
    titulo: str,
    categoria: str,
    usuario: str,
    senha_plana: str,
    url_sistema: str = "",
    observacoes: str = ""
) -> int:
    """
    Criptografa e insere uma nova credencial no cofre.
    Retorna o ID do registro inserido.
    """
    if not titulo or not str(titulo).strip():
        raise ValueError("O título da credencial é obrigatório.")
    if not usuario or not str(usuario).strip():
        raise ValueError("O usuário/login é obrigatório.")
    if not senha_plana or not str(senha_plana).strip():
        raise ValueError("A senha é obrigatória.")

    setup_senhas_table()

    senha_cifrada = encrypt_value(str(senha_plana).strip())
    now_iso = datetime.now().strftime("%Y-%m-%d %H:%M:%S")

    conn = get_connection()
    cursor = conn.cursor()
    cursor.execute("""
        INSERT INTO senhas_cofre (titulo, categoria, url_sistema, usuario, senha_cifrada, observacoes, data_criacao, data_atualizacao)
        VALUES (?, ?, ?, ?, ?, ?, ?, ?)
    """, (
        titulo.strip(),
        categoria.strip() or "Outros",
        url_sistema.strip() if url_sistema else "",
        usuario.strip(),
        senha_cifrada,
        observacoes.strip() if observacoes else "",
        now_iso,
        now_iso
    ))
    new_id = cursor.lastrowid
    conn.commit()
    conn.close()
    return new_id


def listar_senhas(categoria: Optional[str] = None, busca: Optional[str] = None) -> List[Dict[str, Any]]:
    """
    Lista os registros do cofre de senhas de forma segura (sem decifrar a senha no retorno).
    As senhas retornam mascaradas por padrão ('••••••••').
    """
    setup_senhas_table()
    conn = get_connection()
    conn.row_factory = sqlite3.Row
    cursor = conn.cursor()

    query = "SELECT id, titulo, categoria, url_sistema, usuario, observacoes, data_criacao, data_atualizacao FROM senhas_cofre WHERE 1=1"
    params = []

    if categoria and categoria != "Todas" and categoria.strip():
        query += " AND categoria = ?"
        params.append(categoria.strip())

    if busca and busca.strip():
        term = f"%{busca.strip()}%"
        query += " AND (titulo LIKE ? OR usuario LIKE ? OR observacoes LIKE ? OR url_sistema LIKE ?)"
        params.extend([term, term, term, term])

    query += " ORDER BY titulo COLLATE NOCASE ASC;"

    cursor.execute(query, params)
    rows = cursor.fetchall()
    conn.close()

    resultado = []
    for r in rows:
        resultado.append({
            "id": r["id"],
            "titulo": r["titulo"],
            "categoria": r["categoria"],
            "url_sistema": r["url_sistema"] or "",
            "usuario": r["usuario"],
            "senha_mascarada": "••••••••",
            "observacoes": r["observacoes"] or "",
            "data_criacao": r["data_criacao"],
            "data_atualizacao": r["data_atualizacao"]
        })
    return resultado


def get_senhas_df(categoria: Optional[str] = None, busca: Optional[str] = None) -> pd.DataFrame:
    """
    Retorna DataFrame de credenciais com senha mascarada e metadados para exibição.
    """
    lista = listar_senhas(categoria=categoria, busca=busca)
    if not lista:
        return pd.DataFrame(columns=[
            "id", "titulo", "categoria", "url_sistema", "usuario",
            "senha_mascarada", "observacoes", "data_criacao", "data_atualizacao"
        ])
    return pd.DataFrame(lista)


def obter_senha_decifrada(senha_id: int) -> Optional[str]:
    """
    Recupera e decriptografa a senha plana para o registro especificado.
    Retorna None se o registro não for encontrado.
    """
    setup_senhas_table()
    conn = get_connection()
    cursor = conn.cursor()
    cursor.execute("SELECT senha_cifrada FROM senhas_cofre WHERE id = ?", (senha_id,))
    row = cursor.fetchone()
    conn.close()

    if not row or not row[0]:
        return None

    return decrypt_value(row[0])


def obter_credencial_por_id(senha_id: int) -> Optional[Dict[str, Any]]:
    """
    Retorna o dicionário completo do registro (com senha decifrada) para fins de edição.
    """
    setup_senhas_table()
    conn = get_connection()
    conn.row_factory = sqlite3.Row
    cursor = conn.cursor()
    cursor.execute("SELECT * FROM senhas_cofre WHERE id = ?", (senha_id,))
    row = cursor.fetchone()
    conn.close()

    if not row:
        return None

    return {
        "id": row["id"],
        "titulo": row["titulo"],
        "categoria": row["categoria"],
        "url_sistema": row["url_sistema"] or "",
        "usuario": row["usuario"],
        "senha_plana": decrypt_value(row["senha_cifrada"]),
        "observacoes": row["observacoes"] or "",
        "data_criacao": row["data_criacao"],
        "data_atualizacao": row["data_atualizacao"]
    }


def atualizar_senha(
    senha_id: int,
    titulo: str,
    categoria: str,
    usuario: str,
    senha_plana: Optional[str] = None,
    url_sistema: str = "",
    observacoes: str = ""
) -> bool:
    """
    Atualiza uma credencial existente. Se senha_plana for None ou string vazia,
    mantém a senha criptografada anterior intacta.
    """
    if not titulo or not str(titulo).strip():
        raise ValueError("O título da credencial é obrigatório.")
    if not usuario or not str(usuario).strip():
        raise ValueError("O usuário/login é obrigatório.")

    setup_senhas_table()
    now_iso = datetime.now().strftime("%Y-%m-%d %H:%M:%S")

    conn = get_connection()
    cursor = conn.cursor()

    if senha_plana and str(senha_plana).strip():
        nova_cifrada = encrypt_value(str(senha_plana).strip())
        cursor.execute("""
            UPDATE senhas_cofre
            SET titulo = ?, categoria = ?, url_sistema = ?, usuario = ?,
                senha_cifrada = ?, observacoes = ?, data_atualizacao = ?
            WHERE id = ?
        """, (
            titulo.strip(),
            categoria.strip() or "Outros",
            url_sistema.strip() if url_sistema else "",
            usuario.strip(),
            nova_cifrada,
            observacoes.strip() if observacoes else "",
            now_iso,
            senha_id
        ))
    else:
        cursor.execute("""
            UPDATE senhas_cofre
            SET titulo = ?, categoria = ?, url_sistema = ?, usuario = ?,
                observacoes = ?, data_atualizacao = ?
            WHERE id = ?
        """, (
            titulo.strip(),
            categoria.strip() or "Outros",
            url_sistema.strip() if url_sistema else "",
            usuario.strip(),
            observacoes.strip() if observacoes else "",
            now_iso,
            senha_id
        ))

    affected = cursor.rowcount
    conn.commit()
    conn.close()
    return affected > 0


def excluir_senha(senha_id: int) -> bool:
    """Exclui permanentemente uma credencial do cofre."""
    setup_senhas_table()
    conn = get_connection()
    cursor = conn.cursor()
    cursor.execute("DELETE FROM senhas_cofre WHERE id = ?", (senha_id,))
    affected = cursor.rowcount
    conn.commit()
    conn.close()
    return affected > 0


def get_senhas_stats() -> Dict[str, Any]:
    """Retorna métricas para os KPI Cards do cofre de senhas."""
    setup_senhas_table()
    conn = get_connection()
    cursor = conn.cursor()

    cursor.execute("SELECT COUNT(*) FROM senhas_cofre")
    total_senhas = cursor.fetchone()[0]

    cursor.execute("SELECT COUNT(DISTINCT categoria) FROM senhas_cofre WHERE categoria IS NOT NULL AND categoria != ''")
    total_categorias = cursor.fetchone()[0]

    cursor.execute("SELECT COUNT(*) FROM senhas_cofre WHERE url_sistema IS NOT NULL AND url_sistema != ''")
    com_link = cursor.fetchone()[0]

    conn.close()

    return {
        "total_senhas": total_senhas,
        "total_categorias": total_categorias,
        "com_link": com_link
    }
