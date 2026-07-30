import sqlite3
from pathlib import Path

DB_PATH = Path("chamados.db")

def limpar_chamados_automaticos():
    if not DB_PATH.exists():
        print(f"Erro: Banco de dados '{DB_PATH}' não encontrado.")
        return

    conn = sqlite3.connect(DB_PATH)
    cursor = conn.cursor()

    usuarios_remocao = ('Monitoramento Adm Mpms', 'Adm Ticket Por Email')

    # Remove os comentários associados aos chamados
    cursor.execute("""
    DELETE FROM comentarios 
    WHERE chamado_id IN (
        SELECT id FROM chamados 
        WHERE usuario IN (?, ?)
    )
    """, usuarios_remocao)

    comentarios_removidos = cursor.rowcount

    # Remove os chamados
    cursor.execute("""
    DELETE FROM chamados 
    WHERE usuario IN (?, ?)
    """, usuarios_remocao)

    chamados_removidos = cursor.rowcount

    conn.commit()
    conn.close()

    print(f"✅ Limpeza concluída!")
    print(f"   - Chamados removidos: {chamados_removidos}")
    print(f"   - Comentários removidos: {comentarios_removidos}")

def limpar_impressoras_invalidas():
    if not DB_PATH.exists():
        print(f"Erro: Banco de dados '{DB_PATH}' não encontrado.")
        return

    conn = sqlite3.connect(DB_PATH)
    cursor = conn.cursor()

    # Apaga registros com ';' no nome ou que são cabeçalhos de importações antigas
    cursor.execute("""
    DELETE FROM impressoras 
    WHERE nome LIKE '%;%' 
       OR nome LIKE '%Dispositivo;%'
       OR nome LIKE '%Tipo de Dispositivo%'
       OR nome LIKE '%atividade%'
       OR LENGTH(TRIM(nome)) < 2
    """)

    impressoras_removidas = cursor.rowcount
    conn.commit()
    conn.close()

    print(f"✅ Limpeza de impressoras concluída!")
    print(f"   - Registros de impressoras inválidos removidos: {impressoras_removidas}")


if __name__ == "__main__":
    limpar_chamados_automaticos()
    limpar_impressoras_invalidas()

