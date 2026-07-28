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

if __name__ == "__main__":
    limpar_chamados_automaticos()
