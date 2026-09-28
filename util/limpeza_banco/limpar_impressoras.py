import sqlite3
from pathlib import Path

DB_PATH = Path("chamados.db")

def limpar_tabela_impressoras():
    """Remove todos os registros da tabela impressoras no banco SQLite chamados.db."""
    if not DB_PATH.exists():
        print(f"Erro: Banco de dados '{DB_PATH}' não encontrado.")
        return

    conn = sqlite3.connect(DB_PATH)
    cursor = conn.cursor()

    cursor.execute("DELETE FROM impressoras")
    removidos = cursor.rowcount

    conn.commit()
    conn.close()

    print("✅ Limpeza da tabela 'impressoras' concluída com sucesso!")
    print(f"   - Registros de Impressoras/Dispositivos removidos: {removidos}")


if __name__ == "__main__":
    limpar_tabela_impressoras()
