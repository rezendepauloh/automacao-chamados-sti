import sqlite3
from pathlib import Path

DB_PATH = Path("chamados.db")

def limpar_tabelas_plantoes():
    """Remove todos os registros das tabelas plantoes_matutino e plantoes_semanal do SQLite."""
    if not DB_PATH.exists():
        print(f"Erro: Banco de dados '{DB_PATH}' não encontrado.")
        return

    conn = sqlite3.connect(DB_PATH)
    cursor = conn.cursor()

    cursor.execute("DELETE FROM plantoes_matutino")
    matutino_removidos = cursor.rowcount

    cursor.execute("DELETE FROM plantoes_semanal")
    semanal_removidos = cursor.rowcount

    conn.commit()
    conn.close()

    print("✅ Limpeza de plantões concluída!")
    print(f"   - Registros de Plantão Matutino removidos: {matutino_removidos}")
    print(f"   - Registros de Plantão Semanal removidos: {semanal_removidos}")

if __name__ == "__main__":
    limpar_tabelas_plantoes()
