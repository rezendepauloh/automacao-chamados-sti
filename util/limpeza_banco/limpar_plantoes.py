import sys
import sqlite3
from pathlib import Path

BASE_DIR = Path(__file__).resolve().parent.parent.parent
if str(BASE_DIR) not in sys.path:
    sys.path.insert(0, str(BASE_DIR))

from src.database.connection import get_connection

def limpar_tabelas_plantoes():
    """Remove todos os registros das tabelas plantoes_matutino e plantoes_semanal do banco de dados."""
    conn = get_connection()
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
