import sys
import sqlite3
from pathlib import Path

BASE_DIR = Path(__file__).resolve().parent.parent.parent
if str(BASE_DIR) not in sys.path:
    sys.path.insert(0, str(BASE_DIR))

from src.database.connection import get_connection

def limpar_tabela_impressoras():
    """Remove todos os registros da tabela impressoras no banco de dados."""
    conn = get_connection()
    cursor = conn.cursor()

    cursor.execute("DELETE FROM impressoras")
    removidos = cursor.rowcount

    conn.commit()
    conn.close()

    print("✅ Limpeza da tabela 'impressoras' concluída com sucesso!")
    print(f"   - Registros de Impressoras/Dispositivos removidos: {removidos}")


if __name__ == "__main__":
    limpar_tabela_impressoras()
