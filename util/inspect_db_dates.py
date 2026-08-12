import sqlite3
import pandas as pd
from pathlib import Path

ROOT_DIR = Path(__file__).resolve().parent.parent
DB_PATH = ROOT_DIR / "chamados.db"

def inspect_dates():
    if not DB_PATH.exists():
        print(f"❌ Banco não encontrado em: {DB_PATH}")
        return

    conn = sqlite3.connect(DB_PATH)
    cursor = conn.cursor()
    
    print("--- 🔍 20 PRIMEIRAS DATAS NO BANCO DE DADOS ---")
    cursor.execute("SELECT id, data_criacao, base FROM chamados LIMIT 20")
    for row in cursor.fetchall():
        print(f"ID: {row[0]:<12} | Data: {str(row[1]):<22} | Base: {row[2]}")
        
    print("\n--- 🔍 20 ÚLTIMAS DATAS NO BANCO DE DADOS ---")
    cursor.execute("SELECT id, data_criacao, base FROM chamados ORDER BY rowid DESC LIMIT 20")
    for row in cursor.fetchall():
        print(f"ID: {row[0]:<12} | Data: {str(row[1]):<22} | Base: {row[2]}")

    print("\n--- 📊 FORMATOS DE DATA ENCONTRADOS ---")
    cursor.execute("SELECT data_criacao FROM chamados")
    all_dates = [r[0] for r in cursor.fetchall() if r[0]]
    
    iso_count = sum(1 for d in all_dates if str(d).startswith("2026-") or str(d).startswith("2025-"))
    br_count = sum(1 for d in all_dates if "/" in str(d))
    other_count = len(all_dates) - iso_count - br_count
    
    print(f"Total registros: {len(all_dates)}")
    print(f"No formato ISO (YYYY-MM-DD...): {iso_count}")
    print(f"No formato BR (DD/MM/YYYY...): {br_count}")
    print(f"Outros/Inválidos: {other_count}")

    conn.close()

if __name__ == "__main__":
    inspect_dates()
