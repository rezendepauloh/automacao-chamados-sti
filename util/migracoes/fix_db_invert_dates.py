import sys
from pathlib import Path
import sqlite3
import re
from datetime import datetime

ROOT_DIR = Path(__file__).resolve().parent.parent
DB_PATH = ROOT_DIR / "chamados.db"
TODAY = datetime.now()

def fix_invert_dates():
    """Inverte cirurgicamente o mês e o dia nas strings ISO (YYYY-MM-DD) do banco SQLite para datas no futuro."""
    if not DB_PATH.exists():
        print(f"❌ Banco de dados não encontrado em: {DB_PATH}")
        return

    print(f"🔄 Conectando ao banco de dados: {DB_PATH.name}...")
    conn = sqlite3.connect(DB_PATH)
    cursor = conn.cursor()

    try:
        cursor.execute("SELECT id, data_criacao FROM chamados")
        rows = cursor.fetchall()

        total_analisados = len(rows)
        total_corrigidos = 0
        total_mantidos = 0
        total_invalidos = 0

        print(f"📊 Total de registros encontrados na tabela 'chamados': {total_analisados}")
        print("⏳ Iniciando inversão de Mês/Dia para datas no futuro...")

        updates = []
        for cid, raw_date in rows:
            if not raw_date or str(raw_date).strip().lower() in ["none", "nan", "null", ""]:
                total_invalidos += 1
                continue

            s_date = str(raw_date).strip()

            # Padrão ISO: YYYY-MM-DD HH:MM:SS
            match = re.match(r"^(\d{4})-(\d{2})-(\d{2})(.*)$", s_date)
            if match:
                year, month, day, time_part = match.groups()

                try:
                    dt_current = datetime.strptime(f"{year}-{month}-{day}", "%Y-%m-%d")
                    # Só inverte se a data atual registrada estiver no futuro!
                    if dt_current > TODAY:
                        new_month = day
                        new_day = month
                        test_str = f"{year}-{new_month}-{new_day}"
                        test_dt = datetime.strptime(test_str, "%Y-%m-%d")
                        
                        if test_dt <= TODAY:
                            fixed_date_str = f"{year}-{new_month}-{new_day}{time_part}"
                            updates.append((fixed_date_str, str(cid)))
                            total_corrigidos += 1
                        else:
                            total_mantidos += 1
                    else:
                        total_mantidos += 1
                except ValueError:
                    total_mantidos += 1
            else:
                total_invalidos += 1

        if updates:
            print(f"💾 Atualizando {len(updates)} registros no banco de dados...")
            cursor.executemany("""
                UPDATE chamados 
                SET data_criacao = ? 
                WHERE id = ?
            """, updates)
            conn.commit()
            print("✅ Inversão concluída e salva com sucesso!")
        else:
            print("ℹ️ Nenhuma data no futuro precisou ser alterada.")

        print("\n--- RESUMO DA CORREÇÃO ---")
        print(f"✔️ Registros analisados: {total_analisados}")
        print(f"✏️ Registros invertidos/corrigidos: {total_corrigidos}")
        print(f"🆗 Registros mantidos: {total_mantidos}")
        print(f"⚠️ Registros nulos/inválidos: {total_invalidos}")

    except Exception as e:
        conn.rollback()
        print(f"❌ Erro ao processar inversão de datas: {e}")
    finally:
        conn.close()

if __name__ == "__main__":
    fix_invert_dates()
