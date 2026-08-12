import sys
from pathlib import Path
import sqlite3
import re
from datetime import datetime

ROOT_DIR = Path(__file__).resolve().parent.parent
DB_PATH = ROOT_DIR / "chamados.db"
TODAY = datetime.now()

def fix_future_and_invalid_dates():
    """
    Corrige no banco SQLite:
    1. Força a inversão de mês/dia para qualquer data em que o mês seja > 8 em 2026 (pois estamos em Agosto/2026).
    2. Converte qualquer data BR (DD/MM/YYYY) remanescente para ISO.
    """
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
        print(f"⏳ Analisando e corrigindo datas no futuro (Hoje é {TODAY.strftime('%Y-%m-%d')})...")

        updates = []
        for cid, raw_date in rows:
            if not raw_date or str(raw_date).strip().lower() in ["none", "nan", "null", ""]:
                total_invalidos += 1
                continue

            s_date = str(raw_date).strip()
            fixed_date_str = None

            # Caso 1: Formato ISO YYYY-MM-DD HH:MM:SS
            match_iso = re.match(r"^(\d{4})-(\d{2})-(\d{2})(.*)$", s_date)
            if match_iso:
                year, month, day, time_part = match_iso.groups()
                
                # Se ano for 2026 e o mês for maior que 8 (Agosto), está 100% invertido!
                int_year = int(year)
                int_month = int(month)
                int_day = int(day)

                if int_year == 2026 and int_month > TODAY.month:
                    # Inverte mês e dia
                    new_month = f"{int_day:02d}"
                    new_day = f"{int_month:02d}"
                    
                    try:
                        # Valida se a data invertida é uma data real (ex: 2026-08-12)
                        test_str = f"{year}-{new_month}-{new_day}"
                        datetime.strptime(test_str, "%Y-%m-%d")
                        fixed_date_str = f"{year}-{new_month}-{new_day}{time_part}"
                    except ValueError:
                        pass

            # Caso 2: Formato BR DD/MM/YYYY HH:MM:SS
            match_br = re.match(r"^(\d{2})/(\d{2})/(\d{4})(.*)$", s_date)
            if match_br and not fixed_date_str:
                day, month, year, time_part = match_br.groups()
                fixed_date_str = f"{year}-{month}-{day}{time_part}"

            if fixed_date_str and fixed_date_str != s_date:
                updates.append((fixed_date_str, str(cid)))
                total_corrigidos += 1
            else:
                total_mantidos += 1

        if updates:
            print(f"💾 Atualizando {len(updates)} registros com datas corrigidas no banco de dados...")
            cursor.executemany("""
                UPDATE chamados 
                SET data_criacao = ? 
                WHERE id = ?
            """, updates)
            conn.commit()
            print("✅ Atualização concluída e salva com sucesso!")
        else:
            print("ℹ️ Nenhuma data no futuro ou inválida precisou ser alterada.")

        print("\n--- RESUMO DA CORREÇÃO ---")
        print(f"✔️ Registros analisados: {total_analisados}")
        print(f"✏️ Registros corrigidos (invertidos do futuro para o passado): {total_corrigidos}")
        print(f"🆗 Registros mantidos: {total_mantidos}")
        print(f"⚠️ Registros nulos/inválidos: {total_invalidos}")

    except Exception as e:
        conn.rollback()
        print(f"❌ Erro ao processar correção de datas: {e}")
    finally:
        conn.close()

if __name__ == "__main__":
    fix_future_and_invalid_dates()
