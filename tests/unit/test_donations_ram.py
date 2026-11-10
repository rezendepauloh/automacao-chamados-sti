import os
import sys
import unittest
import pandas as pd
from io import BytesIO
from pathlib import Path

ROOT_DIR = Path(__file__).resolve().parent.parent.parent
if str(ROOT_DIR) not in sys.path:
    sys.path.insert(0, str(ROOT_DIR))

import tests.test_helpers
from src.database.donations_db import (
    setup_donations_table,
    sync_donations_from_excel,
    get_donations_data
)

class TestDonationsRAM(unittest.TestCase):
    def setUp(self):
        setup_donations_table()

    def test_sync_donations_with_ram_column(self):
        """Valida que a coluna 'Memória RAM' é lida do Excel e persistida corretamente."""
        data = {
            "Patrimônio": ["098765", "098766"],
            "Modelo": ["HP EliteDesk 800 G5", "Dell OptiPlex 7050"],
            "Serial Number PC": ["BRG1111", "BRG2222"],
            "Equipamento": ["Computador", "Computador"],
            "Doação ou Baixa": ["Doação", "Redistribuição"],
            "Data da doação": ["2026-11-10", "2026-11-10"],
            "Tem chamado?": ["#46270001", "#46270002"],
            "SSD": ["256 GB SSD", "500 GB HD"],
            "Motivo baixa": ["", "Troca preventiva"],
            "Memória RAM": ["16 GB (2x 8GB)", "8 GB (1x 8GB)"]
        }
        df_excel = pd.DataFrame(data)

        buf = BytesIO()
        with pd.ExcelWriter(buf, engine="openpyxl") as writer:
            df_excel.to_excel(writer, sheet_name="Equipamentos doados", index=False)
        buf.seek(0)

        # Executa sincronização
        sync_donations_from_excel(buf)

        # Recupera dados persistidos
        df_saved = get_donations_data()
        self.assertFalse(df_saved.empty)
        self.assertIn("memoria_ram", df_saved.columns)

        row1 = df_saved[df_saved["patrimonio"].astype(str).str.contains("98765")].iloc[0]
        self.assertEqual(row1["memoria_ram"], "16 GB (2x 8GB)")
        self.assertEqual(row1["ssd"], "256 GB SSD")

        row2 = df_saved[df_saved["patrimonio"].astype(str).str.contains("98766")].iloc[0]
        self.assertEqual(row2["memoria_ram"], "8 GB (1x 8GB)")
        self.assertEqual(row2["ssd"], "500 GB HD")

if __name__ == "__main__":
    unittest.main()
