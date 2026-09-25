# -*- coding: utf-8 -*-
"""
Testes unitários para classificação de tags, locais remotos e regras de texto (src.tag_classifier).
"""

import sys
import unittest
from pathlib import Path

ROOT_DIR = Path(__file__).resolve().parent.parent.parent
if str(ROOT_DIR) not in sys.path:
    sys.path.insert(0, str(ROOT_DIR))

import tests.test_helpers
from src.tag_classifier import clean_text, normalize_for_extraction, detect_and_update_remote_locations
from src.preprocess_chamados import clean_otrs_description, normalize_text

class TestTagClassifier(unittest.TestCase):
    def test_normalize_text_accents(self):
        """Valida normalização de strings com acentos e caixas variadas."""
        self.assertEqual(normalize_text("Órgão de Informática STI"), "orgao de informatica sti")
        self.assertEqual(normalize_text("   Espaços   Múltiplos   "), "espacos multiplos")

    def test_clean_otrs_greetings_and_signatures(self):
        """Valida que saudações e assinaturas são limpas do corpo do chamado."""
        raw = "Bom dia,\n\nComputador da promotoria não está ligando.\n\nAtenciosamente,\nSecretaria"
        cleaned = clean_otrs_description(raw)
        self.assertEqual(cleaned, "Computador da promotoria não está ligando.")

    def test_clean_otrs_description_signatures_and_history(self):
        """Testa remoção de assinaturas e histórico anterior (#2) do OTRS."""
        desc = "Chamado para troca de toner.\n\nAtenciosamente,\nWellington\nSecretaria Geral"
        self.assertEqual(clean_otrs_description(desc), "Chamado para troca de toner.")

        desc_hist = "Mensagem nova e atual sobre o problema.\n#2\n21/04/2026 10:00 - Nota anterior do técnico..."
        self.assertEqual(clean_otrs_description(desc_hist), "Mensagem nova e atual sobre o problema.")

    def test_remote_location_specific_cases(self):
        """Testa casos reais de apoio remoto: Costa Rica -> Ricardo Brandão II e Aquidauana -> Chácara Cachoeira II."""
        import pandas as pd
        
        # Caso Felipe Ferrari (Costa Rica -> Ricardo Brandão II)
        df_felipe = pd.DataFrame({
            "Chamado#": [84990],
            "Nome do Usuário": ["Felipe Ferrari Marcolin"],
            "Cidade - Prédio": ["Costa Rica - Sede"],
            "Unidade": ["Costa Rica - Sede"],
            "Descrição": ["Felipe Ferrari Marcolin está lotado em Costa Rica - Sede, mas Angela está trabalhando na Sala do Suporte de Apoio Remoto na Ricardo Brandão - Unidade II e solicitou apoio na instalação do software."]
        })
        res_felipe = detect_and_update_remote_locations(df_felipe)
        self.assertEqual(res_felipe.loc[0, "Localidade física"] if hasattr(res_felipe, "loc") else res_felipe.data.get("Localidade física", [""])[0], "Campo Grande - Ricardo Brandão II")

        # Caso Nadson Borges (Aquidauana -> Chácara Cachoeira II)
        df_nadson = pd.DataFrame({
            "Chamado#": [84995],
            "Nome do Usuário": ["Nadson Matheus Borges"],
            "Cidade - Prédio": ["Aquidauana - Sede"],
            "Unidade": ["Aquidauana - Sede"],
            "Descrição": ["Considerando que passei a desempenhar minhas funções em teletrabalho a partir de hoje, na sala do apoio remoto, na Unidade Chácara Cachoeira II, peço revisão de rede."]
        })
        res_nadson = detect_and_update_remote_locations(df_nadson)
        self.assertEqual(res_nadson.loc[0, "Localidade física"] if hasattr(res_nadson, "loc") else res_nadson.data.get("Localidade física", [""])[0], "Campo Grande - Chácara Cachoeira II")

    def test_location_sanitization_nan_and_ad_not_found(self):
        """Valida que valores 'nan' ou 'Não encontrado no AD' não geram 'nan - Não encontrado no AD'."""
        import pandas as pd

        df_nan = pd.DataFrame({
            "Chamado#": [113099, 113100, 113101],
            "Nome do Usuário": ["User A", "User B", "User C"],
            "Cidade - Prédio": [float('nan'), "Campo Grande - PGJ", None],
            "Unidade": ["Não encontrado no AD", "Não encontrada no AD", "STI Governança"],
            "Descrição": ["Sem menção de localidade", "Outro chamado comum", "Chamado padrão"]
        })
        res = detect_and_update_remote_locations(df_nan)
        self.assertEqual(res.loc[0, "Localidade física"], "Não identificada")
        self.assertEqual(res.loc[1, "Localidade física"], "Campo Grande - PGJ")
        self.assertEqual(res.loc[2, "Localidade física"], "STI Governança")

    def test_interior_cities_location_without_unit_suffix(self):
        """Valida que chamados do interior ficam com a localidade física limpa (apenas nome da cidade)."""
        import pandas as pd

        df_interior = pd.DataFrame({
            "Chamado#": [46463328, 46463162, 46463272],
            "Nome do Usuário": ["User 1", "User 2", "User 3"],
            "Cidade - Prédio": ["Paranaíba - Sede", "Terenos - Sede", "Cassilândia - Sede"],
            "Unidade": ["1ª PJ de Paranaíba", "1ª PJ de Terenos", "1ª PJ de Cassilândia"],
            "Descrição": ["Chamado comum", "Chamado padrão", "Sem menção"]
        })
        res = detect_and_update_remote_locations(df_interior)
        self.assertEqual(res.loc[0, "Localidade física"], "Paranaíba")
        self.assertEqual(res.loc[1, "Localidade física"], "Terenos")
        self.assertEqual(res.loc[2, "Localidade física"], "Cassilândia")

    def test_generate_synthetic_title_with_tag_and_boilerplate(self):
        """Valida que saudações, preâmbulos institucionais e HTML são removidos e a TAG predita é prefixada."""
        from src.tag_classifier import generate_synthetic_title

        desc1 = "Bom dia Prezados! Por determinação do Promotor de Justiça, solicito a instalação da impressora PRT-5394 no setor."
        title1 = generate_synthetic_title("IMPRESSORA", desc1)
        self.assertEqual(title1, "[IMPRESSORA] Instalação da impressora PRT-5394 no setor")

        desc2 = "<div>Boa tarde,<br>Gostaria de solicitar a troca de tonner da máquina.</div>"
        title2 = generate_synthetic_title("IMPRESSORA", desc2)
        self.assertEqual(title2, "[IMPRESSORA] Troca de tonner da máquina")

        desc3 = "Microcomputador não liga após queda de energia na comarca."
        title3 = generate_synthetic_title("HARDWARE", desc3)
        self.assertEqual(title3, "[HARDWARE] Microcomputador não liga após queda de energia na comarca")

    def test_generate_missing_titles_preserves_existing(self):
        """Valida que chamados que já possuem título (ex: OTRS) não são modificados."""
        import pandas as pd
        from src.tag_classifier import generate_missing_titles

        df = pd.DataFrame({
            "Chamado#": [1, 2],
            "TAG": ["REDE", "IMPRESSORA"],
            "Título": ["Título Original OTRS", ""],
            "Descrição": ["Problema de rede", "Por meio deste solicito configuração de impressora laser."]
        })
        res = generate_missing_titles(df)
        self.assertEqual(res.loc[0, "Título"], "Título Original OTRS")
        self.assertEqual(res.loc[1, "Título"], "[IMPRESSORA] Configuração de impressora laser")

if __name__ == "__main__":
    unittest.main()
