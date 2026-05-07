# -*- coding: utf-8 -*-
# test_app.py
import unittest
import pandas as pd
from unittest.mock import MagicMock

# Importação dos módulos do projeto
from preprocess_chamados import clean_otrs_description, normalize_text
from tag_classifier import clean_text
from unidades_scraper import make_sigla
from manual_entries import set_city_into_unidade
from config import save_df_to_excel_formatted, cleanup_old_files


class TestPreprocessChamados(unittest.TestCase):
    """Testes unitários para o pré-processamento de chamados."""

    def test_normalize_text_accents_and_case(self):
        """Testa remoção de acentos, conversão para minúsculas e espaços extras."""
        self.assertEqual(normalize_text("Órgão de Informática"), "orgao de informatica")
        self.assertEqual(normalize_text("   Espaços   Extras   "), "espacos extras")
        self.assertEqual(normalize_text("Promotoria de Justiça - Três Lagoas!!!"), "promotoria de justica - tres lagoas")

    def test_clean_otrs_description_greetings(self):
        """Testa a remoção de saudações iniciais no OTRS."""
        desc = "Bom dia,\n\nPrecisamos de auxílio no sistema."
        self.assertEqual(clean_otrs_description(desc), "Precisamos de auxílio no sistema.")

        desc_2 = "Olá,\nFavor verificar erro no Excel."
        self.assertEqual(clean_otrs_description(desc_2), "Favor verificar erro no Excel.")

    def test_clean_otrs_description_signatures(self):
        """Testa a remoção de assinaturas e despedidas ("Att.", "Atenciosamente")."""
        desc = "Chamado para troca de toner.\n\nAtenciosamente,\nWellington\nSecretaria Geral"
        self.assertEqual(clean_otrs_description(desc), "Chamado para troca de toner.")

        desc_2 = "Problema com rede de internet.\n\nAtt.,\nSuporte STI"
        self.assertEqual(clean_otrs_description(desc_2), "Problema com rede de internet.")

        desc_3 = "Solicito verificação.\n--\nEnviado do meu e-mail"
        self.assertEqual(clean_otrs_description(desc_3), "Solicito verificação.")

    def test_clean_otrs_description_request_block(self):
        """Testa se a linha de 'Descrição do pedido:' é concatenada horizontalmente."""
        desc = "Descrição do pedido:\nInstalar pacote de software\nno computador novo."
        # A quebra de linha após 'Descrição do pedido:' é mantida de acordo com a lógica original do robô
        self.assertEqual(clean_otrs_description(desc), "Descrição do pedido:\nInstalar pacote de software no computador novo.")

    def test_clean_otrs_description_history_split(self):
        """Testa se o histórico anterior de respostas (#2) é corretamente truncado."""
        desc = "Mensagem nova e atual sobre o problema.\n#2\n21/04/2026 10:00 - Nota anterior do técnico..."
        self.assertEqual(clean_otrs_description(desc), "Mensagem nova e atual sobre o problema.")


class TestTagClassifierNLP(unittest.TestCase):
    """Testes unitários para a limpeza de texto (NLP) do Classificador de Tags."""

    def test_clean_text_basic_cleaning(self):
        """Testa normalização e remoção de acentos do classificador."""
        # 'água' lematiza para 'aguo' e 'gás' lematiza para 'ga' no spaCy pt_core_news_sm
        self.assertEqual(clean_text("ÁGUA mineral com GÁS"), "aguo mineral ga")

    def test_clean_text_it_jargon_preservation(self):
        """Testa se os jargões e abreviações de TI são mapeados e preservados corretamente."""
        # 'ram' vira 'memoriaram', que após lematização do spaCy se torna 'memoriar'.
        cleaned = clean_text("Trocar a memória RAM do computador")
        self.assertIn("memoriar", cleaned)

        # 'ip' vira 'enderecoip', que se mantém após lematização
        cleaned_2 = clean_text("Configurar endereço IP da impressora")
        self.assertIn("enderecoip", cleaned_2)

    def test_clean_text_html_and_links(self):
        """Testa se tags HTML e URLs são removidas da extração de tags."""
        desc = "<p>O link do sistema é https://suporte.mpms.mp.br para acessar</p>"
        # 'sistema' é removido como stop word em português no spaCy, resultando em 'link acessar'
        self.assertEqual(clean_text(desc), "link acessar")


class TestUnidadesScraper(unittest.TestCase):
    """Testes unitários para geração de siglas e processamento de unidades."""

    def test_make_sigla_promotoria_campo_grande(self):
        """Testa a geração de siglas para promotorias em Campo Grande (mapeamento por prédio)."""
        # Prédio Chácara Cachoeira -> PJCHA
        row_cha = pd.Series({
            "Tipo": "Promotoria",
            "Cidade": "Campo Grande",
            "Unidade (Prédio)": "Campo Grande - Chácara Cachoeira",
            "Setor": "10ª Promotoria de Justiça de Campo Grande"
        })
        self.assertEqual(make_sigla(row_cha), "10ª PJCHA")

        # Prédio Rua da Paz -> PJCGR
        row_paz = pd.Series({
            "Tipo": "Promotoria",
            "Cidade": "Campo Grande",
            "Unidade (Prédio)": "Campo Grande - Rua da Paz",
            "Setor": "2ª Promotoria de Justiça de Campo Grande"
        })
        self.assertEqual(make_sigla(row_paz), "2ª PJCGR")

        # Prédio Ricardo Brandão -> PJESP
        row_rb = pd.Series({
            "Tipo": "Promotoria",
            "Cidade": "Campo Grande",
            "Unidade (Prédio)": "Campo Grande - Ricardo Brandão",
            "Setor": "40ª Promotoria de Justiça de Campo Grande"
        })
        self.assertEqual(make_sigla(row_rb), "40ª PJESP")

    def test_make_sigla_promotoria_interior(self):
        """Testa a geração de siglas para promotorias no interior (ex: Três Lagoas)."""
        row_tl = pd.Series({
            "Tipo": "Promotoria",
            "Cidade": "Três Lagoas",
            "Unidade (Prédio)": "Três Lagoas - Sede",
            "Setor": "5ª Promotoria de Justiça de Três Lagoas"
        })
        self.assertEqual(make_sigla(row_tl), "5ª PJ de Três Lagoas")

    def test_make_sigla_procuradoria(self):
        """Testa a geração de siglas para procuradorias."""
        row_proc = pd.Series({
            "Tipo": "Procuradoria",
            "Cidade": "Campo Grande",
            "Unidade (Prédio)": "Campo Grande - PGJ",
            "Setor": "1ª Procuradoria de Justiça Cível"
        })
        self.assertEqual(make_sigla(row_proc), "1ª PJ Cível")


class TestManualEntries(unittest.TestCase):
    """Testes unitários para o carregamento e formatação de entradas manuais."""

    def test_set_city_into_unidade(self):
        """Testa se set_city_into_unidade atualiza corretamente o prédio prefixando a cidade."""
        entries = [
            {
                "Cidade": "Campo Grande",
                "Setor": "STI - Suporte PGJ",
                "Unidade (Prédio)": "PGJ"
            }
        ]
        updated = set_city_into_unidade(entries)
        self.assertEqual(updated[0]["Unidade (Prédio)"], "Campo Grande - PGJ")


class TestConfigHelpers(unittest.TestCase):
    """Testes para os utilitários centralizados em config.py."""

    def test_save_df_to_excel_formatted(self):
        """Testa a gravação e formatação de planilhas Excel."""
        import tempfile
        from pathlib import Path

        df = pd.DataFrame({
            "Chamado#": [1, 2],
            "Descrição": ["Linha 1\nSegunda linha", "Simples"]
        })
        
        with tempfile.TemporaryDirectory() as tmpdir:
            test_file = Path(tmpdir) / "teste_formatacao.xlsx"
            widths = {"Chamado#": 10, "Descrição": 50}
            
            # Testa a execução sem erros
            save_df_to_excel_formatted(
                df, test_file, sheet_name="Teste",
                widths=widths, wrap_cols=["Descrição"], height_col="Descrição"
            )
            self.assertTrue(test_file.exists())


class TestCleanupOldFiles(unittest.TestCase):
    """Testes unitários para a função de limpeza de arquivos antigos."""

    def test_cleanup_keeps_correct_amount_of_newest_files(self):
        import tempfile
        from pathlib import Path
        import time
        import os

        with tempfile.TemporaryDirectory() as tmpdir:
            path_dir = Path(tmpdir)
            
            # Cria 15 arquivos falsos timestampados fictícios
            files = []
            for i in range(15):
                f_path = path_dir / f"test_file_{i}.xlsx"
                f_path.write_text("dummy content")
                files.append(f_path)
                # Define mtimes crescentes para os arquivos (o arquivo 14 será o mais recente)
                mtime = time.time() - (15 - i) * 60
                os.utime(f_path, (mtime, mtime))
            
            # Executa a limpeza para manter apenas os 5 mais recentes
            cleanup_old_files(path_dir, "test_file_*.xlsx", keep_count=5)
            
            # Verifica quais arquivos restaram
            remaining_files = sorted(list(path_dir.glob("test_file_*.xlsx")), key=lambda x: x.stat().st_mtime)
            self.assertEqual(len(remaining_files), 5)
            
            # Os sobreviventes devem ser os mais recentes (test_file_10 a test_file_14)
            expected_names = [f"test_file_{i}.xlsx" for i in range(10, 15)]
            remaining_names = [f.name for f in remaining_files]
            self.assertEqual(remaining_names, expected_names)


class TestRequestsBasedUnidadesScraper(unittest.TestCase):
    """Testes para o unidades_scraper utilizando mocks de requisições HTTP."""

    def test_clean_url_prepends_domain(self):
        from unidades_scraper import clean_url
        self.assertEqual(clean_url("/promotorias/agua-clara"), "https://www.mpms.mp.br/promotorias/agua-clara")
        self.assertEqual(clean_url("https://www.mpms.mp.br/outro"), "https://www.mpms.mp.br/outro")
        self.assertEqual(clean_url(""), "")

    def test_get_cities_parses_html_correctly(self):
        from unittest.mock import patch
        from unidades_scraper import get_cities

        fake_html = """
        <div class="innerpage">
            <a href="/promotorias/agua-clara">Água Clara</a>
            <a href="/promotorias/campo-grande">Campo Grande</a>
            <a href="/outralink">Ignorar esse</a>
        </div>
        """
        mock_response = MagicMock()
        mock_response.text = fake_html
        mock_response.status_code = 200

        with patch('requests.get', return_value=mock_response):
            cities = get_cities()
            self.assertEqual(len(cities), 2)
            self.assertEqual(cities[0][0], "Água Clara")
            self.assertEqual(cities[0][1], "https://www.mpms.mp.br/promotorias/agua-clara")
            self.assertEqual(cities[0][2], "agua-clara")


if __name__ == "__main__":
    unittest.main()
