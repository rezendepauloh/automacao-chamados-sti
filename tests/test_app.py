# -*- coding: utf-8 -*-
# test_app.py
import sys
from pathlib import Path

# Adiciona a raiz do projeto e a pasta src ao sys.path
root_dir = Path(__file__).parent.parent
sys.path.insert(0, str(root_dir))
sys.path.insert(0, str(root_dir / "src"))

import unittest
import pandas as pd
from unittest.mock import MagicMock

# Importação dos módulos do projeto
from src.preprocess_chamados import clean_otrs_description, normalize_text
from src.tag_classifier import clean_text, normalize_for_extraction, detect_and_update_remote_locations
from src.scrapers.unidades_scraper import make_sigla
from src.manual_entries import set_city_into_unidade
from src.config import save_df_to_excel_formatted, cleanup_old_files


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
        self.assertEqual(clean_otrs_description(desc), 'Instalar pacote de software\nno computador novo.')


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
        from src.scrapers.unidades_scraper import clean_url
        self.assertEqual(clean_url("/promotorias/agua-clara"), "https://www.mpms.mp.br/promotorias/agua-clara")
        self.assertEqual(clean_url("https://www.mpms.mp.br/outro"), "https://www.mpms.mp.br/outro")
        self.assertEqual(clean_url(""), "")

    def test_get_cities_parses_html_correctly(self):
        from unittest.mock import patch
        from src.scrapers.unidades_scraper import get_cities

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

    def test_scrape_promotoria_parses_html_correctly(self):
        from unittest.mock import patch
        from src.scrapers.unidades_scraper import scrape_promotoria

        fake_html = """
        <div id="promotorias">
            <h2>2ª Promotoria de Justiça de Três Lagoas</h2>
            <p class="titular"><span class="name">Titular: Dr. João da Silva</span></p>
            <address>Rua Elviro Mario Mancini, 860 - Centro - Três Lagoas - CEP 79601-020</address>
            <p class="phone">Telefone: (67) 3521-1234</p>
        </div>
        """
        mock_response = MagicMock()
        mock_response.text = fake_html
        mock_response.status_code = 200

        with patch('requests.get', return_value=mock_response):
            res = scrape_promotoria('Três Lagoas', 'https://www.mpms.mp.br/promotorias/tres-lagoas/2-promotoria')
            self.assertEqual(res['Setor'], '2ª Promotoria de Justiça de Três Lagoas')
            self.assertEqual(res['Titular'], 'Dr. João da Silva')
            self.assertEqual(res['Unidade (Prédio)'], 'Três Lagoas - Sede')
            self.assertEqual(res['Telefone'], '(67) 3521-1234')


class TestDateSanitization(unittest.TestCase):

    """Testes para o tratamento de fuso horário e padronização ISO de datas."""

    def test_date_conversion_to_iso_string(self):
        # Testa se as datas no formato brasileiro (ou qualquer formato válido) são convertidas para string ISO
        df = pd.DataFrame({
            "Data Criação": ["07/05/2026 12:59:00", "19/12/2025 05:08:00"]
        })
        
        # Simula o fluxo de conversão
        dt_col = pd.to_datetime(df['Data Criação'], errors='coerce', dayfirst=True)
        df['Data Criação'] = dt_col.dt.strftime('%Y-%m-%d %H:%M:%S').fillna(df['Data Criação'])
        
        self.assertEqual(df.loc[0, 'Data Criação'], "2026-05-07 12:59:00")
        self.assertEqual(df.loc[1, 'Data Criação'], "2025-12-19 05:08:00")


class TestRemoteLocationExtraction(unittest.TestCase):
    """Testes para o redirecionamento inteligente de técnicos remotos via NLP/Heurísticas."""

    def test_felipe_ferrari_case_ricardo_brandao_ii(self):
        # Caso 1: Felipe Ferrari (Costa Rica -> Ricardo Brandão II)
        df = pd.DataFrame({
            "Chamado#": [84990],
            "Nome do Usuário": ["Felipe Ferrari Marcolin"],
            "Cidade - Prédio": ["Costa Rica - Sede"],
            "Unidade": ["Costa Rica - Sede"],
            "Descrição": ["Felipe Ferrari Marcolin está lotado em Costa Rica - Sede, mas Angela está trabalhando na Sala do Suporte de Apoio Remoto na Ricardo Brandão - Unidade II e solicitou apoio na instalação do software."]
        })
        df_updated = detect_and_update_remote_locations(df)
        self.assertEqual(df_updated.loc[0, "Localidade física"], "Campo Grande - Ricardo Brandão II")
        self.assertEqual(df_updated.loc[0, "Cidade - Prédio"], "Costa Rica - Sede")
        self.assertEqual(df_updated.loc[0, "Unidade"], "Costa Rica - Sede")


    def test_nadson_borges_case_chacara_cachoeira_ii(self):
        # Caso 2: Nadson Borges (Aquidauana -> Chácara Cachoeira II)
        df = pd.DataFrame({
            "Chamado#": [84995],
            "Nome do Usuário": ["Nadson Matheus Borges"],
            "Cidade - Prédio": ["Aquidauana - Sede"],
            "Unidade": ["Aquidauana - Sede"],
            "Descrição": ["Considerando que passei a desempenhar minhas funções em teletrabalho a partir de hoje, na sala do apoio remoto, na Unidade Chácara Cachoeira II, peço revisão de rede."]
        })
        df_updated = detect_and_update_remote_locations(df)
        self.assertEqual(df_updated.loc[0, "Localidade física"], "Campo Grande - Chácara Cachoeira II")
        self.assertEqual(df_updated.loc[0, "Cidade - Prédio"], "Aquidauana - Sede")
        self.assertEqual(df_updated.loc[0, "Unidade"], "Aquidauana - Sede")


    def test_false_positive_no_remote_context(self):
        # Caso 3: Menção à Chácara Cachoeira mas sem contexto de trabalho remoto
        df = pd.DataFrame({
            "Chamado#": [85000],
            "Nome do Usuário": ["José Silva"],
            "Cidade - Prédio": ["Aquidauana - Sede"],
            "Unidade": ["Aquidauana - Sede"],
            "Descrição": ["Minha impressora aqui em Aquidauana quebrou. Preciso de suporte urgente. Não tem nada a ver com Chácara Cachoeira."]
        })
        df_updated = detect_and_update_remote_locations(df)
        # Não deve alterar
        self.assertEqual(df_updated.loc[0, "Cidade - Prédio"], "Aquidauana - Sede")
        self.assertEqual(df_updated.loc[0, "Unidade"], "Aquidauana - Sede")

    def test_remote_context_without_ii_suffix(self):
        # Caso 4: Presença remota em prédio padrão (sem II)
        df = pd.DataFrame({
            "Chamado#": [85010],
            "Nome do Usuário": ["Reginaldo Vilanova"],
            "Cidade - Prédio": ["Ponta Porã - Sede"],
            "Unidade": ["Ponta Porã - Sede"],
            "Descrição": ["Estou trabalhando temporariamente no prédio da Rua da Paz para acompanhar o treinamento da equipe."]
        })
        df_updated = detect_and_update_remote_locations(df)
        self.assertEqual(df_updated.loc[0, "Localidade física"], "Campo Grande - Rua da Paz")
        self.assertEqual(df_updated.loc[0, "Cidade - Prédio"], "Ponta Porã - Sede")
        self.assertEqual(df_updated.loc[0, "Unidade"], "Ponta Porã - Sede")


    def test_normalize_for_extraction(self):
        # Caso 5: Limpeza e normalização fina de acentuações, casing e tipos nulos
        self.assertEqual(normalize_for_extraction("Olá, Chácara Cachoeira, Ricardo Brandão, GAECO!"), "ola, chacara cachoeira, ricardo brandao, gaeco!")
        self.assertEqual(normalize_for_extraction("AQUIDAÚANA, CORUMBÁ, TRABALHO REMÔTO."), "aquidauana, corumba, trabalho remoto.")
        self.assertEqual(normalize_for_extraction(None), "")
        self.assertEqual(normalize_for_extraction(123), "")


class TestOtrsCommentsCleaning(unittest.TestCase):
    """Testes unitários para a limpeza de comentários do OTRS."""

    def test_clean_otrs_comments_filtering(self):
        from src.config import clean_otrs_comments
        comments = [
            {'data': '2026-05-21 10:00:00', 'autor': 'suporte@mpms.mp.br', 'texto': 'Comentário automático que deve ser ignorado'},
            {'data': '2026-05-21 10:05:00', 'autor': 'Central de Atendimento ao Usuário', 'texto': 'Outro comentário automático a ser ignorado'},
            {'data': '2026-05-21 10:10:00', 'autor': 'paulo.goncalves', 'texto': 'Comentário legítimo que deve ser mantido'}
        ]
        cleaned = clean_otrs_comments(comments)
        self.assertEqual(len(cleaned), 1)
        self.assertEqual(cleaned[0]['autor'], 'paulo.goncalves')

    def test_clean_otrs_comments_invalid_inputs(self):
        from src.config import clean_otrs_comments
        self.assertEqual(clean_otrs_comments(None), [])
        self.assertEqual(clean_otrs_comments('[]'), [])
        self.assertEqual(clean_otrs_comments('invalid json'), [])
        self.assertEqual(clean_otrs_comments(float('nan')), [])


class TestSccmAndHostnameParsing(unittest.TestCase):
    """Testes unitários para a extração de dados do SCCM via PowerShell/WMI/JSON."""

    @unittest.mock.patch('subprocess.run')
    def test_fetch_sccm_data_success_json(self, mock_run):
        from src.config import fetch_sccm_data, _sccm_cache
        _sccm_cache.clear()
        
        fake_stdout = """
        {
            "Name": "SRV-TESTE-01",
            "IPAddresses": ["192.168.1.5", "10.10.20.30"]
        }
        """
        mock_run.return_value = MagicMock(returncode=0, stdout=fake_stdout, stderr='')
        
        res = fetch_sccm_data('testuser')
        self.assertEqual(res['ip'], '10.10.20.30')
        self.assertEqual(res['hostname'], 'SRV-TESTE-01')

    @unittest.mock.patch('subprocess.run')
    def test_fetch_sccm_data_access_denied(self, mock_run):
        from src.config import fetch_sccm_data, _sccm_cache
        _sccm_cache.clear()
        
        mock_run.return_value = MagicMock(returncode=1, stdout='', stderr='Get-WmiObject : Acesso negado')
        
        res = fetch_sccm_data('testuser')
        self.assertEqual(res['ip'], 'Acesso Negado')
        self.assertEqual(res['hostname'], 'Acesso Negado')

    @unittest.mock.patch('subprocess.run')
    def test_fetch_sccm_data_regex_fallback(self, mock_run):
        from src.config import fetch_sccm_data, _sccm_cache
        _sccm_cache.clear()
        
        fake_stdout = 'Name : DESKTOP-ABC123\nIPAddresses : {192.168.0.10, 10.50.60.70}'
        mock_run.return_value = MagicMock(returncode=0, stdout=fake_stdout, stderr='')
        
        res = fetch_sccm_data('fallbackuser')
        self.assertEqual(res['ip'], '10.50.60.70')
        self.assertEqual(res['hostname'], 'DESKTOP-ABC123')


class TestAdLookupRobustness(unittest.TestCase):
    """Testes unitários para a busca de departamento/unidade no AD de forma exata e robusta."""

    def test_fetch_ad_department_username_exact_filter(self):
        from src.config import fetch_ad_department
        mock_conn = MagicMock()
        
        mock_entry = {
            'department': ['Secretaria de Tecnologia da Informação'],
            'physicalDeliveryOfficeName': ['Edifício Sede']
        }
        mock_conn.entries = [MagicMock(entry_attributes_as_dict=mock_entry)]
        
        dept = fetch_ad_department(mock_conn, 'larasantos', is_username=True)
        self.assertEqual(dept, 'Secretaria de Tecnologia da Informação')
        
        args, kwargs = mock_conn.search.call_args
        self.assertIn('(sAMAccountName=larasantos)', kwargs.get('search_filter', ''))

    def test_fetch_ad_department_fallback_to_office(self):
        from src.config import fetch_ad_department
        mock_conn = MagicMock()
        
        mock_entry = {
            'department': [],
            'physicalDeliveryOfficeName': ['Promotoria de Três Lagoas']
        }
        mock_conn.entries = [MagicMock(entry_attributes_as_dict=mock_entry)]
        
        dept = fetch_ad_department(mock_conn, 'testuser', is_username=True)
        self.assertEqual(dept, 'Promotoria de Três Lagoas')


class TestCloseMissingTickets(unittest.TestCase):
    """Testes unitários para a funcionalidade de fechamento automático de chamados ausentes."""

    def test_close_missing_tickets_by_base(self):
        from src.database import save_tickets_to_db, close_missing_tickets_by_base, load_data
        
        df_fake = pd.DataFrame([
            {
                'Chamado#': '99901',
                'Título': 'Teste Fechamento 1',
                'Nome do Usuário': 'Usuário Teste',
                'Base': 'OTRS'
            },
            {
                'Chamado#': '99902',
                'Título': 'Teste Fechamento 2',
                'Nome do Usuário': 'Usuário Teste 2',
                'Base': 'OTRS'
            },
            {
                'Chamado#': '99903',
                'Título': 'Teste Fechamento 3',
                'Nome do Usuário': 'Usuário Teste 3',
                'Base': 'OTRS'
            }
        ])
        save_tickets_to_db(df_fake)
        
        # Simula nova rodada de raspagem onde o chamado 99902 foi fechado no OTRS e não está mais nos ativos
        active_ids_otrs = ['99901', '99903', '99904']
        closed_count = close_missing_tickets_by_base(active_ids_otrs, 'OTRS')
        self.assertGreaterEqual(closed_count, 1)
        
        df_updated = load_data()
        ticket_99902 = df_updated[df_updated['id'] == '99902']
        if not ticket_99902.empty:
            self.assertEqual(ticket_99902.iloc[0]['status'], 'Fechado')


if __name__ == '__main__':
    unittest.main()

