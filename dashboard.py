import warnings
warnings.filterwarnings("ignore", category=UserWarning, module="streamlit")
warnings.filterwarnings("ignore", message=".*use_container_width.*")
warnings.filterwarnings("ignore", message=".*st.components.v1.html.*")

import sys
import asyncio
import locale
import importlib
from pathlib import Path
import streamlit as st
from dotenv import load_dotenv

# Adiciona a raiz do projeto e a pasta src ao sys.path para importações de módulos
root_dir = Path(__file__).parent
sys.path.insert(0, str(root_dir))
sys.path.insert(0, str(root_dir / "src"))

# Carrega variáveis do arquivo .env
load_dotenv()

# Inicialização de cores ANSI e UTF-8 no terminal
from src.terminal import log, print_header, CYAN, GREEN, YELLOW

# Silencia o aviso WinError 10054 (Connection Reset) comum no Windows asyncio
if sys.platform == 'win32':
    try:
        from asyncio.proactor_events import _ProactorBasePipeTransport
        _orig_call_connection_lost = _ProactorBasePipeTransport._call_connection_lost

        def _silenced_call_connection_lost(self, exc):
            try:
                _orig_call_connection_lost(self, exc)
            except (ConnectionResetError, ConnectionAbortedError, OSError):
                pass

        _ProactorBasePipeTransport._call_connection_lost = _silenced_call_connection_lost
    except Exception:
        pass

def handle_asyncio_exception(loop, context):
    exception = context.get('exception')
    if isinstance(exception, (ConnectionResetError, ConnectionAbortedError, BrokenPipeError, OSError)):
        return
    message = context.get('message', '')
    if 'ConnectionResetError' in message or '10054' in message or 'connection_lost' in message:
        return
    loop.default_exception_handler(context)

try:
    loop = asyncio.get_event_loop()
    loop.set_exception_handler(handle_asyncio_exception)
except Exception:
    pass


# Configuração do locale para Português do Brasil
try:
    locale.setlocale(locale.LC_ALL, 'pt_BR.utf8')
except Exception:
    try:
        locale.setlocale(locale.LC_ALL, 'Portuguese_Brazil.1252')
    except Exception:
        pass

# Configuração de Página Streamlit
st.set_page_config(page_title="Painel de Chamados - STI", layout="wide")

# Carrega arquivo de estilos CSS globais da pasta assets
css_path = root_dir / "assets" / "css" / "styles.css"
if css_path.exists():
    with open(css_path, "r", encoding="utf-8") as f:
        st.markdown(f"<style>{f.read()}</style>", unsafe_allow_html=True)

# Importa o componente de navegação no header
import src.components.header
importlib.reload(src.components.header)
selected_page = src.components.header.render_header_navigation()

if selected_page == "🏢 Catálogo de Unidades":
    import src.tabs.unidades
    importlib.reload(src.tabs.unidades)
    src.tabs.unidades.render_unidades_page()

elif selected_page == "📞 Central Telefônica (OXE)":
    import src.tabs.central_telefonica
    importlib.reload(src.tabs.central_telefonica)
    src.tabs.central_telefonica.render_central_telefonica_page()

elif selected_page == "📅 Plantões da Bancada":
    import src.tabs.plantoes
    importlib.reload(src.tabs.plantoes)
    src.tabs.plantoes.render_plantoes_page()

elif selected_page == "📅 Calendário Geral":
    import src.tabs.calendario_geral
    importlib.reload(src.tabs.calendario_geral)
    src.tabs.calendario_geral.render_calendario_geral_page()

elif selected_page == "📍 Mapa & Localização":
    import src.tabs.mapas
    importlib.reload(src.tabs.mapas)
    src.tabs.mapas.render_mapa_page()

elif selected_page == "🖥️ Doação & Redistribuição":
    import src.tabs.redistribuicao
    importlib.reload(src.tabs.redistribuicao)
    src.tabs.redistribuicao.render_donations_page()

elif selected_page == "📜 Fiscalização de Contratos":
    import src.tabs.fiscalizacao
    importlib.reload(src.tabs.fiscalizacao)
    src.tabs.fiscalizacao.render_contracts_page()

elif selected_page == "✈️ Viagens da Bancada":
    import src.tabs.viagens
    importlib.reload(src.tabs.viagens)
    src.tabs.viagens.render_viagens_page()

elif selected_page == "🛡️ Controle de Garantia":
    import src.tabs.garantia
    importlib.reload(src.tabs.garantia)
    src.tabs.garantia.render_garantia_page()


elif selected_page == "📚 FAQ & Tutoriais":
    import src.tabs.links_faqs
    importlib.reload(src.tabs.links_faqs)
    src.tabs.links_faqs.render_faq_page()

elif selected_page == "📜 Portarias da Bancada":
    import src.tabs.portarias
    importlib.reload(src.tabs.portarias)
    src.tabs.portarias.render_portarias_page()

elif selected_page == "🖨️ Impressoras (PaperCut)":
    import src.tabs.impressoras
    importlib.reload(src.tabs.impressoras)
    src.tabs.impressoras.render_impressoras_page()

elif selected_page == "⚡ Scripts de Automação":
    import src.tabs.scripts_automacao
    importlib.reload(src.tabs.scripts_automacao)
    src.tabs.scripts_automacao.render_scripts_automacao_page()

elif selected_page == "🔔 Central de Notificações":

    import src.tabs.notificacoes
    importlib.reload(src.tabs.notificacoes)
    src.tabs.notificacoes.render_notificacoes_page()

elif selected_page == "⚙️ Configurações":
    import src.tabs.configuracoes
    importlib.reload(src.tabs.configuracoes)
    src.tabs.configuracoes.render_configuracoes_page()

else:
    import src.tabs.chamados
    importlib.reload(src.tabs.chamados)
    src.tabs.chamados.render_chamados_page()


