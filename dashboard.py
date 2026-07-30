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

# Silencia o aviso WinError 10054 (Connection Reset) comum no Windows asyncio
if sys.platform == 'win32':
    try:
        asyncio.set_event_loop_policy(asyncio.WindowsSelectorEventLoopPolicy())
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

# Roteamento centralizado do Dashboard (Orquestrador) com Hot-Reload habilitado
if selected_page == "📅 Plantões da Bancada":
    import src.tabs.plantoes
    importlib.reload(src.tabs.plantoes)
    src.tabs.plantoes.render_plantoes_page()

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

else:
    import src.tabs.chamados
    importlib.reload(src.tabs.chamados)
    src.tabs.chamados.render_chamados_page()


