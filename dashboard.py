import sys
from pathlib import Path

# Adiciona a raiz do projeto e a pasta src ao sys.path para importações de módulos
root_dir = Path(__file__).parent
sys.path.insert(0, str(root_dir))
sys.path.insert(0, str(root_dir / "src"))

import asyncio
import os
from dotenv import load_dotenv

# Carrega variáveis do arquivo .env
load_dotenv()

# Silencia o aviso WinError 10054 (Connection Reset) comum no Windows asyncio
if sys.platform == 'win32':
    try:
        asyncio.set_event_loop_policy(asyncio.WindowsSelectorEventLoopPolicy())
    except:
        pass

import streamlit as st
import streamlit.components.v1 as components
import pandas as pd
import sqlite3
from pathlib import Path
from datetime import datetime
import locale
from src.config import DEBUG_DIR_LEAFLET, setup_logging
logger = setup_logging(DEBUG_DIR_LEAFLET / "leaflet.log", "leaflet")


# Declara o componente customizado do Select Multiple nativo com clique-e-arraste
custom_select_dir = Path(__file__).parent / "custom_select_component"
custom_select = components.declare_component("custom_select", path=str(custom_select_dir))

# Configuração do locale para Português do Brasil
try:
    locale.setlocale(locale.LC_ALL, 'pt_BR.utf8')
except:
    try:
        locale.setlocale(locale.LC_ALL, 'Portuguese_Brazil.1252')
    except:
        pass

# Configuração da página para ocupar a tela toda e ter um título
st.set_page_config(page_title="Painel de Chamados - STI", layout="wide")

# CSS para ocultar apenas o botão Deploy e o rodapé padrão em inglês, mantendo o menu de três pontinhos
st.markdown("""
    <style>
    /* Oculta apenas o botão de Deploy do cabeçalho */
    [data-testid="stAppDeployButton"], .stAppDeployButton {
        display: none !important;
    }
    /* Remove rodapé padrão */
    footer {
        display: none !important;
    }
    /* Borda fina vermelha e glow suave para o modal (st.dialog) */
    div[data-testid="stDialog"] > div:first-child,
    div[role="dialog"] {
        border: 2px solid #ff4b4b !important;
        box-shadow: 0 0 15px rgba(255, 75, 75, 0.3) !important;
        border-radius: 8px !important;
    }
    /* Reduz drasticamente o padding superior e ancora o posicionamento absoluto */
    .block-container, div[data-testid="stMainBlockContainer"] {
        padding-top: 1.5rem !important;
        padding-bottom: 1rem !important;
        position: relative !important;
    }
    /* Posiciona o botão Popover de navegação no header nativo do Streamlit, compacto no canto superior direito */
    header[data-testid="stHeader"] {
        background: transparent !important;
        pointer-events: none !important;
    }
    header[data-testid="stHeader"] * {
        pointer-events: auto !important;
    }
    div[data-testid="stMainBlockContainer"] > div:first-child > div[data-testid="stPopover"],
    .main div[data-testid="stPopover"] {
        position: fixed !important;
        top: 0.4rem !important;
        right: 4.5rem !important;
        width: auto !important;
        z-index: 999999 !important;
    }
    div[data-testid="stMainBlockContainer"] > div:first-child > div[data-testid="stPopover"] > button,
    .main div[data-testid="stPopover"] > button {
        background-color: #1e1f25 !important;
        border: 1px solid #343541 !important;
        padding: 4px 12px !important;
        height: auto !important;
        min-height: 0px !important;
        width: auto !important;
        font-size: 0.85rem !important;
        color: #ffffff !important;
        border-radius: 6px !important;
    }
    div[data-testid="stPopover"] > button:hover {
        border-color: #ff4b4b !important;
        color: #ff4b4b !important;
    }
    /* Garante que o dataframe ocupe 100% da tela quando entrar no modo Tela Cheia (Fullscreen) */
    div[data-testid="stDataFrame"][data-st-mode="fullscreen"],
    div[data-testid="stElementContainer"]:aria-modal,
    :fullscreen div[data-testid="stDataFrame"],
    :fullscreen div[data-testid="stDataFrame"] > div,
    [data-testid="stDataFrame"]:fullscreen,
    [data-testid="stDataFrame"]:fullscreen iframe {
        width: 100vw !important;
        height: 100vh !important;
        max-width: 100vw !important;
        max-height: 100vh !important;
    }
    </style>
""", unsafe_allow_html=True)



def check_orquestrador_running() -> bool:
    """Verifica se o orquestrador está rodando de forma ativa analisando o arquivo de lock no Windows."""
    import tempfile
    import ctypes
    from pathlib import Path
    
    lock_file = Path(tempfile.gettempdir()) / "automated_otrs_citsmart.lock"
    if not lock_file.exists():
        return False
        
    try:
        with open(lock_file, "r") as f:
            pid = int(f.read().strip())
        
        # Verifica se o processo com esse PID está ativo
        PROCESS_QUERY_LIMITED_INFORMATION = 0x1000
        kernel32 = ctypes.windll.kernel32
        handle = kernel32.OpenProcess(PROCESS_QUERY_LIMITED_INFORMATION, False, pid)
        if handle:
            exit_code = ctypes.c_ulong()
            if kernel32.GetExitCodeProcess(handle, ctypes.byref(exit_code)):
                kernel32.CloseHandle(handle)
                return exit_code.value == 259  # 259 significa STILL_ACTIVE
            kernel32.CloseHandle(handle)
    except:
        pass
    return False

def read_last_log_lines(n: int = 15) -> str:
    """Lê as últimas N linhas do arquivo de log do orquestrador."""
    log_path = Path("debug_logs") / "orquestrador" / "orquestrador.log"
    if not log_path.exists():
        return "Nenhum log gerado ainda. Aguardando início..."
    try:
        with open(log_path, "r", encoding="utf-8", errors="replace") as f:
            lines = f.readlines()
            return "".join(lines[-n:])
    except Exception as e:
        return f"Erro ao ler arquivo de log: {e}"

def get_image_dimensions(image_path: Path):
    """Retorna largura e altura da imagem, ou fallback caso falhe."""
    try:
        from PIL import Image
        with Image.open(image_path) as img:
            return img.width, img.height
    except Exception:
        return 1000, 1000


def get_image_base64(image_path: Path) -> str:
    """Carrega a imagem física e retorna como Data URI base64."""
    import base64
    try:
        if not image_path.exists():
            return ""
        with open(image_path, "rb") as f:
            encoded = base64.b64encode(f.read()).decode("utf-8")
            ext = image_path.suffix.lower()
            mimetype = "image/png"
            if ext in [".jpg", ".jpeg"]:
                mimetype = "image/jpeg"
            elif ext == ".gif":
                mimetype = "image/gif"
            return f"data:{mimetype};base64,{encoded}"
    except Exception:
        return ""

import heapq
import math

def calculate_dijkstra_route(caminhos: dict, start_pin: dict, end_pin: dict) -> list:
    """
    Calcula a rota mais curta entre o pin de origem e o pin de destino
    usando a malha de caminhos (nós e arestas) com o algoritmo de Dijkstra.
    """
    nos = caminhos.get("nós", [])
    if not nos:
        return []
        
    # 1. Encontra o nó de navegação mais próximo para o pin de origem (mesmo pavimento)
    start_no = None
    min_start_dist = float("inf")
    for no in nos:
        if no["pavimento_id"] == start_pin["pavimento_id"]:
            dist = math.sqrt((no["x"] - start_pin["x"])**2 + (no["y"] - start_pin["y"])**2)
            if dist < min_start_dist:
                min_start_dist = dist
                start_no = no
                
    # 2. Encontra o nó de navegação mais próximo para o pin de destino (mesmo pavimento)
    end_no = None
    min_end_dist = float("inf")
    for no in nos:
        if no["pavimento_id"] == end_pin["pavimento_id"]:
            dist = math.sqrt((no["x"] - end_pin["x"])**2 + (no["y"] - end_pin["y"])**2)
            if dist < min_end_dist:
                min_end_dist = dist
                end_no = no
                
    if not start_no or not end_no:
        return []
        
    # 3. Constrói adjacências do Grafo
    nodes_map = {n["id"]: n for n in nos}
    adj = {nid: [] for nid in nodes_map}
    
    for edge in caminhos.get("arestas", []):
        u = edge.get("de")
        v = edge.get("para")
        if u in nodes_map and v in nodes_map:
            n1 = nodes_map[u]
            n2 = nodes_map[v]
            
            # Custo: distância euclidiana física (ou penalidade se mudar de andar)
            if n1["pavimento_id"] != n2["pavimento_id"]:
                weight = 300.0  # Custo fixo para mudar de pavimento
            else:
                weight = math.sqrt((n1["x"] - n2["x"])**2 + (n1["y"] - n2["y"])**2)
                
            adj[u].append((v, weight))
            adj[v].append((u, weight))
            
    # Dijkstra
    queue = [(0.0, start_no["id"], [start_no["id"]])]
    visited = set()
    
    while queue:
        dist, curr, path = heapq.heappop(queue)
        if curr in visited:
            continue
        visited.add(curr)
        
        if curr == end_no["id"]:
            # Reconstrói a rota final em formato de lista de dicionários contendo os nós
            return [nodes_map[nid] for nid in path]
            
        for neighbor, weight in adj[curr]:
            if neighbor not in visited:
                heapq.heappush(queue, (dist + weight, neighbor, path + [neighbor]))
                
    return []


import threading
from http.server import BaseHTTPRequestHandler, HTTPServer
import json

class SaveConfigHandler(BaseHTTPRequestHandler):
    def do_OPTIONS(self):
        self.send_response(200)
        self.send_header('Access-Control-Allow-Origin', '*')
        self.send_header('Access-Control-Allow-Methods', 'POST, GET, OPTIONS')
        self.send_header('Access-Control-Allow-Headers', 'Content-Type')
        self.end_headers()

    def do_GET(self):
        import streamlit as st
        if not hasattr(st, "_global_route"):
            st._global_route = {"origem": "", "destino": ""}
            
        from urllib.parse import urlparse, parse_qs
        parsed = urlparse(self.path)
        logger.debug(f"🌐 GET request recebida no servidor Leaflet backend: {parsed.path}")
        if parsed.path == '/set_route':
            query = parse_qs(parsed.query)
            logger.debug(f"📍 Parâmetros query recebidos: {query}")
            
            # Atualiza st._global_route
            if 'origem' in query:
                st._global_route['origem'] = query['origem'][0]
                logger.debug(f"👉 Origem atualizada no global_route para: {query['origem'][0]}")
            if 'destino' in query:
                st._global_route['destino'] = query['destino'][0]
                logger.debug(f"👉 Destino atualizado no global_route para: {query['destino'][0]}")
                
            self.send_response(200)
            self.send_header('Access-Control-Allow-Origin', '*')
            self.send_header('Content-Type', 'application/json')
            self.end_headers()
            self.wfile.write(json.dumps({"status": "success"}).encode())
            
            # Atualiza também o st.session_state da aplicação ativa de forma segura
            try:
                from streamlit.runtime import get_instance
                runtime = get_instance()
                active_sessions = runtime._session_mgr.list_active_sessions()
                logger.debug(f"👥 Total de sessões Streamlit ativas encontradas: {len(active_sessions)}")
                for session_info in active_sessions:
                    session_state = session_info.session.session_state
                    
                    # Carrega pins para saber os nomes de exibição correspondentes
                    from src.database import get_map_config
                    config = get_map_config()
                    todos_pins = []
                    for pr in config.get("predios", []):
                        for p in pr.get("pins", []):
                            todos_pins.append(p)
                            
                    if 'origem' in query:
                        orig_id = query['origem'][0]
                        orig_match = next((p for p in todos_pins if p["id"] == orig_id), None)
                        if orig_match:
                            display_name = f"{orig_match['sala']} ({orig_match['pavimento_id']}º Andar)" if orig_match['pavimento_id'] > 0 else f"{orig_match['sala']} (Térreo)"
                            session_state["sb_origem"] = display_name
                            logger.debug(f" Sincronizado sb_origem na sessão {session_info.session.id}: {display_name}")
                    
                    if 'destino' in query:
                        dest_id = query['destino'][0]
                        dest_match = next((p for p in todos_pins if p["id"] == dest_id), None)
                        if dest_match:
                            display_name = f"{dest_match['sala']} ({dest_match['pavimento_id']}º Andar)" if dest_match['pavimento_id'] > 0 else f"{dest_match['sala']} (Térreo)"
                            session_state["sb_destino"] = display_name
                            logger.debug(f" Sincronizado sb_destino na sessão {session_info.session.id}: {display_name}")
                            
                # Dispara o rerun de forma assíncrona com delay de 250ms usando uma thread leve
                # Isso dá tempo para o Leaflet fechar a requisição pendente e evita descartar cliques do usuário na UI.
                def delayed_rerun():
                    import time
                    time.sleep(0.25)
                    logger.debug("⏰ Rerunning active sessions pós-delay...")
                    try:
                        for s_info in active_sessions:
                            s_info.session.request_rerun(None)
                    except Exception as ex:
                        logger.error(f"Erro no delayed rerun: {ex}")
                        
                threading.Thread(target=delayed_rerun, daemon=True).start()
            except Exception as e:
                logger.error(f"❌ Erro ao atualizar session_state na rota set_route: {e}", exc_info=True)
            return
        self.send_response(404)
        self.end_headers()

    def do_POST(self):
        logger.debug(f"🌐 POST request recebida no servidor Leaflet backend: {self.path}")
        if self.path == '/save_config':
            content_length = int(self.headers['Content-Length'])
            post_data = self.rfile.read(content_length)
            try:
                config_data = json.loads(post_data.decode('utf-8'))
                from src.database import save_map_config
                save_map_config(config_data)
                logger.debug("💾 Configurações do mapa salvas com sucesso no banco SQLite.")
                
                # Salva também no arquivo físico uploads/map_config_TEMPLATE.json
                template_path = Path("uploads/map_config_TEMPLATE.json")
                with open(template_path, "w", encoding="utf-8") as f:
                    json.dump(config_data, f, indent=2, ensure_ascii=False)
                logger.debug(f"💾 Cópia física de backup salva em: {template_path}")
                
                self.send_response(200)
                self.send_header('Access-Control-Allow-Origin', '*')
                self.send_header('Content-Type', 'application/json')
                self.end_headers()
                self.wfile.write(json.dumps({"status": "success"}).encode())
                return
            except Exception as e:
                logger.error(f"❌ Erro ao salvar configurações no POST do backend: {e}", exc_info=True)
                self.send_response(500)
                self.send_header('Access-Control-Allow-Origin', '*')
                self.end_headers()
                self.wfile.write(str(e).encode())
                return
        self.send_response(404)
        self.end_headers()

def start_backend_server():
    if not hasattr(st, "_backend_server_running"):
        st._backend_server_running = True
        def run_server():
            server = HTTPServer(('localhost', 8099), SaveConfigHandler)
            server.serve_forever()
        t = threading.Thread(target=run_server, daemon=True)
        t.start()


def render_mapa_page():
    """Renderiza a página/aba de Mapa & Localização."""
    start_backend_server()
    from src.database import get_map_config, get_map_pins, save_map_config
    import json
    
    # Obtém a rota ativa via backend
    if not hasattr(st, "_global_route"):
        st._global_route = {"origem": "", "destino": ""}
    url_origem = st._global_route.get("origem", "")
    url_destino = st._global_route.get("destino", "")
    print(f"DEBUG: st._global_route={st._global_route}, url_origem={url_origem}, url_destino={url_destino}")
    
    st.title("📍 Mapa & Localização de Chamados")
    st.write("Visualize no mapa/planta baixa a localização exata das salas de atendimento.")
    
    # 1. Seção de Importação de JSON
    #with st.sidebar.expander("📥 Configurações & Upload JSON", expanded=False):
    #    st.write("Atualize a planta e os locais enviando um JSON formatado:")
    #    uploaded_file = st.file_uploader("Escolher arquivo JSON", type=["json"])
    #    if uploaded_file is not None:
    #        try:
    #            config_data = json.load(uploaded_file)
    #            if "predios" in config_data:
    #                save_map_config(config_data)
    #                st.success("Configurações do mapa e pins importadas com sucesso!")
    #                st.cache_resource.clear()
    #                st.cache_data.clear()
    #                st.rerun()
    #            else:
    #                st.error("JSON inválido! Deve conter a chave 'predios'.")
    #        except Exception as e:
    #            st.error(f"Erro ao processar arquivo: {e}")
                
    # 2. Carrega as configurações do banco
    config = get_map_config()
    predios = config.get("predios", [])
    
    if not predios:
        st.info("Nenhum prédio cadastrado no banco de dados. Faça o upload de um JSON de configurações na barra lateral.")
        return
        
    # Adiciona os seletores e busca diretamente na barra lateral, liberando espaço total para a imagem
    # st.sidebar.markdown("---")
    # st.sidebar.subheader("📍 Seleção do Local")
    
    # Seleção de prédio
    predio_nomes = [p.get("nome") for p in predios]
    selected_predio_nome = st.sidebar.selectbox("Selecione o Prédio", predio_nomes)
    selected_predio = next(p for p in predios if p.get("nome") == selected_predio_nome)
    predio_id = selected_predio.get("id")
    logger.debug(f"🏢 Prédio selecionado na UI: {selected_predio_nome} (ID: {predio_id})")
    
    # Sincroniza a rota global com o Session State do Streamlit para evitar perdas de estado
    todos_pins = get_map_pins(predio_id)
    print(f"DEBUG: todos_pins count={len(todos_pins)}")
    if url_origem:
        orig_match = next((p for p in todos_pins if p["id"] == url_origem), None)
        print(f"DEBUG: url_origem={url_origem}, orig_match={orig_match}")
        if orig_match:
            display_name = f"{orig_match['sala']} ({orig_match['pavimento_id']}º Andar)" if orig_match['pavimento_id'] > 0 else f"{orig_match['sala']} (Térreo)"
            st.session_state.sb_origem = display_name
            print(f"DEBUG: Set st.session_state.sb_origem={display_name}")
    elif "sb_origem" not in st.session_state:
        st.session_state.sb_origem = "-- Selecione a Origem --"
        
    if url_destino:
        dest_match = next((p for p in todos_pins if p["id"] == url_destino), None)
        print(f"DEBUG: url_destino={url_destino}, dest_match={dest_match}")
        if dest_match:
            display_name = f"{dest_match['sala']} ({dest_match['pavimento_id']}º Andar)" if dest_match['pavimento_id'] > 0 else f"{dest_match['sala']} (Térreo)"
            st.session_state.sb_destino = display_name
            print(f"DEBUG: Set st.session_state.sb_destino={display_name}")
    elif "sb_destino" not in st.session_state:
        st.session_state.sb_destino = "-- Selecione o Destino --"
    
    # Seleção de pavimento
    pavimentos = selected_predio.get("pavimentos", [])
    if not pavimentos:
        st.sidebar.warning("Sem pavimentos.")
        return
        
    pavimento_nomes = [pav.get("nome") for pav in pavimentos]
    
    # Gerencia estado do pavimento selecionado
    if "prev_predio_id" not in st.session_state or st.session_state.prev_predio_id != predio_id:
        st.session_state.prev_predio_id = predio_id
        st.session_state.selected_pavimento_id = pavimentos[0].get("id")
        
    try:
        default_index = next(idx for idx, pav in enumerate(pavimentos) if pav.get("id") == st.session_state.selected_pavimento_id)
    except StopIteration:
        default_index = 0
 
    selected_pav_nome = st.sidebar.selectbox("Selecione o Pavimento", pavimento_nomes, index=default_index, key="sb_pavimento")
    selected_pav = next(pav for pav in pavimentos if pav.get("nome") == selected_pav_nome)
    pavimento_id = selected_pav.get("id")
    st.session_state.selected_pavimento_id = pavimento_id
    logger.debug(f"📐 Pavimento selecionado na UI: {selected_pav_nome} (ID: {pavimento_id})")
    
    # Obter caminho físico da imagem
    img_path_str = selected_pav.get("imagem")
    img_path = Path(img_path_str)
    
    if not img_path.exists():
        st.error(f"Imagem da planta baixa não encontrada no caminho: `{img_path_str}`")
        return
        
    # Carregar dimensões e base64
    w, h = get_image_dimensions(img_path)
    b64_image = get_image_base64(img_path)
    
    if not b64_image:
        st.error("Erro ao processar a imagem da planta baixa.")
        return
        
    # Carrega os pins do banco para esse pavimento
    pins = get_map_pins(predio_id, pavimento_id)
    
    # 3. Caixa de seleção de pins e busca com botão de limpar
    st.sidebar.markdown("---")
    
    col_sub, col_clear = st.sidebar.columns([2, 1])
    with col_sub:
        st.subheader("🎯 Salas")
    with col_clear:
        st.markdown("<div style='height: 5px;'></div>", unsafe_allow_html=True)
        if st.button("🧹 Limpar", use_container_width=True):
            st._global_route["origem"] = ""
            st._global_route["destino"] = ""
            st.session_state.sb_sala = "-- Selecione uma Sala --"
            st.session_state.txt_busca = ""
            st.rerun()
            
    # Seletor de Sala
    pin_nomes = ["-- Selecione uma Sala --"] + [p["sala"] for p in pins]
    
    default_sb_index = 0
    if "sb_sala" in st.session_state and st.session_state.sb_sala in pin_nomes:
        default_sb_index = pin_nomes.index(st.session_state.sb_sala)
        
    selected_pin_nome = st.sidebar.selectbox("Ir para a Sala", pin_nomes, index=default_sb_index, key="sb_sala")
    active_pin_ids = []
    
    if selected_pin_nome != "-- Selecione uma Sala --":
        active_pin = next(p for p in pins if p["sala"] == selected_pin_nome)
        active_pin_ids.append(active_pin["id"])
        
    # Busca de sala global (busca em todos os andares do prédio)
    default_search_val = st.session_state.get("txt_busca", "")
    search_query = st.sidebar.text_input("🔍 Buscar Sala ou Local (ex: TI, Dr. Fulano)", value=default_search_val, key="txt_busca").strip()
    if search_query:
        all_pins = get_map_pins(predio_id)
        matching_pins = [
            p for p in all_pins 
            if search_query.lower() in p.get("sala", "").lower() or search_query.lower() in p.get("descricao", "").lower()
        ]
        if matching_pins:
            st.sidebar.success(f"✨ Encontrado: {len(matching_pins)} correspondência(s)")
            for p in matching_pins:
                if p["id"] not in active_pin_ids:
                    active_pin_ids.append(p["id"])
            
            # Se o pin encontrado estiver em um pavimento diferente do atual, altera e recarrega
            first_match = matching_pins[0]
            if first_match.get("pavimento_id") != pavimento_id and search_query != "":
                st.session_state.selected_pavimento_id = first_match.get("pavimento_id")
                st.rerun()
        else:
            st.sidebar.warning("⚠️ Nenhum local encontrado.")
            
    # 4. Traçado de Rotas (Pathfinding)
    caminhos = selected_predio.get("caminhos", {})
    route_coords = []
    route_distance_meters = 0.0

    
    if caminhos and caminhos.get("nós") and caminhos.get("arestas"):
        st.sidebar.markdown("---")
        col_route, col_clear_route = st.sidebar.columns([2, 1])
        with col_route:
            st.subheader("🚶 Rota Interna")
        with col_clear_route:
            st.markdown("<div style='height: 5px;'></div>", unsafe_allow_html=True)
            if st.button("🧹 Limpar", key="btn_limpar_rota", use_container_width=True):
                st._global_route["origem"] = ""
                st._global_route["destino"] = ""
                st.session_state.sb_origem = "-- Selecione a Origem --"
                st.session_state.sb_destino = "-- Selecione o Destino --"
                st.rerun()
        
        # Pega pins de todos os andares para origem/destino
        todos_pins = get_map_pins(predio_id)
        pin_origem_nomes = [f"{p['sala']} ({p['pavimento_id']}º Andar)" if p['pavimento_id'] > 0 else f"{p['sala']} (Térreo)" for p in todos_pins]
        
        selected_origem_display = st.sidebar.selectbox("Ponto de Origem", ["-- Selecione a Origem --"] + pin_origem_nomes, key="sb_origem")
        selected_destino_display = st.sidebar.selectbox("Ponto de Destino", ["-- Selecione o Destino --"] + pin_origem_nomes, key="sb_destino")
        
        orig_pin = None
        dest_pin = None
        
        if selected_origem_display != "-- Selecione a Origem --":
            orig_idx = pin_origem_nomes.index(selected_origem_display)
            orig_pin = todos_pins[orig_idx]
            st._global_route["origem"] = orig_pin["id"]
        else:
            st._global_route["origem"] = ""
            
        if selected_destino_display != "-- Selecione o Destino --":
            dest_idx = pin_origem_nomes.index(selected_destino_display)
            dest_pin = todos_pins[dest_idx]
            st._global_route["destino"] = dest_pin["id"]
        else:
            st._global_route["destino"] = ""
            
        if orig_pin and dest_pin:
            if orig_pin["id"] == dest_pin["id"]:
                st.sidebar.info("Origem e Destino são idênticos.")
            else:
                route_nodes = calculate_dijkstra_route(caminhos, orig_pin, dest_pin)
                if route_nodes:
                    # Calcula a distância total percorrida
                    total_dist_pixels = 0.0
                    for idx_n in range(len(route_nodes) - 1):
                        n1 = route_nodes[idx_n]
                        n2 = route_nodes[idx_n+1]
                        if n1["pavimento_id"] != n2["pavimento_id"]:
                            total_dist_pixels += 100.0  # Custo aproximado para mudança de andar (ex: escada/elevador)
                        else:
                            total_dist_pixels += math.sqrt((n1["x"] - n2["x"])**2 + (n1["y"] - n2["y"])**2)
                    
                    # Fator de escala padrão aproximado: 1 pixel = 0.05 metros
                    route_distance_meters = total_dist_pixels * 0.05
                    st.sidebar.success(f"🎉 Rota calculada com sucesso! ({route_distance_meters:.1f} m)")
                    
                    # Filtra nós da rota para o pavimento ativo
                    active_floor_nodes = [n for n in route_nodes if n["pavimento_id"] == pavimento_id]
                    
                    coords_to_draw = []
                    # Se o pin de origem estiver no andar ativo, adiciona-o no início
                    if orig_pin["pavimento_id"] == pavimento_id:
                        coords_to_draw.append([orig_pin["y"], orig_pin["x"]])
                        
                    for n in active_floor_nodes:
                        coords_to_draw.append([n["y"], n["x"]])
                        
                    # Se o pin de destino estiver no andar ativo, adiciona-o no fim
                    if dest_pin["pavimento_id"] == pavimento_id:
                        coords_to_draw.append([dest_pin["y"], dest_pin["x"]])
                        
                    route_coords = coords_to_draw
                    
                    # Se a rota passa por outros andares, sinaliza
                    outros_andares = [n for n in route_nodes if n["pavimento_id"] != pavimento_id]
                    if outros_andares:
                        st.sidebar.warning("⚠️ Rota exige mudança de pavimento! Siga até a escada/elevador e alterne para o pavimento destino para ver a continuação.")
                else:
                    st.sidebar.error("Não foi possível calcular uma rota válida.")

    # Filtra os nós do pavimento ativo para visualização em desenvolvimento
    active_nodes = []
    if caminhos and caminhos.get("nós"):
        active_nodes = [n for n in caminhos.get("nós", []) if n.get("pavimento_id") == pavimento_id]
    active_nodes_json_str = json.dumps(active_nodes)

    # Filtra as arestas do pavimento ativo para visualização em desenvolvimento
    active_arestas = []
    if caminhos and caminhos.get("arestas") and caminhos.get("nós"):
        nodes_dict = {n["id"]: n for n in caminhos.get("nós", [])}
        for edge in caminhos.get("arestas", []):
            u_id = edge.get("de")
            v_id = edge.get("para")
            if u_id in nodes_dict and v_id in nodes_dict:
                u = nodes_dict[u_id]
                y = nodes_dict[v_id]
                if u.get("pavimento_id") == pavimento_id and y.get("pavimento_id") == pavimento_id:
                    active_arestas.append({
                        "de_id": u_id,
                        "de_coords": [u["y"], u["x"]],
                        "para_id": v_id,
                        "para_coords": [y["y"], y["x"]],
                        "tipo": edge.get("tipo", "caminho")
                    })
    active_arestas_json_str = json.dumps(active_arestas)

    # 5. Leaflet HTML/JS
    pins_json_str = json.dumps(pins)
    route_coords_json_str = json.dumps(route_coords)
    config_json_str = json.dumps(config)
    active_pin_ids_json_str = json.dumps(active_pin_ids)
    
    leaflet_html = f"""
    <!DOCTYPE html>
    <html>
    <head>
      <title>Planta Baixa</title>
      <meta charset="utf-8" />
      <meta name="viewport" content="width=device-width, initial-scale=1.0">
      
      <!-- Leaflet CSS e JS -->
      <link rel="stylesheet" href="https://unpkg.com/leaflet@1.9.4/dist/leaflet.css" />
      <script src="https://unpkg.com/leaflet@1.9.4/dist/leaflet.js"></script>
      
      <!-- Fullscreen Plugin CSS e JS -->
      <link rel="stylesheet" href="https://api.mapbox.com/mapbox.js/plugins/leaflet-fullscreen/v1.0.1/leaflet.fullscreen.css" />
      <script src="https://api.mapbox.com/mapbox.js/plugins/leaflet-fullscreen/v1.0.1/Leaflet.fullscreen.min.js"></script>
      
      <style>
        html, body {{
          margin: 0;
          padding: 0;
          height: 100%;
          width: 100%;
          overflow: hidden;
          background-color: transparent;
        }}
        #map {{
          height: 100%;
          width: 100%;
          margin: 0;
          padding: 0;
          background: #0e1117;
          border: 1px solid #464855;
          border-radius: 8px;
          box-sizing: border-box;
        }}
        /* Estilos premium para o pop-up e formulário do Modo Dev */
        .leaflet-popup-content-wrapper, .leaflet-popup-tip {{
          background: #1e1f25 !important;
          color: #ffffff !important;
          border: 1px solid #464855 !important;
          border-radius: 8px !important;
          box-shadow: 0 4px 15px rgba(0,0,0,0.5) !important;
        }}
        .dev-form {{
          display: flex;
          flex-direction: column;
          gap: 8px;
          min-width: 210px;
          font-family: 'Inter', sans-serif;
          padding: 4px;
        }}
        .dev-form label {{
          font-weight: bold;
          font-size: 11px;
          color: #a0a5b5;
          margin-bottom: 2px;
          display: block;
        }}
        .dev-form input[type="text"], .dev-form select {{
          background: #2a2b36;
          border: 1px solid #464855;
          color: #fff;
          padding: 5px 8px;
          border-radius: 4px;
          font-size: 12px;
          outline: none;
          width: 90%;
        }}
        .dev-form input[type="text"]:focus {{
          border-color: #4b9cff;
        }}
        .dev-form-row {{
          display: flex;
          align-items: center;
          gap: 6px;
          font-size: 12px;
        }}
        .dev-btn-group {{
          display: flex;
          justify-content: flex-end;
          gap: 8px;
          margin-top: 8px;
        }}
        .dev-btn {{
          padding: 6px 12px;
          border-radius: 4px;
          border: none;
          cursor: pointer;
          font-size: 11px;
          font-weight: bold;
          transition: background-color 0.2s;
        }}
        .dev-btn-save {{
          background-color: #2ecc71;
          color: #fff;
        }}
        .dev-btn-save:hover {{
          background-color: #27ae60;
        }}
        .dev-btn-cancel {{
          background-color: #e74c3c;
          color: #fff;
        }}
        .dev-btn-cancel:hover {{
          background-color: #c0392b;
        }}
      </style>
    </head>
    <body>
      <div id="map" style="height: 650px; width: 100%;"></div>
      <script>
        // Estados e controles do Modo Desenvolvedor (declarados no topo para evitar hoisting/ReferenceError)
        var devMode = false;
        var devState = {{
          lastNodeId: sessionStorage.getItem('dev_lastNodeId') || null,
          unsavedElements: JSON.parse(sessionStorage.getItem('dev_unsavedElements') || '[]')
        }};

        var activeRoute = {{
          origem: "{url_origem}",
          destino: "{url_destino}"
        }};

        var w = {w};
        var h = {h};
        var bounds = [[0, 0], [h, w]];

        // Configura o mapa simples limitando arrasto (maxBounds) somente dentro da planta
        var map = L.map('map', {{
          crs: L.CRS.Simple,
          minZoom: -2,
          maxZoom: 3,
          attributionControl: false,
          maxBounds: bounds,
          maxBoundsViscosity: 1.0
        }});
        
        // Carrega a imagem da planta
        var image = L.imageOverlay('{b64_image}', bounds).addTo(map);
        map.fitBounds(bounds);

        // Adiciona controle de Fullscreen no canto superior direito
        map.addControl(new L.Control.Fullscreen({{
          position: 'topright',
          title: {{
            'false': 'Ver em Tela Cheia',
            'true': 'Sair da Tela Cheia'
          }}
        }}));

        // Identificadores de controle do Streamlit passados ao JS
        var activePinIds = {active_pin_ids_json_str};
        var activeBuildingId = "{predio_id}";
        var floorId = {pavimento_id};
        var fullConfig = {config_json_str};

        // Layers para pins e malha (nós e arestas)
        var pinsLayer = L.layerGroup().addTo(map);
        var debugLayer = L.layerGroup(); // Começa oculto ou visível via controle do olho

        // Função unificada para desenhar todos os elementos (existentes e novos)
        window.redrawAllLayers = function() {{
          pinsLayer.clearLayers();
          debugLayer.clearLayers();

          var predio = fullConfig.predios.find(function(p) {{ return p.id === activeBuildingId; }});
          if (!predio) return;

          // Garante a existência da estrutura no JSON
          if (!predio.caminhos) predio.caminhos = {{ "nós": [], "arestas": [] }};
          if (!predio.caminhos.nós) predio.caminhos.nós = [];
          if (!predio.caminhos.arestas) predio.caminhos.arestas = [];
          if (!predio.pins) predio.pins = [];

          // 1. Desenha as Arestas
          predio.caminhos.arestas.forEach(function(edge) {{
            var deNode = predio.caminhos.nós.find(function(n) {{ return n.id === edge.de && n.pavimento_id === floorId; }});
            var paraNode = predio.caminhos.nós.find(function(n) {{ return n.id === edge.para && n.pavimento_id === floorId; }});
            if (deNode && paraNode) {{
              var polyline = L.polyline([[deNode.y, deNode.x], [paraNode.y, paraNode.x]], {{
                color: '#2ecc71',
                weight: 3,
                opacity: 0.5,
                dashArray: '5, 5'
              }}).addTo(debugLayer);

              polyline.bindTooltip("<b>Aresta:</b> " + edge.de + " ➔ " + edge.para + (edge.tipo && edge.tipo !== 'caminho' ? " (" + edge.tipo + ")" : ""), {{sticky: true}});

              // No modo desenvolvedor, permite excluir a aresta ao clicar nela
              polyline.on('click', function(e) {{
                if (devMode) {{
                  L.DomEvent.stopPropagation(e);
                  var popupContent = `
                    <div style="color:#fff; font-size:12px; font-family:sans-serif; padding:4px;">
                      <b>Aresta:</b> ${{edge.de}} ➔ ${{edge.para}}<br><br>
                      <button class="dev-btn dev-btn-cancel" onclick="window.removeEdge('${{edge.de}}', '${{edge.para}}')">Excluir Aresta</button>
                    </div>
                  `;
                  L.popup()
                    .setLatLng(e.latlng)
                    .setContent(popupContent)
                    .openOn(map);
                }}
              }});
            }}
          }});

          // 2. Desenha os Nós
          predio.caminhos.nós.forEach(function(node) {{
            if (node.pavimento_id !== floorId) return;

            var isUnsaved = devState.unsavedElements.includes(node.id);
            var color = isUnsaved ? "#ffd700" : "#8a2be2";

            var nodeIcon = L.divIcon({{
              className: 'debug-node',
              html: '<div style="background-color: ' + color + '; width: 8px; height: 8px; border-radius: 50%; opacity: 0.8; box-shadow: 0 0 3px rgba(0,0,0,0.5);"></div>',
              iconSize: [8, 8],
              iconAnchor: [4, 4]
            }});

            var marker = L.marker([node.y, node.x], {{
              icon: nodeIcon,
              draggable: devMode
            }}).addTo(debugLayer);

            marker.bindTooltip("<b>Nó:</b> " + node.id + "<br>" + (node.nome || ''), {{sticky: true}});

            // Atualiza coordenadas no JSON ao arrastar
            marker.on('dragend', function(e) {{
              var latlng = marker.getLatLng();
              node.x = Math.round(latlng.lng);
              node.y = Math.round(latlng.lat);
              console.log("📍 Nó " + node.id + " movido para: x=" + node.x + ", y=" + node.y);
              if (!devState.unsavedElements.includes(node.id)) {{
                devState.unsavedElements.push(node.id);
                sessionStorage.setItem('dev_unsavedElements', JSON.stringify(devState.unsavedElements));
              }}
              window.redrawAllLayers();
            }});

            // Clique no nó
            marker.on('click', function(e) {{
              L.DomEvent.stopPropagation(e);
              if (devMode) {{
                var connectBtnHtml = "";
                if (devState.lastNodeId && devState.lastNodeId !== node.id) {{
                  connectBtnHtml = '<button class="dev-btn" onclick="window.connectToLastNode(\\\'' + node.id + '\\\')" style="background-color:#2ecc71; color:#fff; font-size:10px;">Ligar a ' + devState.lastNodeId + '</button>';
                }}
                var editPopup = `
                  <div class="dev-form">
                    <div style="font-weight: bold; color: #ffd700; margin-bottom: 5px;">🛠️ Editar Nó</div>
                    <label>ID do Nó</label>
                    <input type="text" value="${{node.id}}" disabled style="background:#1e1f25; color:#888; border:1px solid #464855;">
                    
                    <label>Nome do Nó</label>
                    <input type="text" id="edit_node_nome" value="${{node.nome || ''}}">
                    
                    <div class="dev-form-row">
                      <label style="margin:0;">X:</label>
                      <input type="text" id="edit_node_x" value="${{node.x}}" style="width:55px;">
                      <label style="margin:0;">Y:</label>
                      <input type="text" id="edit_node_y" value="${{node.y}}" style="width:55px;">
                    </div>
                    
                    <div class="dev-btn-group">
                      <button class="dev-btn dev-btn-cancel" onclick="window.removeNode('${{node.id}}')" style="background-color:#e74c3c;">Excluir</button>
                      ${{connectBtnHtml}}
                      <button class="dev-btn" onclick="window.setLastNode('${{node.id}}')" style="background-color:#3498db; color:#fff;">Partir</button>
                      <button class="dev-btn dev-btn-save" onclick="window.updateNode('${{node.id}}')">Salvar</button>
                    </div>
                  </div>
                `;
                L.popup()
                  .setLatLng([node.y, node.x])
                  .setContent(editPopup)
                  .openOn(map);
              }}
            }});
          }});

          // 3. Desenha os Pins (Marcadores de Sala)
          predio.pins.forEach(function(pin) {{
            if (pin.pavimento_id !== floorId) return;

            var isActive = activePinIds.includes(pin.id);
            var isUnsaved = devState.unsavedElements.includes(pin.id);
            var color = isActive ? "#ff4b4b" : (isUnsaved ? "#e67e22" : "#4b9cff");
            var size = isActive ? "24px" : "16px";
            var border = isActive ? "3px solid white" : "2px solid white";

            var customIcon = L.divIcon({{
              className: 'custom-pin',
              html: '<div style="background-color: ' + color + '; width: ' + size + '; height: ' + size + '; border-radius: 50%; border: ' + border + '; box-shadow: 0 0 10px rgba(0,0,0,0.5);"></div>',
              iconSize: isActive ? [24, 24] : [16, 16],
              iconAnchor: isActive ? [12, 12] : [8, 8]
            }});

            var marker = L.marker([pin.y, pin.x], {{
              icon: customIcon,
              draggable: devMode
            }}).addTo(pinsLayer);

            var popupContent = "<b>📌 " + pin.sala + "</b><br>" + pin.descricao;
            if (!devMode) {{
              popupContent += "<br><br><div class='dev-btn-group' style='justify-content:center; gap:6px;'>" +
                              "<button class='dev-btn' style='background-color:#3498db; color:#fff; font-size:10px; padding:4px 8px; margin:0;' onclick='window.setRouteOrigin(\\\"" + pin.id + "\\\")'>Definir Origem</button>" +
                              "<button class='dev-btn' style='background-color:#2ecc71; color:#fff; font-size:10px; padding:4px 8px; margin:0;' onclick='window.setRouteDestination(\\\"" + pin.id + "\\\")'>Definir Destino</button>" +
                              "</div>";
            }}
            marker.bindPopup(popupContent);

            if (isActive && !devMode) {{
              if (activePinIds.indexOf(pin.id) === 0) {{
                marker.openPopup();
                map.setView([pin.y, pin.x], 1);
              }}
            }}

            // Atualiza coordenadas no JSON ao arrastar
            marker.on('dragend', function(e) {{
              var latlng = marker.getLatLng();
              pin.x = Math.round(latlng.lng);
              pin.y = Math.round(latlng.lat);
              console.log("📍 Pin " + pin.id + " movido para: x=" + pin.x + ", y=" + pin.y);
              if (!devState.unsavedElements.includes(pin.id)) {{
                devState.unsavedElements.push(pin.id);
                sessionStorage.setItem('dev_unsavedElements', JSON.stringify(devState.unsavedElements));
              }}
              window.redrawAllLayers();
            }});

            // Clique no Pin no modo Dev
            marker.on('click', function(e) {{
              if (devMode) {{
                L.DomEvent.stopPropagation(e);
                var editPopup = `
                  <div class="dev-form">
                    <div style="font-weight: bold; color: #4b9cff; margin-bottom: 5px;">🛠️ Editar Pin</div>
                    <label>ID do Pin</label>
                    <input type="text" value="${{pin.id}}" disabled style="background:#1e1f25; color:#888; border:1px solid #464855;">
                    
                    <label>Nome da Sala</label>
                    <input type="text" id="edit_pin_sala" value="${{pin.sala || ''}}">
                    
                    <label>Descrição</label>
                    <input type="text" id="edit_pin_desc" value="${{pin.descricao || ''}}">
                    
                    <div class="dev-form-row">
                      <label style="margin:0;">X:</label>
                      <input type="text" id="edit_pin_x" value="${{pin.x}}" style="width:55px;">
                      <label style="margin:0;">Y:</label>
                      <input type="text" id="edit_pin_y" value="${{pin.y}}" style="width:55px;">
                    </div>
                    
                    <div class="dev-btn-group">
                      <button class="dev-btn dev-btn-cancel" onclick="window.removePin('${{pin.id}}')" style="background-color:#e74c3c;">Excluir</button>
                      <button class="dev-btn dev-btn-cancel" onclick="map.closePopup();">Fechar</button>
                      <button class="dev-btn dev-btn-save" onclick="window.updatePin('${{pin.id}}')">Salvar</button>
                    </div>
                  </div>
                `;
                L.popup()
                  .setLatLng([pin.y, pin.x])
                  .setContent(editPopup)
                  .openOn(map);
              }}
            }});
          }});
        }};

        // Inicializa o desenho do mapa
        window.redrawAllLayers();

        // Callbacks de manipulação do estado em tempo de execução
        window.setRouteOrigin = function(pinId) {{
          activeRoute.origem = pinId;
          map.closePopup();
          fetch('http://localhost:8099/set_route?origem=' + pinId + '&destino=' + activeRoute.destino)
            .then(function() {{
                // Envia uma mensagem ou recarrega a página pai do Streamlit para aplicar as rotas e seletores instantaneamente
                if (window.parent) {{
                    window.parent.postMessage({{type: 'streamlit:render'}}, '*');
                }}
            }})
            .catch(function(err) {{ console.error("Erro ao definir origem:", err); }});
        }};

        window.setRouteDestination = function(pinId) {{
          activeRoute.destino = pinId;
          map.closePopup();
          fetch('http://localhost:8099/set_route?origem=' + activeRoute.origem + '&destino=' + pinId)
            .then(function() {{
                if (window.parent) {{
                    window.parent.postMessage({{type: 'streamlit:render'}}, '*');
                }}
            }})
            .catch(function(err) {{ console.error("Erro ao definir destino:", err); }});
        }};

        window.setLastNode = function(id) {{
          devState.lastNodeId = id;
          sessionStorage.setItem('dev_lastNodeId', id);
          console.log("📌 Nó de partida definido como: " + id);
          alert("Nó de partida definido como: " + id);
          map.closePopup();
        }};

        window.connectToLastNode = function(id) {{
          if (!devState.lastNodeId || devState.lastNodeId === id) return;
          var predio = fullConfig.predios.find(function(p) {{ return p.id === activeBuildingId; }});
          if (predio) {{
            if (!predio.caminhos) predio.caminhos = {{ "nós": [], "arestas": [] }};
            if (!predio.caminhos.arestas) predio.caminhos.arestas = [];
            
            // Verifica se a aresta já existe
            var exists = predio.caminhos.arestas.some(function(edge) {{
              return (edge.de === devState.lastNodeId && edge.para === id) || 
                     (edge.de === id && edge.para === devState.lastNodeId);
            }});
            
            if (!exists) {{
              predio.caminhos.arestas.push({{
                de: devState.lastNodeId,
                para: id
              }});
              console.log("🔗 Aresta criada: " + devState.lastNodeId + " -> " + id);
            }}
            
            // Define o nó recém conectado como o novo nó de partida para permitir encadeamento fácil
            devState.lastNodeId = id;
            sessionStorage.setItem('dev_lastNodeId', id);
            window.redrawAllLayers();
          }}
          map.closePopup();
        }};

        window.removeNode = function(id) {{
          var predio = fullConfig.predios.find(function(p) {{ return p.id === activeBuildingId; }});
          if (predio && predio.caminhos && predio.caminhos.nós) {{
            predio.caminhos.nós = predio.caminhos.nós.filter(function(n) {{ return n.id !== id; }});
            if (predio.caminhos.arestas) {{
              predio.caminhos.arestas = predio.caminhos.arestas.filter(function(a) {{ return a.de !== id && a.para !== id; }});
            }}
            if (devState.lastNodeId === id) {{
              devState.lastNodeId = null;
              sessionStorage.removeItem('dev_lastNodeId');
            }}
            devState.unsavedElements = devState.unsavedElements.filter(function(x) {{ return x !== id; }});
            sessionStorage.setItem('dev_unsavedElements', JSON.stringify(devState.unsavedElements));
            window.redrawAllLayers();
            console.log("❌ Nó removido: " + id);
          }}
          map.closePopup();
        }};

        window.removePin = function(id) {{
          var predio = fullConfig.predios.find(function(p) {{ return p.id === activeBuildingId; }});
          if (predio && predio.pins) {{
            predio.pins = predio.pins.filter(function(p) {{ return p.id !== id; }});
            devState.unsavedElements = devState.unsavedElements.filter(function(x) {{ return x !== id; }});
            sessionStorage.setItem('dev_unsavedElements', JSON.stringify(devState.unsavedElements));
            window.redrawAllLayers();
            console.log("❌ Pin removido: " + id);
          }}
          map.closePopup();
        }};

        window.removeEdge = function(de, para) {{
          var predio = fullConfig.predios.find(function(p) {{ return p.id === activeBuildingId; }});
          if (predio && predio.caminhos && predio.caminhos.arestas) {{
            predio.caminhos.arestas = predio.caminhos.arestas.filter(function(a) {{
              return !(a.de === de && a.para === para);
            }});
            window.redrawAllLayers();
            console.log("❌ Aresta removida: " + de + " -> " + para);
          }}
          map.closePopup();
        }};

        window.updateNode = function(id) {{
          var nome = document.getElementById("edit_node_nome").value.trim();
          var x = parseInt(document.getElementById("edit_node_x").value.trim());
          var y = parseInt(document.getElementById("edit_node_y").value.trim());
          
          var predio = fullConfig.predios.find(function(p) {{ return p.id === activeBuildingId; }});
          if (predio && predio.caminhos && predio.caminhos.nós) {{
            var node = predio.caminhos.nós.find(function(n) {{ return n.id === id; }});
            if (node) {{
              node.nome = nome;
              if (!isNaN(x)) node.x = x;
              if (!isNaN(y)) node.y = y;
              if (!devState.unsavedElements.includes(id)) {{
                devState.unsavedElements.push(id);
                sessionStorage.setItem('dev_unsavedElements', JSON.stringify(devState.unsavedElements));
              }}
              window.redrawAllLayers();
              console.log("✔️ Nó atualizado:", node);
            }}
          }}
          map.closePopup();
        }};

        window.updatePin = function(id) {{
          var sala = document.getElementById("edit_pin_sala").value.trim();
          var desc = document.getElementById("edit_pin_desc").value.trim();
          var x = parseInt(document.getElementById("edit_pin_x").value.trim());
          var y = parseInt(document.getElementById("edit_pin_y").value.trim());
          
          var predio = fullConfig.predios.find(function(p) {{ return p.id === activeBuildingId; }});
          if (predio && predio.pins) {{
            var pin = predio.pins.find(function(p) {{ return p.id === id; }});
            if (pin) {{
              pin.sala = sala;
              pin.descricao = desc;
              if (!isNaN(x)) pin.x = x;
              if (!isNaN(y)) pin.y = y;
              if (!devState.unsavedElements.includes(id)) {{
                devState.unsavedElements.push(id);
                sessionStorage.setItem('dev_unsavedElements', JSON.stringify(devState.unsavedElements));
              }}
              window.redrawAllLayers();
              console.log("✔️ Pin atualizado:", pin);
            }}
          }}
          map.closePopup();
        }};

        // Criação de botão customizado de controle de visibilidade da malha
        var MeshToggleControl = L.Control.extend({{
          options: {{
            position: 'topright'
          }},
          onAdd: function (map) {{
            var container = L.DomUtil.create('div', 'leaflet-bar leaflet-control leaflet-custom-control');
            container.style.backgroundColor = '#1e1f25';
            container.style.width = '34px';
            container.style.height = '34px';
            container.style.cursor = 'pointer';
            container.style.display = 'flex';
            container.style.alignItems = 'center';
            container.style.justifyContent = 'center';
            container.style.borderRadius = '4px';
            container.style.border = '1px solid #464855';
            container.style.transition = 'all 0.2s';
            container.style.opacity = '0.5';
            container.title = "Exibir Malha de Caminhos";

            // Ícone do Olho para Visibilidade
            container.innerHTML = '<span style="font-size: 16px; line-height: 1; filter: grayscale(100%);">👁️</span>';

            var isVisible = false;
            container.onclick = function(e) {{
              L.DomEvent.stopPropagation(e);
              if (isVisible) {{
                map.removeLayer(debugLayer);
                container.style.opacity = '0.5';
              }} else {{
                map.addLayer(debugLayer);
                container.style.opacity = '1.0';
              }}
              isVisible = !isVisible;
            }};
            
            container.onmouseover = function() {{
              container.style.backgroundColor = '#2a2b36';
            }};
            container.onmouseout = function() {{
              container.style.backgroundColor = '#1e1f25';
            }};

            return container;
          }}
        }});

        map.addControl(new MeshToggleControl());


        var exportCtrlInstance = null;

        var DevModeControl = L.Control.extend({{
          options: {{
            position: 'topright'
          }},
          onAdd: function (map) {{
            var container = L.DomUtil.create('div', 'leaflet-bar leaflet-control leaflet-custom-control');
            container.style.backgroundColor = '#1e1f25';
            container.style.width = '34px';
            container.style.height = '34px';
            container.style.cursor = 'pointer';
            container.style.display = 'flex';
            container.style.alignItems = 'center';
            container.style.justifyContent = 'center';
            container.style.borderRadius = '4px';
            container.style.border = '1px solid #464855';
            container.style.transition = 'all 0.2s';
            container.style.opacity = '0.5';
            container.title = "Modo Desenvolvedor (Editar Mapa)";

            container.innerHTML = '<span style="font-size: 16px; line-height: 1;">🛠️</span>';

            container.onclick = function(e) {{
              L.DomEvent.stopPropagation(e);
              devMode = !devMode;
              if (devMode) {{
                container.style.opacity = '1.0';
                container.style.borderColor = '#2ecc71';
                container.style.boxShadow = '0 0 8px rgba(46, 204, 113, 0.6)';
                map.getContainer().style.cursor = 'crosshair';
                map.addLayer(debugLayer); // Habilita a visualização da malha para poder editá-la
                window.redrawAllLayers();
                if (exportCtrlInstance) {{
                  exportCtrlInstance.show();
                }}
              }} else {{
                container.style.opacity = '0.5';
                container.style.borderColor = '#464855';
                container.style.boxShadow = 'none';
                map.getContainer().style.cursor = '';
                window.redrawAllLayers();
                if (exportCtrlInstance) {{
                  exportCtrlInstance.hide();
                }}
              }}
            }};

            container.onmouseover = function() {{
              container.style.backgroundColor = '#2a2b36';
            }};
            container.onmouseout = function() {{
              container.style.backgroundColor = '#1e1f25';
            }};

            return container;
          }}
        }});

        var SaveControl = L.Control.extend({{
          options: {{
            position: 'topright'
          }},
          onAdd: function (map) {{
            var container = L.DomUtil.create('div', 'leaflet-bar leaflet-control leaflet-custom-control');
            container.style.backgroundColor = '#1e1f25';
            container.style.width = '34px';
            container.style.height = '34px';
            container.style.cursor = 'pointer';
            container.style.display = 'none';
            container.style.alignItems = 'center';
            container.style.justifyContent = 'center';
            container.style.borderRadius = '4px';
            container.style.border = '1px solid #464855';
            container.style.transition = 'all 0.2s';
            container.title = "Salvar alterações no Banco de Dados";

            container.innerHTML = '<span style="font-size: 16px; line-height: 1;">💾</span>';

            container.onclick = function(e) {{
              L.DomEvent.stopPropagation(e);
              window.saveConfigToDb();
            }};

            container.onmouseover = function() {{
              container.style.backgroundColor = '#2a2b36';
            }};
            container.onmouseout = function() {{
              container.style.backgroundColor = '#1e1f25';
            }};

            this._container = container;
            return container;
          }},
          show: function() {{
            if (this._container) this._container.style.display = 'flex';
          }},
          hide: function() {{
            if (this._container) this._container.style.display = 'none';
          }}
        }});

        var devModeCtrl = new DevModeControl();
        var exportCtrl = new SaveControl();
        map.addControl(devModeCtrl);
        map.addControl(exportCtrl);
        exportCtrlInstance = exportCtrl;

        // Callback para salvar elementos temporários criados em tela
        window.saveDevElement = function(x, y) {{
          var elemType = document.querySelector('input[name="elem_type"]:checked').value;
          var floor = {pavimento_id};
          var predioId = "{predio_id}";
          
          var predio = fullConfig.predios.find(function(p) {{ return p.id === predioId; }});
          if (!predio) return;

          if (!predio.caminhos) predio.caminhos = {{ "nós": [], "arestas": [] }};
          if (!predio.caminhos.nós) predio.caminhos.nós = [];
          if (!predio.caminhos.arestas) predio.caminhos.arestas = [];
          if (!predio.pins) predio.pins = [];

          if (elemType === "node") {{
            var id = document.getElementById("dev_node_id").value.trim() || ("no_" + Date.now());
            var nome = document.getElementById("dev_node_nome").value.trim();
            var connect = document.getElementById("dev_node_connect") ? document.getElementById("dev_node_connect").checked : false;

            var newNode = {{
              id: id,
              pavimento_id: floor,
              x: x,
              y: y,
              nome: nome
            }};

            predio.caminhos.nós.push(newNode);

            if (connect && devState.lastNodeId) {{
              var newEdge = {{
                de: devState.lastNodeId,
                para: id
              }};
              predio.caminhos.arestas.push(newEdge);
            }}

            devState.lastNodeId = id;
            sessionStorage.setItem('dev_lastNodeId', id);
            if (!devState.unsavedElements.includes(id)) {{
              devState.unsavedElements.push(id);
              sessionStorage.setItem('dev_unsavedElements', JSON.stringify(devState.unsavedElements));
            }}
            console.log("✔️ Novo Nó adicionado diretamente ao fullConfig:", newNode);
          }} else {{
            var id = document.getElementById("dev_pin_id").value.trim() || ("pin_" + Date.now());
            var sala = document.getElementById("dev_pin_sala").value.trim() || "Nova Sala";
            var desc = document.getElementById("dev_pin_desc").value.trim();

            var newPin = {{
              id: id,
              predio_id: predioId,
              pavimento_id: floor,
              sala: sala,
              x: x,
              y: y,
              descricao: desc
            }};

            predio.pins.push(newPin);
            if (!devState.unsavedElements.includes(id)) {{
              devState.unsavedElements.push(id);
              sessionStorage.setItem('dev_unsavedElements', JSON.stringify(devState.unsavedElements));
            }}
            console.log("✔️ Novo Pin adicionado diretamente ao fullConfig:", newPin);
          }}

          window.redrawAllLayers();
          map.closePopup();
        }};

        // Salva as configurações atualizadas no banco e no JSON físico do projeto
        window.saveConfigToDb = function() {{
          fetch('http://localhost:8099/save_config', {{
            method: 'POST',
            headers: {{
              'Content-Type': 'application/json'
            }},
            body: JSON.stringify(fullConfig)
          }})
          .then(function(response) {{
            if (response.ok) {{
              devState.unsavedElements = [];
              sessionStorage.removeItem('dev_unsavedElements');
              window.redrawAllLayers();
              alert("💾 Configurações salvas com sucesso no Banco de Dados SQLite e no arquivo JSON uploads/map_config_TEMPLATE.json!");
            }} else {{
              alert("❌ Erro ao salvar configurações no servidor.");
            }}
          }})
          .catch(function(error) {{
            console.error(error);
            alert("❌ Erro de rede ao tentar salvar no servidor backend.");
          }});
        }};

        // Desenha a rota de pathfinding se houver coordenadas válidas
        var routeCoords = {route_coords_json_str};
        if (routeCoords.length > 1) {{
          var polyline = L.polyline(routeCoords, {{
            color: '#ff4b4b',
            weight: 6,
            opacity: 0.8,
            dashArray: '10, 10',
            lineJoin: 'round'
          }}).addTo(map);
          
          // Adiciona popup informativo de distância ao clicar na rota
          polyline.bindPopup("<b>🚶 Rota Interna Calculada</b><br>Distância total estimada: <b>" + "{route_distance_meters:.1f}" + " m</b>");
          
          // Enquadra a visão do mapa para englobar toda a rota percorrida
          map.fitBounds(polyline.getBounds());
        }}

        // Ações de clique no mapa
        map.on('click', function(e) {{
          var coord = e.latlng;
          var x = Math.round(coord.lng);
          var y = Math.round(coord.lat);
          
          if (x >= 0 && x <= w && y >= 0 && y <= h) {{
            var floor = {pavimento_id};
            
            if (devMode) {{
              var tempIdNode = "no_" + Date.now();
              var tempIdPin = "pin_" + Date.now();
              
              var popupContent = `
                <div class="dev-form">
                  <div style="font-weight: bold; margin-bottom: 5px; color: #4b9cff;">🛠️ Criar Elemento</div>
                  
                  <div class="dev-form-row" style="margin-bottom: 6px;">
                    <input type="radio" id="type_node" name="elem_type" value="node" checked onchange="document.getElementById('node_fields').style.display='flex'; document.getElementById('pin_fields').style.display='none';">
                    <label for="type_node" style="margin:0; cursor:pointer; color:#fff;">Nó</label>
                    
                    <input type="radio" id="type_pin" name="elem_type" value="pin" onchange="document.getElementById('node_fields').style.display='none'; document.getElementById('pin_fields').style.display='flex';">
                    <label for="type_pin" style="margin:0; cursor:pointer; color:#fff;">Pin (Sala)</label>
                  </div>
                  
                  <!-- Campos do Nó -->
                  <div id="node_fields" style="display: flex; flex-direction: column; gap: 8px;">
                    <label>ID do Nó</label>
                    <input type="text" id="dev_node_id" value="${{tempIdNode}}">
                    
                    <label>Nome do Nó</label>
                    <input type="text" id="dev_node_nome" placeholder="Ex: Corredor Ala A" value="">
                    
                    <div class="dev-form-row" style="margin-top: 4px;">
                      <input type="checkbox" id="dev_node_connect" ${{devState.lastNodeId ? 'checked' : 'disabled'}}>
                      <label for="dev_node_connect" style="margin:0; cursor:pointer; font-size:11px;">Conectar ao nó anterior (${{devState.lastNodeId || 'Nenhum'}})</label>
                    </div>
                  </div>
                  
                  <!-- Campos do Pin -->
                  <div id="pin_fields" style="display: none; flex-direction: column; gap: 8px;">
                    <label>ID do Pin</label>
                    <input type="text" id="dev_pin_id" value="${{tempIdPin}}">
                    
                    <label>Nome da Sala / Local</label>
                    <input type="text" id="dev_pin_sala" placeholder="Ex: Sala 102" value="">
                    
                    <label>Descrição</label>
                    <input type="text" id="dev_pin_desc" placeholder="Ex: Suporte Técnico" value="">
                  </div>
                  
                  <div class="dev-btn-group">
                    <button class="dev-btn dev-btn-cancel" onclick="map.closePopup();">Cancelar</button>
                    <button class="dev-btn dev-btn-save" onclick="window.saveDevElement(${{x}}, ${{y}})">Adicionar</button>
                  </div>
                </div>
              `;
              
              L.popup()
                .setLatLng(coord)
                .setContent(popupContent)
                .openOn(map);
            }} else {{
              console.log("📍 Coordenada Clicada -> x: " + x + ", y: " + y);
              
              // Log individual do Nó
              console.log("📦 [JSON NÓ]:", '{{\"id\": \"n_novo\", \"pavimento_id\": ' + floor + ', \"x\": ' + x + ', \"y\": ' + y + ', \"nome\": \"\"}}');
              
              // Log individual do Pin
              console.log("📌 [JSON PIN]:", '{{\"id\": \"pin_novo\", \"pavimento_id\": ' + floor + ', \"sala\": \"\", \"x\": ' + x + ', \"y\": ' + y + ', \"descricao\": \"\"}}');
              
              // Encontra o nó mais próximo no pavimento ativo
              var nearestNode = null;
              var minDist = Infinity;
              if (typeof activeNodes !== "undefined" && activeNodes.length > 0) {{
                activeNodes.forEach(function(node) {{
                  var dx = node.x - x;
                  var dy = node.y - y;
                  var dist = Math.sqrt(dx*dx + dy*dy);
                  if (dist < minDist) {{
                    minDist = dist;
                    nearestNode = node;
                  }}
                }});
              }}
              
              // Log sugerido de Aresta conectando ao nó mais próximo
              if (nearestNode) {{
                var distRound = Math.round(minDist);
                console.log("🔗 [JSON ARESTA] (Nó mais próximo: " + nearestNode.id + ", dist: " + distRound + "px):", 
                            '{{\"de\": \"' + nearestNode.id + '\", \"para\": \"n_novo\"}}');
              }}
            }}
          }}
        }});
      </script>
    </body>
    </html>
    """
    
    leaflet_html += f"\n<!-- key: map_{url_origem}_{url_destino}_{pavimento_id} -->"
    components.html(leaflet_html, height=670)


def render_donations_page():
    """Renderiza a página de Doação & Redistribuição de Máquinas."""
    from src.database import get_donations_data, sync_donations_from_excel
    
    st.title("🖥️ Sistema de Doação & Redistribuição de Máquinas")
    st.write("Acompanhe o inventário de equipamentos destinados a doação, redistribuição, garantia ou baixados.")
    
    from src.config import DONATIONS_FILE_PATH
    EXCEL_PATH = str(DONATIONS_FILE_PATH)

    
    # Cabeçalho com botão de sincronização
    col1, col2 = st.columns([3, 1])
    with col1:
        st.info("Os dados exibidos abaixo são sincronizados a partir da planilha oficial no SharePoint.")
    with col2:
        if st.button("🔄 Sincronizar Planilha", type="primary", use_container_width=True):
            with st.spinner("Lendo dados da planilha Excel..."):
                try:
                    sync_donations_from_excel(EXCEL_PATH)
                    st.toast("✅ Dados sincronizados com sucesso!", icon="💾")
                    st.rerun()
                except Exception as e:
                    st.error(f"Erro ao sincronizar planilha: {e}")
                    
    # Carrega dados do SQLite
    df = get_donations_data()
    
    if df.empty:
        st.warning("⚠️ Nenhum dado encontrado no cache local. Por favor, clique em 'Sincronizar Planilha' para carregar os registros.")
        return
        
    # Extrai o Ano da data de movimentação para fins de filtros e gráficos
    df['Ano'] = pd.to_datetime(df['data_movimentacao'], errors='coerce').dt.year
    df['Ano'] = df['Ano'].fillna("Sem Data").astype(str).str.replace(".0", "", regex=False)
        
    # Barra lateral de filtros e ferramentas
    st.sidebar.title("🖥️ Painel de Controle")
    st.sidebar.subheader("🔍 Filtros de Equipamentos")
    
    # Filtro de Movimentação
    mov_options = ["Todos"] + sorted(list(df['tipo_movimentacao'].unique()))
    selected_mov = st.sidebar.selectbox("Tipo de Movimentação", mov_options, key="donations_mov")
    
    # Filtro de Equipamento
    equip_options = ["Todos"] + sorted(list(df['equipamento'].unique()))
    selected_equip = st.sidebar.selectbox("Tipo de Equipamento", equip_options, key="donations_equip")
    
    # Filtro de Modelo
    model_options = ["Todos"] + sorted(list(df['modelo'].unique()))
    selected_model = st.sidebar.selectbox("Modelo", model_options, key="donations_model")
    
    # Filtro de Ano
    year_options = ["Todos"] + sorted(list(df['Ano'].unique()), reverse=True)
    selected_year = st.sidebar.selectbox("Ano da Movimentação", year_options, key="donations_year")

    
    # Filtro de SSD
    ssd_options = ["Todos"] + sorted(list(df['ssd'].unique()))
    selected_ssd = st.sidebar.selectbox("SSD", ssd_options, key="donations_ssd")
    
    # Filtro por Busca Geral (Patrimônio, Modelo ou Chamado)
    search_query = st.sidebar.text_input("🔎 Buscar (Patrimônio, Modelo, Chamado)", "", key="donations_search").strip()

    st.sidebar.markdown("---")
    st.sidebar.subheader("📋 Gerador de Texto (Preparo)")
    
    # Filtra as datas disponíveis no banco (limpando vazias)
    valid_dates = df[df['data_movimentacao'] != '']['data_movimentacao'].unique()
    valid_dates = sorted(list(valid_dates), reverse=True)
    
    # Função para formatar a data de YYYY-MM-DD para DD/MM/YYYY no selectbox
    def format_date_br(date_str):
        from datetime import datetime
        try:
            return datetime.strptime(date_str, "%Y-%m-%d").strftime("%d/%m/%Y")
        except:
            return date_str
            
    selected_date_str = st.sidebar.selectbox("Selecione a Data de Preparo", valid_dates, format_func=format_date_br)
    generate_btn = st.sidebar.button("📝 Gerar Texto do Chamado", use_container_width=True)

    @st.dialog("📋 Texto de Preparo de Chamado", width="large")
    def show_preparo_text(date_str, df_all):
        # Filtra os dados daquela data
        df_date = df_all[df_all['data_movimentacao'] == date_str]
        
        if df_date.empty:
            st.warning("Nenhum equipamento encontrado nesta data.")
            return
            
        from datetime import datetime
        try:
            # Tenta formatar a data de YYYY-MM-DD para DD/MM/YYYY
            dt_obj = datetime.strptime(date_str, "%Y-%m-%d")
            formatted_date = dt_obj.strftime("%d/%m/%Y")
        except:
            formatted_date = date_str
            
        # Constrói o assunto do chamado dinamicamente com base nas movimentações do dia
        movs = sorted(list(df_date['tipo_movimentacao'].unique()))
        movs_clean = [m.strip().capitalize() for m in movs if m.strip()]
        if len(movs_clean) == 1:
            movs_str = movs_clean[0]
        elif len(movs_clean) > 1:
            movs_str = ", ".join(movs_clean[:-1]) + " e " + movs_clean[-1]
        else:
            movs_str = "Movimentação"
            
        subject = f"[DOAÇÃO] - Preparação de {movs_str} de equipamentos do dia {formatted_date}"
        
        st.write("📋 **Assunto do Chamado:**")
        st.code(subject, language="text")
        st.markdown("---")
            
        html_parts = []
        html_parts.append("<div style='font-family: Arial, Helvetica, sans-serif; color: #000000; line-height: 1.5;'>")
        html_parts.append("<p>Prezados, boa tarde.</p>")
        html_parts.append(f"<p>Na tarde de hoje (<strong>{formatted_date}</strong>), preparamos os seguintes equipamentos, sendo eles:</p>")
        
        # Agrupa os equipamentos por tipo de movimentação
        for mov_type, grp in df_date.groupby('tipo_movimentacao'):
            html_parts.append(f"<p style='margin-top: 20px; margin-bottom: 8px;'><strong>🔹 Equipamentos para {mov_type.upper()}:</strong></p>")
            
            # Verifica se há valores reais nas colunas opcionais para este grupo específico
            has_ssd = grp['ssd'].astype(str).str.strip().any()
            has_obs = grp['motivo_baixa'].astype(str).str.strip().any()
            
            # Tabela HTML com bordas explícitas para garantir que copie e cole com formatação no OTRS/Outlook
            table_html = [
                "<table border='2' cellpadding='6' cellspacing='0' style='border-collapse: collapse; width: 100%; border: 2px solid #cccccc; font-family: Arial, Helvetica, sans-serif; font-size: 13px; color: #000000;'>"
            ]
            # Cabeçalho da tabela - cores azul corporativo com texto branco, aplicados direto no th
            th_style = "background-color: #2f5597; border: 2px solid #cccccc; padding: 6px 10px; text-align: left;"
            
            headers_html = [
                f"<tr>",
                f"<th style='{th_style}'><span style=\"color:#ffffff\">Patrimônio</span></th>",
                f"<th style='{th_style}'><span style=\"color:#ffffff\">Modelo</span></th>",
                f"<th style='{th_style}'><span style=\"color:#ffffff\">Serial Number PC</span></th>",
                f"<th style='{th_style}'><span style=\"color:#ffffff\">Equipamento</span></th>"
            ]
            if has_ssd:
                headers_html.append(f"<th style='{th_style}'><span style=\"color:#ffffff\">SSD</span></th>")
            if has_obs:
                headers_html.append(f"<th style='{th_style}'><span style=\"color:#ffffff\">Motivo/Obs</span></th>")
            headers_html.append("</tr>")
            
            table_html.append("".join(headers_html))
            
            # Linhas da tabela
            for idx, (_, row) in enumerate(grp.iterrows()):
                pat = str(row.get('patrimonio', '')).strip()
                mod = str(row.get('modelo', '')).strip()
                ser = str(row.get('serial_number', '')).strip()
                eqp = str(row.get('equipamento', '')).strip()
                ssd = str(row.get('ssd', '')).strip()
                obs = str(row.get('motivo_baixa', '')).strip()
                
                # Zebra striping (linhas alternadas com azul claro do Excel)
                bg_style = "background-color: #d9e1f2;" if idx % 2 == 1 else ""
                
                row_html = [
                    f"<tr>",
                    f"<td style='border: 2px solid #cccccc; padding: 6px 10px; {bg_style}'>{pat}</td>",
                    f"<td style='border: 2px solid #cccccc; padding: 6px 10px; {bg_style}'>{mod}</td>",
                    f"<td style='border: 2px solid #cccccc; padding: 6px 10px; {bg_style}'>{ser}</td>",
                    f"<td style='border: 2px solid #cccccc; padding: 6px 10px; {bg_style}'>{eqp}</td>"
                ]
                if has_ssd:
                    row_html.append(f"<td style='border: 2px solid #cccccc; padding: 6px 10px; {bg_style}'>{ssd}</td>")
                if has_obs:
                    row_html.append(f"<td style='border: 2px solid #cccccc; padding: 6px 10px; {bg_style}'>{obs}</td>")
                row_html.append("</tr>")
                
                table_html.append("".join(row_html))
 
            table_html.append("</table>")
            html_parts.append("".join(table_html))
            
        html_parts.append("</div>")
        full_html = "\n".join(html_parts)
        
        st.write("💡 Selecione o texto abaixo com o mouse, copie (Ctrl+C) e cole diretamente no chamado do OTRS:")
        st.markdown(
            f'<div style="background-color: #ffffff; padding: 20px; border-radius: 6px; border: 1px solid #dddddd; max-height: 400px; overflow-y: auto;">{full_html}</div>', 
            unsafe_allow_html=True
        )
        
        st.markdown("---")
        st.write("💻 Ou copie o código-fonte HTML abaixo (clique no botão **'Código-Fonte'** no OTRS e cole):")
        st.code(full_html, language="html")

    if generate_btn:
        show_preparo_text(selected_date_str, df)

    # Aplicação dos filtros
    df_filtered = df.copy()
    if selected_mov != "Todos":
        df_filtered = df_filtered[df_filtered['tipo_movimentacao'] == selected_mov]
    if selected_equip != "Todos":
        df_filtered = df_filtered[df_filtered['equipamento'] == selected_equip]
    if selected_model != "Todos":
        df_filtered = df_filtered[df_filtered['modelo'] == selected_model]
    if selected_year != "Todos":
        df_filtered = df_filtered[df_filtered['Ano'] == selected_year]
    if selected_ssd != "Todos":
        df_filtered = df_filtered[df_filtered['ssd'] == selected_ssd]


    if search_query:
        query_lower = search_query.lower()
        df_filtered = df_filtered[
            df_filtered['patrimonio'].str.lower().str.contains(query_lower) |
            df_filtered['modelo'].str.lower().str.contains(query_lower) |
            df_filtered['chamado'].str.lower().str.contains(query_lower)
        ]


        
    # KPIs rápidos
    st.markdown("---")
    kpi_col1, kpi_col2, kpi_col3, kpi_col4, kpi_col5 = st.columns(5)
    
    total_equip = len(df_filtered)
    # Conta doações (case-insensitive)
    doados = len(df_filtered[df_filtered['tipo_movimentacao'].str.lower() == 'doação'])
    # Conta redistribuições (case-insensitive)
    redistribuicoes = len(df_filtered[df_filtered['tipo_movimentacao'].str.lower() == 'redistribuição'])
    # Conta baixas (case-insensitive)
    baixas = len(df_filtered[df_filtered['tipo_movimentacao'].str.lower() == 'baixa'])
    # Conta garantias (case-insensitive)
    garantias = len(df_filtered[df_filtered['tipo_movimentacao'].str.lower() == 'garantia'])
    
    kpi_col1.metric("Todos os Equipamentos", total_equip)
    kpi_col2.metric("Doações", doados)
    kpi_col3.metric("Redistribuições", redistribuicoes)
    kpi_col4.metric("Baixas", baixas)
    kpi_col5.metric("Garantias", garantias)

    
    st.markdown("---")
    
    # Gráficos
    g_col1, g_col2 = st.columns(2)
    
    with g_col1:
        st.subheader("📊 Distribuição por Movimentação")
        if not df_filtered.empty:
            mov_counts = df_filtered['tipo_movimentacao'].value_counts().reset_index()
            mov_counts.columns = ['Movimentação', 'Quantidade']
            st.bar_chart(data=mov_counts, x='Movimentação', y='Quantidade', use_container_width=True)
        else:
            st.info("Sem dados para exibir o gráfico.")
            
    with g_col2:
        st.subheader("📅 Histórico de Movimentações por Ano")
        if not df_filtered.empty:
            # Converte a data_movimentacao para obter o ano
            df_filtered['Ano'] = pd.to_datetime(df_filtered['data_movimentacao'], errors='coerce').dt.year
            df_filtered['Ano'] = df_filtered['Ano'].fillna("Sem Data").astype(str).str.replace(".0", "", regex=False)
            
            ano_counts = df_filtered.groupby(['Ano', 'tipo_movimentacao']).size().unstack(fill_value=0)
            st.bar_chart(ano_counts, use_container_width=True)
        else:
            st.info("Sem dados para exibir o gráfico.")
            
    # Tabela principal
    st.markdown("---")
    st.subheader("📋 Detalhamento dos Equipamentos")
    
    st.dataframe(
        df_filtered,
        column_config={
            "patrimonio": st.column_config.TextColumn("Patrimônio"),
            "modelo": st.column_config.TextColumn("Modelo"),
            "serial_number": st.column_config.TextColumn("Número de Série"),
            "equipamento": st.column_config.TextColumn("Equipamento"),
            "tipo_movimentacao": st.column_config.TextColumn("Movimentação"),
            "data_movimentacao": st.column_config.DateColumn("Data da Movimentação", format="DD/MM/YYYY"),
            "chamado": st.column_config.TextColumn("Chamado relacionado"),
            "ssd": st.column_config.TextColumn("SSD"),
            "motivo_baixa": st.column_config.TextColumn("Motivo da Baixa"),
        },
        hide_index=True,
        use_container_width=True
    )



def render_faq_page():
    """Renderiza a página de FAQs, Tutoriais do SharePoint e Links Úteis da Bancada."""
    st.title("📚 FAQ, Tutoriais & Links Úteis da Bancada")
    st.write("Base de conhecimento centralizada com tutoriais da equipe e atalhos rápidos para sistemas externos.")
    st.markdown("---")

    db_path = root_dir / "chamados.db"
    json_faq_path = root_dir / "temp" / "faqs_template.json"
    json_links_path = root_dir / "temp" / "links_uteis_template.json"
    
    # Sincroniza/Cria a tabela faqs no SQLite a partir do JSON se necessário
    try:
        conn = sqlite3.connect(db_path)
        cursor = conn.cursor()
        cursor.execute("""
            CREATE TABLE IF NOT EXISTS faqs (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                titulo TEXT NOT NULL,
                tipo_faq TEXT NOT NULL,
                url TEXT NOT NULL UNIQUE,
                conteudo TEXT,
                data_atualizacao DATETIME DEFAULT CURRENT_TIMESTAMP
            )
        """)
        
        # Garante a existência da coluna conteudo
        cursor.execute("PRAGMA table_info(faqs)")
        cols_db = [col[1] for col in cursor.fetchall()]
        if "conteudo" not in cols_db:
            cursor.execute("ALTER TABLE faqs ADD COLUMN conteudo TEXT")
            conn.commit()

        # Importa registros do JSON de FAQs caso o banco esteja vazio
        cursor.execute("SELECT COUNT(*) FROM faqs")
        count = cursor.fetchone()[0]

        if count == 0 and json_faq_path.exists():
            import json
            with open(json_faq_path, "r", encoding="utf-8") as f:
                faqs_json = json.load(f)
            
            for item in faqs_json:
                cursor.execute("""
                    INSERT OR IGNORE INTO faqs (titulo, tipo_faq, url, conteudo)
                    VALUES (?, ?, ?, ?)
                """, (item.get("titulo"), item.get("tipo_faq", "Geral"), item.get("url"), item.get("conteudo")))
            conn.commit()

        df_faqs = pd.read_sql_query("SELECT id, titulo, tipo_faq, url, conteudo FROM faqs", conn)
        conn.close()
    except Exception as e:
        st.error(f"Erro ao carregar banco de dados de FAQs: {e}")
        df_faqs = pd.DataFrame()

    # Carrega Links Úteis
    links_uteis = []
    if json_links_path.exists():
        import json
        try:
            with open(json_links_path, "r", encoding="utf-8") as f:
                links_uteis = json.load(f)
        except Exception as e:
            logger.error(f"Erro ao ler links_uteis_template.json: {e}")

    # Criação das Abas Principais
    tab_faqs, tab_links = st.tabs(["📚 FAQs & Tutoriais (SharePoint)", "🔗 Links Úteis da Bancada"])

    with tab_faqs:
        if df_faqs.empty:
            st.info("Nenhum FAQ cadastrado no momento.")
        else:
            # Modal para Leitura do Conteúdo do FAQ
            @st.dialog("📖 Leitor de FAQ", width="large")
            def open_faq_modal(faq_id):
                faq_item = df_faqs[df_faqs['id'] == faq_id].iloc[0]
                
                st.markdown("""
                <style>
                .faq-container {
                    font-family: 'Segoe UI', system-ui, -apple-system, sans-serif;
                    color: #e0e0e0;
                    line-height: 1.6;
                }
                .faq-container h1, .faq-container h2, .faq-container h3, .faq-container h4 {
                    color: #ffffff !important;
                    margin-top: 1.5rem;
                    margin-bottom: 0.75rem;
                    font-weight: 600;
                }
                .faq-container p {
                    margin-bottom: 1rem;
                    font-size: 0.95rem;
                }
                .faq-container img {
                    max-width: 100% !important;
                    height: auto !important;
                    border-radius: 8px !important;
                    margin: 16px 0 !important;
                    box-shadow: 0 4px 12px rgba(0,0,0,0.4) !important;
                    border: 1px solid #343541 !important;
                }
                .faq-container ol, .faq-container ul {
                    padding-left: 1.5rem;
                    margin-bottom: 1rem;
                }
                .faq-container li {
                    margin-bottom: 0.4rem;
                }
                .faq-container code {
                    background-color: #2a2b36;
                    color: #ff4b4b;
                    padding: 2px 6px;
                    border-radius: 4px;
                    font-size: 0.9rem;
                }
                </style>
                """, unsafe_allow_html=True)

                st.subheader(faq_item['titulo'])
                st.caption(f"Categoria: **{faq_item['tipo_faq']}**")
                st.markdown("---")
                
                if faq_item['conteudo'] and str(faq_item['conteudo']).strip():
                    st.markdown(f'<div class="faq-container">{faq_item["conteudo"]}</div>', unsafe_allow_html=True)
                else:
                    st.info("O conteúdo detalhado deste FAQ ainda não foi sincronizado localmente.")
                    st.write("Você pode visualizar o tutorial completo diretamente no SharePoint pelo botão abaixo.")
                    
                st.markdown("---")
                st.markdown(f'<a href="{faq_item["url"]}" target="_blank" style="display: inline-block; background-color: #ff4b4b; color: white; text-decoration: none; font-weight: bold; padding: 8px 16px; border-radius: 6px;">🔗 Abrir no SharePoint (Nova Aba) ↗</a>', unsafe_allow_html=True)

            # Filtros e Busca na Sidebar Lateral Esquerda para FAQs
            st.sidebar.markdown("## 🔍 Filtros do FAQ")
            search_query = st.sidebar.text_input("Buscar por palavra-chave:", "", key="faq_search")
            
            tipos_disponiveis = ["Todos"] + sorted(df_faqs['tipo_faq'].dropna().unique().tolist())
            selected_tipo = st.sidebar.selectbox("📂 Categoria:", tipos_disponiveis, key="faq_cat")

            # Aplicação dos Filtros
            filtered_df = df_faqs.copy()
            if search_query:
                filtered_df = filtered_df[filtered_df['titulo'].str.contains(search_query, case=False, na=False)]
            if selected_tipo != "Todos":
                filtered_df = filtered_df[filtered_df['tipo_faq'] == selected_tipo]

            st.markdown(f"**Exibindo {len(filtered_df)} de {len(df_faqs)} FAQs / Tutoriais**")
            st.markdown("<br>", unsafe_allow_html=True)

            # Exibição dos FAQs em Cards num grid de 2 colunas
            cols = st.columns(2)
            for index, row in filtered_df.iterrows():
                col_target = cols[index % 2]
                with col_target:
                    with st.container(border=True):
                        st.caption(f"📌 {row['tipo_faq']}")
                        st.subheader(row['titulo'])
                        
                        c_btn1, c_btn2 = st.columns([1, 1])
                        with c_btn1:
                            if st.button("📖 Ler Tutorial", key=f"btn_read_{row['id']}", use_container_width=True):
                                open_faq_modal(row['id'])
                        with c_btn2:
                            st.markdown(f'<a href="{row["url"]}" target="_blank" style="display: block; text-align: center; background-color: #2a2b36; border: 1px solid #343541; color: white; text-decoration: none; font-size: 0.85rem; padding: 6px; border-radius: 6px; font-weight: bold;">🔗 SharePoint ↗</a>', unsafe_allow_html=True)

    with tab_links:
        st.subheader("🌐 Links e Atalhos Rápidos da Bancada")
        st.write("Acesso direto aos sistemas operacionais, filas de atendimento e ferramentas externas.")
        st.markdown("<br>", unsafe_allow_html=True)

        if not links_uteis:
            st.info("Nenhum link útil cadastrado em `temp/links_uteis_template.json`.")
        else:
            # Busca de Links Rápidos
            search_link = st.text_input("🔍 Pesquisar por Nome do Sistema / Link:", "", key="search_link_input")
            
            filtered_links = links_uteis
            if search_link:
                filtered_links = [l for l in links_uteis if search_link.lower() in l.get("titulo", "").lower()]

            # Exibe os links em um grid elegante de 3 colunas
            link_cols = st.columns(3)
            for idx, item in enumerate(filtered_links):
                col_lk = link_cols[idx % 3]
                with col_lk:
                    with st.container(border=True):
                        st.markdown(f"#### 🚀 {item.get('titulo', 'Sem Título')}")
                        st.caption(item.get("url", ""))
                        st.markdown("<br>", unsafe_allow_html=True)
                        st.markdown(
                            f'<a href="{item.get("url")}" target="_blank" style="display: block; text-align: center; background-color: #ff4b4b; color: white; text-decoration: none; font-weight: bold; padding: 10px; border-radius: 6px;">🔗 Acessar Sistema ↗</a>', 
                            unsafe_allow_html=True
                        )



def render_contracts_page():
    """Renderiza a página de Fiscalização de Contratos a partir da planilha oficial do OneDrive/SharePoint."""
    st.title("📜 Fiscalização de Contratos & Processos SAJ")
    st.write("Acompanhamento das indicações de fiscais titulares, suplentes, processos SAJ e portarias publicadas.")
    st.markdown("---")

    relative_path = os.getenv("FISCAL_EXCEL_RELATIVE_PATH", "")
    excel_file = Path.home() / relative_path if relative_path else None

    # Cabeçalho com botão de sincronização
    col1, col2 = st.columns([3, 1])
    with col1:
        st.info("Os dados exibidos abaixo são lidos em tempo real da planilha oficial sincronizada via OneDrive/SharePoint.")
    with col2:
        if st.button("🔄 Sincronizar Planilha", type="primary", use_container_width=True):
            st.cache_data.clear()
            st.toast("✅ Dados da planilha recarregados!", icon="🔄")
            st.rerun()

    if not excel_file or not excel_file.exists():
        st.warning(f"⚠️ Planilha de Fiscais não localizada no caminho:\n`{excel_file}`")
        st.info("Verifique se o OneDrive está sincronizado e o arquivo 'Indicação para atuar como fiscal.xlsx' está disponível.")
        return

    try:
        # Lê as três abas da planilha do Excel
        excel_data = pd.ExcelFile(excel_file)
        
        df_indicacoes = pd.read_excel(excel_data, sheet_name="Indicações") if "Indicações" in excel_data.sheet_names else pd.DataFrame()
        df_publicacoes = pd.read_excel(excel_data, sheet_name="Publicações") if "Publicações" in excel_data.sheet_names else pd.DataFrame()
        df_contador = pd.read_excel(excel_data, sheet_name="Contador") if "Contador" in excel_data.sheet_names else pd.DataFrame()
        
    except Exception as e:
        st.error(f"Erro ao ler a planilha de Fiscais: {e}")
        return

    # Normalização dos nomes das colunas de Indicações
    if not df_indicacoes.empty:
        df_indicacoes.columns = [str(col).strip() for col in df_indicacoes.columns]
    
    # Normalização das Publicações
    if not df_publicacoes.empty:
        df_publicacoes.columns = [str(col).strip() for col in df_publicacoes.columns]

    # --- CARDS KPI (CONTADOR DOS 3 FISCAIS PRINCIPAIS) ---
    fiscais_foco = [
        "Paulo Henrique Gonçalves Rezende",
        "Reginaldo da Silva Bandeira",
        "Luiz Leonardo Villalba"
    ]

    st.subheader("📊 Resumo de Contratos por Fiscal")
    kpi_cols = st.columns(3)

    for i, fiscal in enumerate(fiscais_foco):
        count_titular = 0
        count_suplente = 0
        
        if not df_indicacoes.empty:
            if "Fiscal titular" in df_indicacoes.columns:
                count_titular = (df_indicacoes["Fiscal titular"].astype(str).str.strip() == fiscal).sum()
            if "Fiscal suplente" in df_indicacoes.columns:
                count_suplente = (df_indicacoes["Fiscal suplente"].astype(str).str.strip() == fiscal).sum()
        
        total_fiscal = count_titular + count_suplente
        primeiro_nome = fiscal.split()[0] + " " + fiscal.split()[-1]

        with kpi_cols[i % 3]:
            st.markdown(f"""
            <div style="
                background-color: #1e1f25;
                border: 1px solid #343541;
                border-radius: 8px;
                padding: 16px;
                text-align: center;
                box-shadow: 0 4px 6px rgba(0,0,0,0.1);
            ">
                <h4 style="margin:0; color:#ff4b4b; font-size:1.1rem;">👤 {primeiro_nome}</h4>
                <h2 style="margin: 8px 0; color:#ffffff; font-size:2rem;">{total_fiscal} <span style="font-size:0.9rem; color:#a0a0a0;">processos</span></h2>
                <div style="display:flex; justify-content:space-around; margin-top:8px; font-size:0.8rem; color:#cccccc;">
                    <span>📌 Titular: <b>{count_titular}</b></span>
                    <span>🔄 Suplente: <b>{count_suplente}</b></span>
                </div>
            </div>
            """, unsafe_allow_html=True)

    st.markdown("<br>", unsafe_allow_html=True)

    # --- FILTROS SIDEBAR ---
    st.sidebar.markdown("## 🔍 Filtros de Contratos")
    
    # Filtro por Fiscal
    opcoes_fiscais = ["Todos"] + fiscais_foco
    selected_fiscal_filter = st.sidebar.selectbox("👤 Filtrar por Fiscal:", opcoes_fiscais)
    
    # Busca por texto livre (Objeto / Nº SAJ / Contrato)
    search_text = st.sidebar.text_input("🔍 Buscar por Nº SAJ, Objeto ou Contrato:", "")

    # Aplicação dos filtros em Indicações
    df_filtered_ind = df_indicacoes.copy()
    
    if selected_fiscal_filter != "Todos":
        cond_titular = df_filtered_ind["Fiscal titular"].astype(str).str.strip() == selected_fiscal_filter if "Fiscal titular" in df_filtered_ind.columns else False
        cond_suplente = df_filtered_ind["Fiscal suplente"].astype(str).str.strip() == selected_fiscal_filter if "Fiscal suplente" in df_filtered_ind.columns else False
        df_filtered_ind = df_filtered_ind[cond_titular | cond_suplente]

    if search_text:
        mask = pd.Series(False, index=df_filtered_ind.index)
        for col in df_filtered_ind.columns:
            mask = mask | df_filtered_ind[col].astype(str).str.contains(search_text, case=False, na=False)
        df_filtered_ind = df_filtered_ind[mask]

    # --- ABAS DE EXIBIÇÃO ---
    tab_ind, tab_charts, tab_pub, tab_raw_count = st.tabs(["📋 Indicações de Fiscais", "📈 Gráficos & Estatísticas", "📰 Publicações & Portarias", "📊 Tabela Contadora"])

    with tab_ind:
        c_head1, c_head2 = st.columns([3, 1])
        with c_head1:
            st.subheader(f"📋 Processos e Indicações ({len(df_filtered_ind)} registros)")
        with c_head2:
            if not df_filtered_ind.empty:
                # Botão de download dos dados filtrados para Excel
                import io
                output_buffer = io.BytesIO()
                with pd.ExcelWriter(output_buffer, engine='openpyxl') as writer:
                    df_filtered_ind.to_excel(writer, index=False, sheet_name='Fiscais')
                excel_bytes = output_buffer.getvalue()
                
                st.download_button(
                    label="📥 Exportar Excel",
                    data=excel_bytes,
                    file_name="contratos_fiscais_filtrados.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    use_container_width=True
                )

        if not df_filtered_ind.empty:
            st.dataframe(
                df_filtered_ind,
                use_container_width=True,
                hide_index=True,
                column_config={
                    "nº Saj": st.column_config.TextColumn("Nº SAJ"),
                    "Fiscal titular": st.column_config.TextColumn("Fiscal Titular"),
                    "Fiscal suplente": st.column_config.TextColumn("Fiscal Suplente"),
                    "Objeto": st.column_config.TextColumn("Objeto / Descrição"),
                    "Contrato": st.column_config.TextColumn("Contrato / Empenho"),
                }
            )
        else:
            st.info("Nenhum registro encontrado para os filtros selecionados.")

    with tab_charts:
        st.subheader("📈 Visão Geral da Carga de Trabalho dos Fiscais")
        if not df_indicacoes.empty and "Fiscal titular" in df_indicacoes.columns and "Fiscal suplente" in df_indicacoes.columns:
            # Prepara dataframe comparativo de Titulares vs Suplentes
            df_t = df_indicacoes["Fiscal titular"].dropna().astype(str).str.strip().value_counts().reset_index()
            df_t.columns = ["Fiscal", "Como Titular"]
            
            df_s = df_indicacoes["Fiscal suplente"].dropna().astype(str).str.strip().value_counts().reset_index()
            df_s.columns = ["Fiscal", "Como Suplente"]
            
            df_comp = pd.merge(df_t, df_s, on="Fiscal", how="outer").fillna(0)
            df_comp["Como Titular"] = df_comp["Como Titular"].astype(int)
            df_comp["Como Suplente"] = df_comp["Como Suplente"].astype(int)
            df_comp["Total Processos"] = df_comp["Como Titular"] + df_comp["Como Suplente"]
            df_comp = df_comp.sort_values(by="Total Processos", ascending=False)
            
            g_col1, g_col2 = st.columns(2)
            with g_col1:
                st.markdown("#### 📌 Distribuição de Titularidades")
                st.bar_chart(df_comp.set_index("Fiscal")[["Como Titular"]], use_container_width=True)
            with g_col2:
                st.markdown("#### 🔄 Distribuição de Suplências")
                st.bar_chart(df_comp.set_index("Fiscal")[["Como Suplente"]], use_container_width=True)
                
            st.markdown("---")
            st.markdown("#### 📊 Carga Total Comparativa de Fiscais")
            st.bar_chart(df_comp.set_index("Fiscal")[["Como Titular", "Como Suplente"]], use_container_width=True)

            # --- NOVO GRÁFICO POR TIPO DE OBJETO ---
            st.markdown("---")
            st.subheader("📦 Agrupamento por Tipo de Objeto / Equipamento")
            
            if "Objeto" in df_indicacoes.columns:
                df_obj = df_indicacoes["Objeto"].dropna().astype(str).str.strip().str.lower()
                
                # Mapeamento / Categorização dos objetos
                def categorizar_objeto(desc):
                    if "monitor" in desc:
                        return "🖥️ Monitores"
                    elif "desktop" in desc or "computador" in desc:
                        return "💻 Desktops / Computadores"
                    elif "fone" in desc or "headset" in desc:
                        return "🎧 Fones / Headsets"
                    elif "webcam" in desc or "mouse" in desc:
                        return "🖱️ Periféricos (Webcam/Mouse/Teclado)"
                    elif "notebook" in desc or "laptop" in desc:
                        return "💻 Notebooks"
                    elif "telefone" in desc or "ramal" in desc:
                        return "📞 Telefonia / Ramais"
                    elif "hd" in desc or "ssd" in desc:
                        return "💾 Armazenamento (HD/SSD)"
                    elif "internet" in desc or "satélite" in desc:
                        return "📡 Internet / Conectividade"
                    elif "scanner" in desc:
                        return "🖨️ Scanners / Impressão"
                    elif "tablet" in desc:
                        return "📱 Tablets"
                    else:
                        return "📦 Outros Suprimentos / Serviços"

                df_indicacoes_cats = df_indicacoes.copy()
                df_indicacoes_cats["Categoria_Objeto"] = df_indicacoes_cats["Objeto"].astype(str).apply(categorizar_objeto)
                
                counts_obj = df_indicacoes_cats["Categoria_Objeto"].value_counts().reset_index()
                counts_obj.columns = ["Categoria", "Quantidade"]
                
                o_col1, o_col2 = st.columns([2, 1])
                with o_col1:
                    st.bar_chart(counts_obj.set_index("Categoria"), use_container_width=True)
                with o_col2:
                    st.markdown("##### 📌 Quantidade por Tipo:")
                    for _, r in counts_obj.iterrows():
                        st.markdown(f"- **{r['Categoria']}**: `{r['Quantidade']}` processos")
        else:
            st.info("Dados insuficientes para renderização dos gráficos.")

    with tab_pub:
        st.subheader("📰 Publicações em Diário Oficial & Portarias")
        df_filtered_pub = df_publicacoes.copy()
        
        if selected_fiscal_filter != "Todos" and not df_filtered_pub.empty:
            cond_t = df_filtered_pub["Fiscal titular"].astype(str).str.strip() == selected_fiscal_filter if "Fiscal titular" in df_filtered_pub.columns else False
            cond_s = df_filtered_pub["Fiscal suplente"].astype(str).str.strip() == selected_fiscal_filter if "Fiscal suplente" in df_filtered_pub.columns else False
            df_filtered_pub = df_filtered_pub[cond_t | cond_s]

        if search_text and not df_filtered_pub.empty:
            mask_p = pd.Series(False, index=df_filtered_pub.index)
            for col in df_filtered_pub.columns:
                mask_p = mask_p | df_filtered_pub[col].astype(str).str.contains(search_text, case=False, na=False)
            df_filtered_pub = df_filtered_pub[mask_p]

        if not df_filtered_pub.empty:
            st.dataframe(
                df_filtered_pub,
                use_container_width=True,
                hide_index=True,
                column_config={
                    "nº Saj": st.column_config.TextColumn("Nº SAJ"),
                    "Fiscal titular": st.column_config.TextColumn("Fiscal Titular"),
                    "Fiscal suplente": st.column_config.TextColumn("Fiscal Suplente"),
                    "Objeto": st.column_config.TextColumn("Objeto / Nota de Empenho"),
                    "Obs.:": st.column_config.TextColumn("Portaria / Observações"),
                }
            )
        else:
            st.info("Nenhuma publicação/portaria encontrada.")

    with tab_raw_count:
        st.subheader("📊 Tabela de Contagem Geral")
        if not df_contador.empty:
            st.dataframe(df_contador, use_container_width=True, hide_index=True)
        else:
            st.info("Aba Contador indisponível ou vazia na planilha.")


# Estado da página ativa
if "current_page" not in st.session_state:
    st.session_state["current_page"] = "📋 Painel de Chamados"

# Menu Hambúrguer (Popover) fixado no Header Superior Direito via CSS
with st.popover("☰ Menu"):
    st.markdown("### 📌 Sistemas / Páginas")
    if st.button("📋 Painel de Chamados", use_container_width=True):
        st.session_state["current_page"] = "📋 Painel de Chamados"
        st.rerun()
    if st.button("📍 Mapa & Localização", use_container_width=True):
        st.session_state["current_page"] = "📍 Mapa & Localização"
        st.rerun()
    if st.button("🖥️ Doação & Redistribuição", use_container_width=True):
        st.session_state["current_page"] = "🖥️ Doação & Redistribuição"
        st.rerun()
    if st.button("📜 Fiscalização de Contratos", use_container_width=True):
        st.session_state["current_page"] = "📜 Fiscalização de Contratos"
        st.rerun()
    if st.button("📚 FAQ & Tutoriais", use_container_width=True):
        st.session_state["current_page"] = "📚 FAQ & Tutoriais"
        st.rerun()

page = st.session_state["current_page"]

if page == "📍 Mapa & Localização":
    render_mapa_page()
    st.stop()

if page == "🖥️ Doação & Redistribuição":
    render_donations_page()
    st.stop()

if page == "📜 Fiscalização de Contratos":
    render_contracts_page()
    st.stop()

if page == "📚 FAQ & Tutoriais":
    render_faq_page()
    st.stop()



# Cabeçalho com Título e Botão de Sincronização
col_title, col_btn = st.columns([3, 1])
with col_title:
    st.title("📊 Painel de Chamados Centralizado")
    st.write("Visualize e interaja com os chamados do OTRS e CitSmart.")

# Verifica estado de execução global do robô
robo_ativo = check_orquestrador_running()

with col_btn:
    st.markdown("<div style='height: 15px;'></div>", unsafe_allow_html=True)
    if robo_ativo:
        # Se o robô estiver rodando, desativa o botão e exibe um sinal visual ativo
        st.button("🤖 Robô em Execução...", use_container_width=True, disabled=True)
    else:
        run_orquestrador = st.button(
            "🔄 Atualizar Chamados", 
            use_container_width=True, 
            help="Executa o orquestrador completo em segundo plano.",
            type="primary"
        )
        if run_orquestrador:
            import subprocess
            import sys
            # Dispara o orquestrador em segundo plano sem bloquear a aplicação Streamlit
            subprocess.Popen([sys.executable, "orquestrador.py"])
            st.toast("🚀 Robô iniciado em segundo plano!", icon="🤖")
            st.cache_data.clear() # Limpa caches para receber novos dados no término
            st.rerun()

# Se o robô estiver ativo, exibe uma seção bonita mostrando o progresso em tempo real (sem travar)
if robo_ativo:
    with st.expander("🤖 Robô Rodando em Segundo Plano – Acompanhar Progresso", expanded=False):
        st.info("O robô está coletando novos chamados e classificando com IA neste momento. Você pode continuar usando o painel normalmente!")
        
        # Lê e exibe os logs dinamicamente
        logs = read_last_log_lines(15)
        st.code(logs, language="text")
        
        # Botão rápido para atualizar o status dos logs manualmente
        st.button("🔄 Atualizar Progresso", help="Recarrega as últimas linhas de log do robô")


DB_PATH = Path("chamados.db")

# Cores oficiais das TAGs para uso unificado na tabela e no modal
TAG_COLORS = {
    "BACKUP": "#dd5358",
    "EVENTO": "#ce66ce",
    "FORMATAÇÃO": "#d38a62",
    "GARANTIA": "#518bbb",
    "IMPRESSORA": "#C6EFCE",
    "INSTALAÇÃO HARDWARE": "#FCE4D6",
    "INSTALAÇÃO SOFTWARE": "#86BEEE",
    "MANUTENÇÃO": "#E9CF69",
    "MONITOR": "#cbdd6f",
    "MUDANÇA": "#21ffe0",
    "PREPARAÇÃO COMPUTADORES": "#f09c72",
    "REDE": "#B7F391",
    "SOLICITAÇÃO SSD": "#f5a89b",
    "SUPORTE": "#FFE699",
    "TELEFONIA FIXA": "#e273a1",
    "VIAGEM": "#61e7c6",
    "VISTORIA CPDS": "#b2740e",
}

@st.cache_resource
def load_spacy_model():
    """Carrega o modelo spaCy local em português, com fallback caso falhe."""
    import spacy
    try:
        return spacy.load("pt_core_news_sm")
    except Exception:
        return None

@st.cache_data(show_spinner=False)
def summarize_ticket_locally(description: str, comments: str, max_sentences: int = 2) -> str:
    """
    Resume o chamado técnico localmente usando Processamento de Linguagem Natural (spaCy).
    Remove saudações, cortesia, jargões de encaminhamento e administrative boilerplate antes de pontuar.
    Usa um sistema de frequência com boost para termos técnicos e penalidades por tamanho de sentença.
    """
    import re
    from collections import Counter
    from src.config import clean_otrs_description
    
    # Pré-processa e limpa metadados e formulários estruturados (especialmente OTRS)
    description = clean_otrs_description(description)
    
    # 1. Função para limpar saudações e formalidades administrativas
    def clean_text(t: str) -> str:
        if not t:
            return ""
        # Remove saudações no início do texto ou de sentenças (ex: "prezados, solicito...")
        t = re.sub(
            r'^\s*(?:prezados?|prezadas?|caros?|caras?|olá|ola|bom\s+dia|boa\s+tarde|boa\s+noite|prezada\s+equipe|prezada\s+sti)\b(?:[^\n\.\?]*[\n\.,\?])?',
            '', t, flags=re.IGNORECASE
        )
        t = re.sub(
            r'^\s*(?:tudo\s+bem\??|espero\s+que\s+esteja\s+tudo\s+bem\??|espero\s+que\s+sim\??)',
            '', t, flags=re.IGNORECASE
        )
        
        # Remove verbos formais de pedido/encaminhamento
        t = re.sub(
            r'\b(?:gostaria\s+de\s+|venho\s+(?:por\s+meio\s+deste\s+)?|favor\s+|por\s+gentileza\s+|gentileza\s+)\b',
            '', t, flags=re.IGNORECASE
        )
        t = re.sub(
            r'\b(?:solicito\s+providências\s+para\s+|solicito\s+(?:a|o|que|os|as)?\s+|encaminho\s+para\s+providências\s*(?:[ao]s?|para|de)?\s+|encaminho\s+para\s+|segue\s+para\s+|segue\s+o\s+chamado\s+(?:para\s+)?)\b',
            '', t, flags=re.IGNORECASE
        )
        
        # Remove referências a anexos
        t = re.sub(
            r'\b(?:conforme|como|conforme\s+mostra\s+a|ver|veja)\s+(?:imagem\s+)?(?:em\s+)?anexo\b',
            '', t, flags=re.IGNORECASE
        )
        t = re.sub(
            r'\b(?:segue[m]?\s+)?(?:em\s+)?anexo\b',
            '', t, flags=re.IGNORECASE
        )
        
        # Remove jargões de encerramento
        t = re.sub(
            r'\b(?:fico|ficamos)\s+(?:à|a)\s+disposição\s+para\s+(?:eventuais|quaisquer)\s+(?:esclarecimentos|dúvidas)\b\.?',
            '', t, flags=re.IGNORECASE
        )
        t = re.sub(
            r'\b(?:desde\s+já\s+)?agradeço[s]?\b\.?',
            '', t, flags=re.IGNORECASE
        )
        t = re.sub(
            r'\b(?:atenciosamente|grato|obrigado|fico\s+no\s+aguardo|aguardo\s+retorno|sem\s+mais)\b\.?',
            '', t, flags=re.IGNORECASE
        )
        
        # Limpezas adicionais de espaçamento e pontuação órfã
        t = re.sub(r'\s+', ' ', t)
        t = re.sub(r'^\s*[,\.\-\:\/]+\s*', '', t)
        return t.strip()

    # Limpa a descrição principal
    desc_clean = clean_text(str(description))
    
    # Limpa e processa os comentários históricos
    comments_clean_list = []
    if comments:
        for line in str(comments).split('\n'):
            # Remove marcadores de metadados do comentário (autor, data): e.g. "- 13/05/2026 Celso: texto"
            line_payload = re.sub(r'^[-\s]*[\d/:\s\[\]\-\#\.]+(?:[\w\s\(\)]+)?:\s*', '', line).strip()
            line_clean = clean_text(line_payload)
            if line_clean and len(line_clean) > 8:
                comments_clean_list.append(line_clean)
                
    # Consolida os textos limpos para análise
    text_parts = []
    if desc_clean:
        text_parts.append(desc_clean)
    if comments_clean_list:
        text_parts.append(" ".join(comments_clean_list))
        
    combined_text = " ".join(text_parts).strip()
    
    if not combined_text:
        return "Sem descrição detalhada."
        
    nlp = load_spacy_model()
    
    # Fallback simples caso o modelo spaCy falhe em carregar
    if nlp is None:
        sentences = [s.strip() for s in combined_text.split('.') if len(s.strip()) > 8]
        if sentences:
            res = ". ".join(sentences[:max_sentences])
            if not res.endswith('.'):
                res += '.'
            return res
        return combined_text[:140] + "..." if len(combined_text) > 140 else combined_text

    doc = nlp(combined_text)
    
    # Termos de suporte técnico STI comuns para receberem "boost" de relevância
    TECHNICAL_BOOST = {
        "ssd", "hd", "windows", "formatação", "formatar", "lentidão", "travamento", "travando",
        "impressora", "imprimir", "rede", "conexão", "erro", "falha", "sistema", "configurar",
        "configuração", "instalação", "instalar", "senha", "usuário", "computador", "máquina",
        "notebook", "monitor", "teclado", "mouse", "backup", "servidor", "internet", "cabo",
        "wi-fi", "wifi", "login", "acesso", "workstation", "driver", "inicialização", "boot",
        "perfil", "outlook", "email", "e-mail", "toner", "cartucho", "suporte", "atualizar",
        "atualização", "office", "word", "excel", "pasta", "rede", "compartilhamento"
    }
    
    # 1. Filtra stopwords/pontuação e calcula a frequência das palavras-chave relevantes
    keywords = []
    for token in doc:
        if token.is_stop or token.is_punct or token.is_space:
            continue
        if token.pos_ in ["NOUN", "VERB", "ADJ", "PROPN"]:
            keywords.append(token.text.lower())
            
    if not keywords:
        sentences = list(doc.sents)
        return " ".join([s.text.strip() for s in sentences[:max_sentences]])
        
    # Calcula frequência normalizada
    word_freq = Counter(keywords)
    max_freq = max(word_freq.values())
    for word in word_freq:
        word_freq[word] = word_freq[word] / max_freq
        
    # 2. Pontua as sentenças reais com base nas palavras-chave, boost técnico e tamanho
    sent_scores = {}
    sentences = list(doc.sents)
    
    for idx, sent in enumerate(sentences):
        words = [t for t in sent if not t.is_punct and not t.is_space]
        if len(words) < 3:
            continue
            
        score = 0
        for token in sent:
            word_lower = token.text.lower()
            if word_lower in word_freq:
                score += word_freq[word_lower]
            # Boost extra para termos de tecnologia cruciais na triagem
            if word_lower in TECHNICAL_BOOST:
                score += 3.0
                
        # Penaliza sentenças longas demais e privilegia tamanhos fáceis de ler no WhatsApp
        word_count = len(words)
        if 8 <= word_count <= 25:
            score *= 1.3
        elif word_count > 30:
            score *= 0.6
        elif word_count < 6:
            score *= 0.7
            
        # A primeira frase geralmente traz o assunto principal
        if idx == 0:
            score += 2.0
            
        # O último comentário costuma trazer as ações mais recentes realizadas
        if idx == len(sentences) - 1 and len(sentences) > 1:
            score += 1.0
            
        sent_scores[sent] = score
        
    if not sent_scores:
        res = " ".join([s.text.strip() for s in sentences[:max_sentences]])
        return res
        
    # 3. Seleciona as frases de maior pontuação e as ordena conforme a aparição no chamado
    sorted_sents = sorted(sent_scores.keys(), key=lambda x: sent_scores[x], reverse=True)
    top_sents = sorted_sents[:max_sentences]
    top_sents = sorted(top_sents, key=lambda x: x.start)
    
    # 4. Formata o retorno garantindo capitalização correta
    formatted_sentences = []
    for s in top_sents:
        sent_text = s.text.strip()
        if not sent_text:
            continue
        # Garante letra maiúscula no início de cada sentença
        sent_text = sent_text[0].upper() + sent_text[1:]
        # Remove eventuais resíduos de pontuação no início
        sent_text = re.sub(r'^[\s,\.\-\:\/]+', '', sent_text)
        if not sent_text.endswith(('.', '!', '?')):
            sent_text += '.'
        formatted_sentences.append(sent_text)
        
    summary = " ".join(formatted_sentences)
    return summary

def load_data():
    if not DB_PATH.exists():
        return pd.DataFrame()
    conn = sqlite3.connect(DB_PATH)
    df = pd.read_sql_query("SELECT * FROM chamados", conn)
    conn.close()
    
    # Limpa " - Sede" de forma inteligente na coluna de exibição da Localidade Física
    if 'localidade_fisica' in df.columns:
        import re
        df['localidade_fisica'] = df['localidade_fisica'].apply(
            lambda x: re.sub(r'\s*-\s*Sede\b', '', str(x), flags=re.IGNORECASE).strip() if pd.notna(x) else x
        )
    return df

df = load_data()

if df.empty:
    st.warning("Nenhum dado encontrado no banco de dados. Execute o orquestrador primeiro!")
else:
    # Tratamento de datas para exibição e filtro
    df['datetime_obj'] = pd.to_datetime(df['data_criacao'], errors='coerce')
    df['Data Formatada'] = df['datetime_obj'].dt.strftime('%d/%m/%Y %H:%M:%S')
    df['Data Formatada'] = df['Data Formatada'].fillna(df['data_criacao'])

    # Barra lateral de filtros com botão de limpar global
    min_date = df['datetime_obj'].dropna().min().date() if not df['datetime_obj'].dropna().empty else datetime.now().date()
    max_date = df['datetime_obj'].dropna().max().date() if not df['datetime_obj'].dropna().empty else datetime.now().date()
    
    # Inicializa os estados no st.session_state para os filtros se não existirem
    if "f_date_range" not in st.session_state:
        st.session_state["f_date_range"] = (min_date, max_date)
    if "f_status" not in st.session_state:
        st.session_state["f_status"] = []
    if "f_tags" not in st.session_state:
        st.session_state["f_tags"] = []
    if "custom_loc_selection" not in st.session_state:
        st.session_state["custom_loc_selection"] = []
    if "f_cities" not in st.session_state:
        st.session_state["f_cities"] = []
    if "f_units" not in st.session_state:
        st.session_state["f_units"] = []
    if "f_bases" not in st.session_state:
        st.session_state["f_bases"] = []
    if "f_user" not in st.session_state:
        st.session_state["f_user"] = ""
    if "f_mode" not in st.session_state:
        st.session_state["f_mode"] = "🟢 Manter Selecionados"
    if "f_ticket_ids" not in st.session_state:
        st.session_state["f_ticket_ids"] = []

    def get_filtered_options(col_name: str) -> list:
        """
        Calcula as opções únicas disponíveis para uma coluna específica,
        aplicando todos os filtros ativos, exceto o filtro da própria coluna,
        respeitando o modo de filtragem (Manter vs Ocultar).
        """
        temp_df = df.copy()
        is_exclude_mode = (st.session_state.get("f_mode") == "🔴 Ocultar Selecionados")
        
        # 1. Filtro de data
        dr = st.session_state.get("f_date_range", (min_date, max_date))
        if isinstance(dr, tuple) and len(dr) == 2:
            start_date, end_date = dr
            temp_df = temp_df[
                (temp_df['datetime_obj'].dt.date >= start_date) & 
                (temp_df['datetime_obj'].dt.date <= end_date)
            ]
            
        # 2. Aplica demais filtros (exceto o próprio) respeitando o modo de exclusão
        if col_name != 'status' and st.session_state.get("f_status"):
            if is_exclude_mode:
                temp_df = temp_df[~temp_df['status'].isin(st.session_state["f_status"])]
            else:
                temp_df = temp_df[temp_df['status'].isin(st.session_state["f_status"])]
            
        if col_name != 'tag' and st.session_state.get("f_tags"):
            if is_exclude_mode:
                temp_df = temp_df[~temp_df['tag'].isin(st.session_state["f_tags"])]
            else:
                temp_df = temp_df[temp_df['tag'].isin(st.session_state["f_tags"])]
            
        if col_name != 'localidade_fisica' and st.session_state.get("custom_loc_selection"):
            if is_exclude_mode:
                temp_df = temp_df[~temp_df['localidade_fisica'].isin(st.session_state["custom_loc_selection"])]
            else:
                temp_df = temp_df[temp_df['localidade_fisica'].isin(st.session_state["custom_loc_selection"])]
            
        if col_name != 'cidade_predio' and st.session_state.get("f_cities"):
            if is_exclude_mode:
                temp_df = temp_df[~temp_df['cidade_predio'].isin(st.session_state["f_cities"])]
            else:
                temp_df = temp_df[temp_df['cidade_predio'].isin(st.session_state["f_cities"])]
            
        if col_name != 'unidade' and st.session_state.get("f_units"):
            if is_exclude_mode:
                temp_df = temp_df[~temp_df['unidade'].isin(st.session_state["f_units"])]
            else:
                temp_df = temp_df[temp_df['unidade'].isin(st.session_state["f_units"])]
            
        if col_name != 'base' and st.session_state.get("f_bases"):
            if is_exclude_mode:
                temp_df = temp_df[~temp_df['base'].isin(st.session_state["f_bases"])]
            else:
                temp_df = temp_df[temp_df['base'].isin(st.session_state["f_bases"])]
            
        if col_name != 'usuario' and st.session_state.get("f_user"):
            if is_exclude_mode:
                temp_df = temp_df[~temp_df['usuario'].str.contains(st.session_state["f_user"], case=False, na=False)]
            else:
                temp_df = temp_df[temp_df['usuario'].str.contains(st.session_state["f_user"], case=False, na=False)]
            
        options = sorted(list(temp_df[col_name].dropna().unique()))
        
        # Garante que qualquer opção selecionada no widget atual permaneça na lista de opções para evitar erros do Streamlit
        key_map = {
            'status': 'f_status',
            'tag': 'f_tags',
            'localidade_fisica': 'custom_loc_selection',
            'cidade_predio': 'f_cities',
            'unidade': 'f_units',
            'base': 'f_bases',
            'usuario': 'f_user'
        }
        session_key = key_map.get(col_name)
        if session_key:
            current_selection = st.session_state.get(session_key, [])
            if current_selection:
                if isinstance(current_selection, list):
                    for val in current_selection:
                        if val not in options:
                            options.append(val)
                elif current_selection not in options:
                    options.append(current_selection)
                    
        return options

    # Cabeçalho dos Filtros na Sidebar
    st.sidebar.markdown("### 🔍 Filtros de Chamados")
    
    # Seletor de Modo do Filtro (Inclusão vs Exclusão)
    filter_mode = st.sidebar.radio(
        "Modo de Filtragem:",
        options=["🟢 Manter Selecionados", "🔴 Ocultar Selecionados"],
        key="f_mode",
        horizontal=True,
        help="🟢 Manter: Mostra apenas os itens escolhidos.\n🔴 Ocultar: Exibe tudo EXCETO os itens escolhidos (ideal para 'Todos Menos Um')."
    )
    
    # Opções carregadas antecipadamente para o botão Marcar Todos
    status_options = get_filtered_options('status')
    tag_options = get_filtered_options('tag')
    city_options = get_filtered_options('cidade_predio')
    unit_options = get_filtered_options('unidade')
    base_options = get_filtered_options('base')

    # Botões rápidos Marcar Todos / Limpar Tudo
    col_btn_sel1, col_btn_sel2 = st.sidebar.columns(2)
    with col_btn_sel1:
        if st.button("☑️ Marcar Todos", use_container_width=True, help="Seleciona todas as opções para que você possa remover apenas uma ou duas"):
            st.session_state["f_status"] = list(status_options)
            st.session_state["f_tags"] = list(tag_options)
            st.session_state["f_cities"] = list(city_options)
            st.session_state["f_units"] = list(unit_options)
            st.session_state["f_bases"] = list(base_options)
            st.rerun()
    with col_btn_sel2:
        if st.button("🧹 Limpar Tudo", use_container_width=True, help="Limpa todos os filtros ativos de uma vez"):
            st.session_state["f_date_range"] = (min_date, max_date)
            st.session_state["f_status"] = []
            st.session_state["f_tags"] = []
            st.session_state["custom_loc_selection"] = []
            st.session_state["f_cities"] = []
            st.session_state["f_units"] = []
            st.session_state["f_bases"] = []
            st.session_state["f_user"] = ""
            st.session_state["f_ticket_ids"] = []
            st.rerun()

    st.sidebar.markdown("---")

    # Filtro de Datas (Calendário)
    date_range = st.sidebar.date_input(
        "Intervalo de Datas",
        value=st.session_state["f_date_range"],
        min_value=min_date,
        max_value=max_date,
        format="DD/MM/YYYY",
        key="f_date_range"
    )
    
    # Filtros Multi-seleção com opções dinâmicas dependentes
    selected_status = st.sidebar.multiselect(
        "Status", 
        options=status_options, 
        key="f_status", 
        placeholder="Escolha as opções..."
    )
    
    selected_tags = st.sidebar.multiselect(
        "TAG (Categoria de IA)", 
        options=tag_options, 
        key="f_tags", 
        placeholder="Escolha as opções..."
    )
    
    # Componente Customizado Select Multiple nativo para Localidades
    loc_options = get_filtered_options('localidade_fisica')
    
    st.sidebar.markdown("---")
    st.sidebar.subheader("📍 Seleção de Localidades")
    st.sidebar.write("Arraste o mouse sobre as opções ou segure **Shift** para selecionar múltiplas de uma vez:")
    
    # Executa o componente customizado passando as opções dinâmicas
    with st.sidebar:
        selected_locs = custom_select(
            options=loc_options,
            default=st.session_state.get("custom_loc_selection", []),
            key="custom_loc_selection"
        )
    if selected_locs is None:
        selected_locs = []
    
    selected_cities = st.sidebar.multiselect(
        "Cidade - Prédio", 
        options=city_options, 
        key="f_cities", 
        placeholder="Escolha as opções..."
    )
    
    selected_units = st.sidebar.multiselect(
        "Unidade", 
        options=unit_options, 
        key="f_units", 
        placeholder="Escolha as opções..."
    )
    
    # Filtro de Base (CitSmart/OTRS)
    selected_bases = st.sidebar.multiselect(
        "Base de Origem", 
        options=base_options, 
        key="f_bases", 
        placeholder="Escolha as opções..."
    )
    
    # Filtro de Usuário (Busca por texto)
    user_search = st.sidebar.text_input(
        "Buscar por Usuário", 
        key="f_user", 
        placeholder="Digite o nome do usuário..."
    )

    # Aplica os filtros principais primeiro
    filtered_df = df.copy()
    is_exclude_mode = (filter_mode == "🔴 Ocultar Selecionados")
    
    # Filtro de data
    if isinstance(date_range, tuple) and len(date_range) == 2:
        start_date, end_date = date_range
        filtered_df = filtered_df[
            (filtered_df['datetime_obj'].dt.date >= start_date) & 
            (filtered_df['datetime_obj'].dt.date <= end_date)
        ]
        
    if selected_status:
        if is_exclude_mode:
            filtered_df = filtered_df[~filtered_df['status'].isin(selected_status)]
        else:
            filtered_df = filtered_df[filtered_df['status'].isin(selected_status)]

    if selected_tags:
        if is_exclude_mode:
            filtered_df = filtered_df[~filtered_df['tag'].isin(selected_tags)]
        else:
            filtered_df = filtered_df[filtered_df['tag'].isin(selected_tags)]

    if selected_locs:
        if is_exclude_mode:
            filtered_df = filtered_df[~filtered_df['localidade_fisica'].isin(selected_locs)]
        else:
            filtered_df = filtered_df[filtered_df['localidade_fisica'].isin(selected_locs)]

    if selected_cities:
        if is_exclude_mode:
            filtered_df = filtered_df[~filtered_df['cidade_predio'].isin(selected_cities)]
        else:
            filtered_df = filtered_df[filtered_df['cidade_predio'].isin(selected_cities)]

    if selected_units:
        if is_exclude_mode:
            filtered_df = filtered_df[~filtered_df['unidade'].isin(selected_units)]
        else:
            filtered_df = filtered_df[filtered_df['unidade'].isin(selected_units)]

    if selected_bases:
        if is_exclude_mode:
            filtered_df = filtered_df[~filtered_df['base'].isin(selected_bases)]
        else:
            filtered_df = filtered_df[filtered_df['base'].isin(selected_bases)]

    if user_search:
        if is_exclude_mode:
            filtered_df = filtered_df[~filtered_df['usuario'].str.contains(user_search, case=False, na=False)]
        else:
            filtered_df = filtered_df[filtered_df['usuario'].str.contains(user_search, case=False, na=False)]

    st.sidebar.markdown("---")
    
    # Filtro Específico por Chamado Individual (alimentado pelos chamados filtrados ou pela base completa se vazio)
    source_df = filtered_df if not filtered_df.empty else df
    df_tickets_list = source_df[['id', 'titulo']].copy()
    ids_clean = df_tickets_list['id'].astype(str).str.replace(r'\.0$', '', regex=True).str.strip()
    titles_clean = df_tickets_list['titulo'].fillna("Sem Título").astype(str).str.strip()
    labels = ids_clean + " - " + titles_clean
    ticket_options = sorted(labels.unique().tolist())
    
    # Garante que qualquer chamado já selecionado continue na lista de opções para evitar erros do Streamlit
    current_sel_tickets = st.session_state.get("f_ticket_ids", [])
    for sel_t in current_sel_tickets:
        if sel_t not in ticket_options:
            ticket_options.append(sel_t)

    selected_tickets = st.sidebar.multiselect(
        "🎫 Chamados Específicos (por ID / Título)",
        options=ticket_options,
        key="f_ticket_ids",
        placeholder="Selecione chamados individuais..."
    )
    
    st.sidebar.markdown("---")

    # Aplica o filtro de chamados específicos selecionados
    if selected_tickets:
        selected_ids = [t.split(" - ")[0].strip() for t in selected_tickets]
        clean_df_ids = filtered_df['id'].astype(str).str.replace(r'\.0$', '', regex=True).str.strip()
        if is_exclude_mode:
            filtered_df = filtered_df[~clean_df_ids.isin(selected_ids)]
        else:
            filtered_df = filtered_df[clean_df_ids.isin(selected_ids)]
        
    # Exibe métricas no topo
    col1, col2, col3 = st.columns(3)
    col1.metric("Total de Chamados", len(filtered_df))
    col2.metric("Abertos", len(filtered_df[filtered_df['status'] == 'Aberto']))
    col3.metric("Fechados", len(filtered_df[filtered_df['status'] == 'Fechado']))
    
    st.write("---")

    # Organização em Abas: Tabela principal vs Gráficos & Estatísticas
    main_tab_list, main_tab_charts = st.tabs(["📋 Tabela Geral de Chamados", "📈 Gráficos & Estatísticas do Painel"])

    with main_tab_charts:
        st.subheader("📊 Análise Estatística dos Chamados Filtrados")
        st.write("Visualização consolidada de abertura de chamados por Prédio, Unidade, Categorias (TAGs) e Usuários.")
        st.markdown("<br>", unsafe_allow_html=True)

        if not filtered_df.empty:
            # Row 1: Prédios e Unidades mais demandantes
            g_col1, g_col2 = st.columns(2)
            with g_col1:
                st.markdown("#### 🏢 Top Prédios / Cidades com Mais Chamados")
                city_counts = filtered_df['cidade_predio'].value_counts().head(10).reset_index()
                city_counts.columns = ['Prédio / Cidade', 'Quantidade']
                st.bar_chart(city_counts.set_index('Prédio / Cidade'), use_container_width=True)

            with g_col2:
                st.markdown("#### 🏛️ Top Unidades / Setores Mais Demandantes")
                unit_counts = filtered_df['unidade'].value_counts().head(10).reset_index()
                unit_counts.columns = ['Unidade / Setor', 'Quantidade']
                st.bar_chart(unit_counts.set_index('Unidade / Setor'), use_container_width=True)

            st.markdown("---")

            # Row 2: TAGs / Categorias de IA e Usuários com Mais Chamados
            g_col3, g_col4 = st.columns(2)
            with g_col3:
                st.markdown("#### 🏷️ Distribuição por Categoria (TAG de IA)")
                tag_counts = filtered_df['tag'].value_counts().reset_index()
                tag_counts.columns = ['Categoria (TAG)', 'Quantidade']
                st.bar_chart(tag_counts.set_index('Categoria (TAG)'), use_container_width=True)

            with g_col4:
                st.markdown("#### 👤 Top Usuários que Mais Abrem Chamados")
                user_counts = filtered_df['usuario'].value_counts().head(10).reset_index()
                user_counts.columns = ['Usuário', 'Quantidade']
                st.bar_chart(user_counts.set_index('Usuário'), use_container_width=True)
                
            st.markdown("---")

            # Row 3: Comparativo de Bases e Status
            g_col5, g_col6 = st.columns(2)
            with g_col5:
                st.markdown("#### 🔄 Origem dos Chamados (Base)")
                base_counts = filtered_df['base'].value_counts().reset_index()
                base_counts.columns = ['Base de Origem', 'Quantidade']
                st.bar_chart(base_counts.set_index('Base de Origem'), use_container_width=True)

            with g_col6:
                st.markdown("#### 📍 Status por Prédio / Cidade (Abertos x Fechados)")
                status_city = filtered_df.groupby(['cidade_predio', 'status']).size().unstack(fill_value=0)
                st.bar_chart(status_city.head(10), use_container_width=True)
        else:
            st.info("Sem chamados no filtro selecionado para renderizar gráficos.")

    @st.dialog("Detalhes do Chamado", width="large")
    def show_ticket_details(row):
        # Cabeçalho do chamado unificado em um Expander aberto por padrão para economizar espaço se necessário
        title = str(row.get('titulo', '')).strip()
        if title and title.lower() not in ["none", "nan", "null", ""]:
            header_text = f"🎫 Chamado #{row['id']} – {title}"
        else:
            header_text = f"🎫 Chamado #{row['id']}"
        
        with st.expander(header_text, expanded=True):
            # Cria as duas colunas para otimizar espaço vertical
            col1, col2 = st.columns(2)
            
            with col1:
                st.markdown("### 👤 Informações do Usuário")
                # Formata o Usuário com o sAMAccountName (id_cliente) se disponível
                user_display = str(row['usuario'])
                client_id = str(row.get('id_cliente', '')).strip()
                if client_id and client_id.lower() not in ["none", "nan", "null", ""]:
                    user_display += f" ({client_id})"
                    
                st.markdown(f"**Usuário:** {user_display}")
                st.markdown(f"**Localidade:** {row['localidade_fisica']}")
                st.markdown(f"**Base de Origem:** `{row['base']}`")
                st.markdown(f"**IP de Origem:** `{row.get('ip_origem') or 'N/A'}`")
                st.markdown(f"**Hostname:** `{row.get('hostname') or 'N/A'}`")
                
                with st.expander("📍 Editar Localização Manual", expanded=False):
                    new_cidade = st.text_input("Cidade - Prédio", value=str(row.get('cidade_predio', '')), key=f"edit_cidade_{row['id']}")
                    new_unidade = st.text_input("Unidade", value=str(row.get('unidade', '')), key=f"edit_unidade_{row['id']}")
                    new_localidade = st.text_input("Localidade Física", value=str(row.get('localidade_fisica', '')), key=f"edit_localidade_{row['id']}")
                    if st.button("💾 Salvar Localização", key=f"save_loc_btn_{row['id']}"):
                        from src.database import update_ticket_location_details
                        update_ticket_location_details(row['id'], new_localidade, new_cidade, new_unidade)
                        st.success("Localização salva! (Fechar para atualizar a tabela)")
                        st.cache_data.clear()
                
            with col2:
                st.markdown("### ⚙️ Classificação & Status")
                # Exibição da TAG com destaque colorido premium baseado nas cores oficiais
                tag_name = str(row['tag']).upper().strip()
                bg_color = TAG_COLORS.get(tag_name, "#262730")
                
                # Calcula contraste excelente calculando a luminância da cor de fundo
                hex_color = bg_color.lstrip('#')
                try:
                    r, g, b = int(hex_color[0:2], 16), int(hex_color[2:4], 16), int(hex_color[4:6], 16)
                    luminance = (0.299 * r + 0.587 * g + 0.114 * b) / 255
                    text_color = "#ffffff" if luminance < 0.6 else "#212529"
                except:
                    text_color = "#ffffff"
                    
                tag_html = f'<span style="background-color: {bg_color}; color: {text_color}; padding: 3px 8px; border-radius: 4px; font-weight: bold; font-family: inherit; font-size: 13px;">{row["tag"]}</span>'
                st.markdown(f"**TAG Atual:** {tag_html}", unsafe_allow_html=True)
                
                # Dropdown para alterar a TAG manualmente de forma rápida
                tag_options = sorted(list(TAG_COLORS.keys()))
                try:
                    default_idx = tag_options.index(tag_name)
                except ValueError:
                    default_idx = 0
                    
                new_tag = st.selectbox("🏷️ Alterar TAG Manualmente", options=tag_options, index=default_idx, key=f"select_tag_{row['id']}")
                if new_tag != tag_name:
                    if st.button("💾 Salvar Nova TAG", key=f"save_tag_btn_{row['id']}"):
                        from src.database import update_ticket_tag
                        update_ticket_tag(row['id'], new_tag)
                        st.success(f"TAG alterada com sucesso para {new_tag}! (Atualizará na tabela ao fechar o modal)")
                        st.cache_data.clear()


                
                # Alteração de status
                #status_options = ["Aberto", "Fechado"]
                #current_idx = status_options.index(row['status']) if row['status'] in status_options else 0
                #new_status = st.selectbox("Status", status_options, index=current_idx, key="status_select_modal")
                
                #if new_status != row['status']:
                #    if st.button("💾 Salvar Alteração de Status", key="save_status_btn_modal"):
                #        conn = sqlite3.connect(DB_PATH)
                #        cursor = conn.cursor()
                #        now = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
                #        cursor.execute("""
                #        UPDATE chamados 
                #        SET status = ?, data_atualizacao = ?
                #        WHERE id = ?
                #        """, (new_status, now, row['id']))
                #        conn.commit()
                #        conn.close()
                #        st.success(f"Status atualizado para {new_status}!")
                #        st.rerun()
                        
                # Botão para abrir o chamado original
                link_url = row.get('link')
                if link_url:
                    st.markdown("---")
                    st.link_button("🔗 Abrir Chamado Original", link_url, width="stretch")
        
        # Accordion para Andamento / Notas Rápidas
        with st.expander("📝 Andamento / Nota de Atendimento", expanded=True):
            current_andamento = str(row.get('andamento', '')).strip()
            if current_andamento.lower() in ["none", "nan", "null", ""]:
                current_andamento = ""
            new_andamento = st.text_area("Nota rápida sobre o andamento do chamado:", value=current_andamento, key="andamento_modal_ta")
            if st.button("💾 Salvar Nota de Andamento", key="save_andamento_modal_btn"):
                from src.database import update_ticket_andamento
                update_ticket_andamento(row['id'], new_andamento)
                st.success("Nota de andamento atualizada com sucesso! (Atualizará na tabela ao fechar o modal)")
                st.cache_data.clear()
        
        # Accordion para a Descrição
        with st.expander(f"📝 #1 - {row['Data Formatada']} (Descrição)", expanded=True):
            st.text(row['descricao'])
            
        # Comentários / Notas históricas
        from src.database import get_comments_by_ticket
        comments = get_comments_by_ticket(row['id'])
        if comments:
            st.markdown("### 💬 Histórico de Notas e Acompanhamentos")
            # Exibe cada comentário em um expander elegante
            for i, c in enumerate(comments, start=2):
                header = f"🕒 #{i} – {c['data']} – por {c['autor']}"
                with st.expander(header):
                    st.text(c['texto'])
            
        if st.button("Fechar", key="close_modal_btn"):
            st.rerun()

    # Colunas para exibir por padrão (Ordem padrão com Andamento inclusa)
    cols_to_show = [
        'id', 'status', 'tag', 'andamento', 'localidade_fisica', 
        'cidade_predio', 'unidade', 'usuario', 'datetime_obj', 'base'
    ]
        
    # Inclui colunas necessárias para geração de links e para estarem disponíveis no picker (como ip_origem)
    cols_to_generate = list(cols_to_show)
    if 'link' not in cols_to_generate and 'link' in filtered_df.columns:
        cols_to_generate.append('link')
    if 'ip_origem' not in cols_to_generate and 'ip_origem' in filtered_df.columns:
        cols_to_generate.append('ip_origem')
        
    df_display = filtered_df[cols_to_generate].copy()
    
    def format_id_link(row):
        cid = str(row['id']).strip()
        link = str(row.get('link', '')).strip() if 'link' in row else ''
        if not link or link.lower() in ["none", "nan", "null", ""]:
            # Fallbacks seguros
            if row['base'] == 'CitSmart':
                link = f"https://suporte.mpms.mp.br/citsmart/pages/serviceRequestIncident/serviceRequestIncident.load?iframe=true&language=pt-BR#/request?idRequest={cid}"
            else:
                link = "https://central.mpms.mp.br/otrs/index.pl"
        return f"{link}#id:{cid}"

    df_display['id'] = df_display.apply(format_id_link, axis=1)
    
    # As colunas finais passadas incluem ip_origem para estar disponível no picker
    cols_for_dataframe = list(cols_to_show)
    if 'ip_origem' in df_display.columns:
        cols_for_dataframe.append('ip_origem')
        
    df_final_display = df_display[cols_for_dataframe]
    
    with main_tab_list:
        col_tbl_head, col_tbl_cap = st.columns([3, 2])
        with col_tbl_head:
            st.subheader(f"📋 Lista de Chamados ({len(filtered_df)} registros)")
            st.write("Dica: Clique no **checkbox (caixinha de seleção)** no início de qualquer linha para abrir os Detalhes no Modal.")
        with col_tbl_cap:
            components.html("""
            <script src="https://cdnjs.cloudflare.com/ajax/libs/html2canvas/1.4.1/html2canvas.min.js"></script>
            <div style="text-align: right; padding-top: 5px;">
                <button id="btn-cap-tbl" onclick="captureTable()" style="
                    background: linear-gradient(135deg, #10b981 0%, #059669 100%);
                    color: white;
                    border: none;
                    padding: 8px 14px;
                    font-size: 0.85rem;
                    font-weight: bold;
                    border-radius: 6px;
                    cursor: pointer;
                    box-shadow: 0 2px 6px rgba(0,0,0,0.3);
                    transition: all 0.2s ease;
                ">
                    📸 Baixar Tabela em Alta Resolução (PNG)
                </button>
            </div>
            <script>
            function captureTable() {
                const btn = document.getElementById("btn-cap-tbl");
                btn.innerText = "⏳ Gerando PNG em alta resolução...";
                btn.disabled = true;

                const tableEl = window.parent.document.querySelector('div[data-testid="stDataFrame"]') || 
                                window.parent.document.querySelector('.stDataFrame');

                if (!tableEl) {
                    alert("Não foi possível encontrar a tabela na tela.");
                    btn.innerText = "📸 Baixar Tabela em Alta Resolução (PNG)";
                    btn.disabled = false;
                    return;
                }

                html2canvas(tableEl, {
                    scale: 2.5,
                    useCORS: true,
                    backgroundColor: "#0e1117",
                    logging: false
                }).then(canvas => {
                    const link = document.createElement("a");
                    const d = new Date();
                    const dateStr = d.getFullYear() + "-" + String(d.getMonth()+1).padStart(2,'0') + "-" + String(d.getDate()).padStart(2,'0');
                    link.download = `tabela_chamados_sti_${dateStr}.png`;
                    link.href = canvas.toDataURL("image/png");
                    link.click();

                    btn.innerText = "✅ Imagem Salva!";
                    setTimeout(() => {
                        btn.innerText = "📸 Baixar Tabela em Alta Resolução (PNG)";
                        btn.disabled = false;
                    }, 3000);
                }).catch(err => {
                    console.error("Erro no html2canvas:", err);
                    alert("Erro ao capturar tabela: " + err);
                    btn.innerText = "📸 Baixar Tabela em Alta Resolução (PNG)";
                    btn.disabled = false;
                });
            }
            </script>
            """, height=65)
        
        # Controle de estado para evitar loop do modal
        if "last_selected" not in st.session_state:
            st.session_state["last_selected"] = None
        
        # Função para colorir as linhas do DataFrame de acordo com as TAGs e suas cores oficiais do Excel
        def style_dataframe(row):
            tag = str(row.get('tag', '')).upper().strip()
            bg_color = TAG_COLORS.get(tag, "")
            if bg_color:
                # Garante contraste excelente calculando a luminância da cor de fundo
                hex_color = bg_color.lstrip('#')
                r, g, b = int(hex_color[0:2], 16), int(hex_color[2:4], 16), int(hex_color[4:6], 16)
                luminance = (0.299 * r + 0.587 * g + 0.114 * b) / 255
                text_color = "#ffffff" if luminance < 0.6 else "#212529"
                style = f"background-color: {bg_color}; color: {text_color};"
            else:
                style = ""
                
            return [style] * len(row)

        # Configuramos o st.dataframe com seleção nativa e estilização de cores (Altamente compatível)
        selection_event = st.dataframe(
            df_final_display.style.apply(style_dataframe, axis=1),
            column_order=cols_to_show, # Especifica quais colunas aparecem por padrão (oculta ip_origem)
            column_config={
                "id": st.column_config.LinkColumn("Chamado #", display_text=r".*#id:(.*)"),
                "status": st.column_config.TextColumn("Status"),
                "tag": st.column_config.TextColumn("TAG"),
                "andamento": st.column_config.TextColumn("Andamento"),
                "localidade_fisica": st.column_config.TextColumn("Localidade Física"),
                "cidade_predio": st.column_config.TextColumn("Cidade - Prédio"),
                "unidade": st.column_config.TextColumn("Unidade"),
                "usuario": st.column_config.TextColumn("Usuário"),
                "datetime_obj": st.column_config.DatetimeColumn("Data Criação", format="DD/MM/YYYY HH:mm:ss"),
                "ip_origem": st.column_config.TextColumn("IP de Origem"),
                "base": st.column_config.TextColumn("Base"),
            },
            hide_index=True,
            width="stretch",
            height=600,
            on_select="rerun",
            selection_mode="single-row",
            key="tabela_chamados_datagrid"
        )

        # Lógica para exibir o Modal baseado na seleção da linha
        selected_rows = selection_event.selection.rows if hasattr(selection_event, "selection") else []
        
        if selected_rows:
            current_selected = selected_rows[0]
            if st.session_state["last_selected"] != current_selected:
                st.session_state["last_selected"] = current_selected
                row_data = filtered_df.iloc[current_selected]
                show_ticket_details(row_data)
        else:
            st.session_state["last_selected"] = None

        # NOVO: Seção para geração rápida de resumo para WhatsApp com interpretação de IA e problema completo
        st.markdown("---")
        st.subheader("📲 Compartilhar Fila por WhatsApp")
        with st.expander("💬 Gerar Resumo Formatado (Pronto para copiar e enviar)", expanded=False):
            if filtered_df.empty:
                st.info("Nenhum chamado na fila filtrada.")
            else:
                # Opção de resumo com NLP Local
                usar_resumo_ia = st.checkbox(
                    "✨ Usar Resumos Inteligentes (NLP Local)", 
                    value=True, 
                    help="Usa Processamento de Linguagem Natural (spaCy) rodando totalmente local para resumir o chamado em poucas palavras."
                )
                
                def get_ai_diagnostico(tag, desc):
                    tag = str(tag).upper().strip()
                    desc = str(desc).strip()
                    
                    # Limpeza de saudações iniciais para extrair sintoma puro
                    import re
                    clean_desc = re.sub(r'^(bom dia|boa tarde|boa noite|ola|prezados|favor|solicito|gostaria de)\b.*?\n', '', desc, flags=re.IGNORECASE)
                    clean_desc = clean_desc.strip()
                    
                    sentences = re.split(r'[.!?\n]', clean_desc)
                    first_sentence = ""
                    for s in sentences:
                        s = s.strip()
                        if len(s) > 10:
                            first_sentence = s
                            break
                    if not first_sentence:
                        first_sentence = desc[:100]
                        if len(desc) > 100:
                            first_sentence += "..."
                            
                    diagnosticos = {
                        "BACKUP": "Cópia de segurança ou restauração de arquivos pendente.",
                        "EVENTO": "Suporte técnico para eventos ou solenidades institucionais.",
                        "FORMATAÇÃO": "Computador com lentidão extrema/travamento exigindo formatação e reinstalação de OS.",
                        "GARANTIA": "Defeito físico de fábrica em equipamento que exige acionamento de suporte terceirizado.",
                        "IMPRESSORA": "Instabilidade na fila de impressão local, papel atolado ou configuração de nova impressora de rede.",
                        "INSTALAÇÃO HARDWARE": "Necessidade de substituição física ou acréscimo de componente de hardware na máquina.",
                        "INSTALAÇÃO SOFTWARE": "Instalação, licenciamento ou atualização corretiva de programas corporativos.",
                        "MANUTENÇÃO": "Necessidade de intervenção mecânica/elétrica, limpeza interna ou reaperto de conexões físicas.",
                        "MONITOR": "Sem sinal de vídeo, tela preta, piscando ou distorcendo imagens de saída.",
                        "MUDANÇA": "Deslocamento físico completo de equipamentos de informática entre salas ou comarcas.",
                        "PREPARAÇÃO COMPUTADORES": "Configuração inicial de máquinas novas e perfis de rede para novos servidores.",
                        "REDE": "Ausência total de internet, falha de rede física ou lentidão no tráfego de dados locais.",
                        "SOLICITAÇÃO SSD": "Melhoria de desempenho físico de máquina lenta via substituição por disco de estado sólido (SSD).",
                        "SUPORTE": "Instruções de uso básico ou esclarecimento de dúvidas técnicas em sistemas internos.",
                        "TELEFONIA FIXA": "Aparelho de telefone mudo, ramal com ruídos/chiado ou necessidade de transferência de ramal.",
                        "VIAGEM": "Deslocamento programado da equipe STI para atendimento em promotoria regional externa.",
                        "VISTORIA CPDS": "Check-up preventivo completo nos servidores e no centro de processamento de dados local."
                    }
                    
                    diag = diagnosticos.get(tag, "Análise e resolução de ticket técnico STI.")
                    return f"🧠 *Possível Problema:* {diag}\n🩺 *Sintoma:* _{first_sentence}_"
     
                from src.database import get_comments_by_ticket
     
                lines = []
                lines.append("📋 *LISTA DE CHAMADOS STI - MPMS* 📋\n")
                for _, row in filtered_df.iterrows():
                    cid = str(row['id']).strip()
                    link = str(row.get('link', '')).strip()
                    if not link or link.lower() in ["none", "nan", "null", ""]:
                        if row['base'] == 'CitSmart':
                            link = f"https://suporte.mpms.mp.br/citsmart/pages/serviceRequestIncident/serviceRequestIncident.load?iframe=true&language=pt-BR#/request?idRequest={cid}"
                        else:
                            link = "https://central.mpms.mp.br/otrs/index.pl"
                    
                    user = str(row['usuario'])
                    loc = str(row['localidade_fisica'])
                    tag = str(row['tag'])
                    desc = str(row['descricao']).strip()
                    
                    # Recupera os comentários históricos do banco para enviar junto
                    comments_list = get_comments_by_ticket(row['id'])
                    comments_text = ""
                    comments_summary_input = ""
                    if comments_list:
                        comments_text = "💬 *Histórico de Acompanhamento:*"
                        comments_summary_input = "\n".join([f"- {c['data']} ({c['autor']}): {c['texto']}" for c in comments_list])
                        for i, c in enumerate(comments_list, start=1):
                            comments_text += f"\n  • #{i} [{c['data']}] – {c['autor']}: {c['texto']}"
                    
                    # Gera diagnóstico inteligente
                    diagnostico_ia = get_ai_diagnostico(tag, desc)
                    
                    lines.append(f"🎫 *Chamado #{cid}* ({row['base']})")
                    lines.append(f"👤 *Usuário:* {user}")
                    lines.append(f"📍 *Local:* {loc}")
                    lines.append(f"🏷️ *TAG:* {tag}")
                    lines.append(f"{diagnostico_ia}")
                    
                    if usar_resumo_ia:
                        resumo_nlp = summarize_ticket_locally(desc, comments_summary_input)
                        lines.append(f"📝 *Resumo Inteligente:* {resumo_nlp}")
                    else:
                        lines.append(f"📝 *Problema Completo:*")
                        lines.append(f"{desc}")
                        if comments_text:
                            lines.append(comments_text)
                    
                    lines.append(f"🔗 *Link Direto:* {link}")
                    lines.append("--------------------------------------------------")
                
                whats_text = "\n".join(lines)
                st.write("Dica: Use o botão de **copiar** no canto superior direito do bloco de código abaixo:")
                st.code(whats_text, language="text")
