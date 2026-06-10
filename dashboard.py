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
    /* Reposiciona as abas de forma absoluta ao invés de fixas, alinhando nativamente com o menu lateral */
    div[data-testid="stRadio"] {
        position: absolute;
        top: -35px;
        left: 0px;
        z-index: 999999;
        background-color: transparent;
        margin: 0 !important;
        padding: 0 !important;
        width: max-content !important; /* Impede o colapso de largura na posição absoluta */
    }
    /* Alinha opções do radio em linha horizontal compacta */
    div[data-testid="stRadio"] [role="radiogroup"] {
        flex-direction: row !important;
        gap: 8px !important;
    }
    /* Oculta a bolinha padrão do radio button de forma precisa (apenas o filho direto contendo o círculo) */
    div[data-testid="stRadio"] label > div:first-child {
        display: none !important;
    }
    /* Estiliza as abas de forma premium e elegante no cabeçalho */
    div[data-testid="stRadio"] label {
        background-color: #1e1f25 !important;
        border: 1px solid #343541 !important;
        padding: 4px 16px !important;
        border-radius: 6px !important;
        cursor: pointer !important;
        transition: all 0.2s ease-in-out !important;
        margin: 0 !important;
        white-space: nowrap !important; /* Impede a quebra de linhas do texto */
    }
    div[data-testid="stRadio"] label:hover {
        border-color: #ff4b4b !important;
        background-color: #2a2b36 !important;
    }
    /* Destaca a aba ativa com a cor vermelha tema do painel */
    div[data-testid="stRadio"] label:has(input:checked) {
        background-color: #ff4b4b !important;
        color: white !important;
        border-color: #ff4b4b !important;
        font-weight: bold !important;
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


def render_mapa_page():
    """Renderiza a página/aba de Mapa & Localização."""
    from src.database import get_map_config, get_map_pins, save_map_config
    import json
    
    st.title("📍 Mapa & Localização de Chamados")
    st.write("Visualize no mapa/planta baixa a localização exata das salas de atendimento.")
    
    # 1. Seção de Importação de JSON
    with st.sidebar.expander("📥 Configurações & Upload JSON", expanded=False):
        st.write("Atualize a planta e os locais enviando um JSON formatado:")
        uploaded_file = st.file_uploader("Escolher arquivo JSON", type=["json"])
        if uploaded_file is not None:
            try:
                config_data = json.load(uploaded_file)
                if "predios" in config_data:
                    save_map_config(config_data)
                    st.success("Configurações do mapa e pins importadas com sucesso!")
                    st.cache_resource.clear()
                    st.cache_data.clear()
                    st.rerun()
                else:
                    st.error("JSON inválido! Deve conter a chave 'predios'.")
            except Exception as e:
                st.error(f"Erro ao processar arquivo: {e}")
                
    # 2. Carrega as configurações do banco
    config = get_map_config()
    predios = config.get("predios", [])
    
    if not predios:
        st.info("Nenhum prédio cadastrado no banco de dados. Faça o upload de um JSON de configurações na barra lateral.")
        return
        
    # Adiciona os seletores e busca diretamente na barra lateral, liberando espaço total para a imagem
    st.sidebar.markdown("---")
    st.sidebar.subheader("📍 Seleção do Local")
    
    # Seleção de prédio
    predio_nomes = [p.get("nome") for p in predios]
    selected_predio_nome = st.sidebar.selectbox("Selecione o Prédio", predio_nomes)
    selected_predio = next(p for p in predios if p.get("nome") == selected_predio_nome)
    predio_id = selected_predio.get("id")
    
    # Seleção de pavimento
    pavimentos = selected_predio.get("pavimentos", [])
    if not pavimentos:
        st.sidebar.warning("Sem pavimentos.")
        return
        
    pavimento_nomes = [pav.get("nome") for pav in pavimentos]
    selected_pav_nome = st.sidebar.selectbox("Selecione o Pavimento", pavimento_nomes)
    selected_pav = next(pav for pav in pavimentos if pav.get("nome") == selected_pav_nome)
    pavimento_id = selected_pav.get("id")
    
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
    
    # 3. Caixa de seleção de pins e busca
    st.sidebar.markdown("---")
    st.sidebar.subheader("🎯 Localização de Salas")
    
    # Seletor de Sala
    pin_nomes = ["-- Selecione uma Sala --"] + [p["sala"] for p in pins]
    selected_pin_nome = st.sidebar.selectbox("Ir para a Sala", pin_nomes)
    active_pin_id = ""
    
    if selected_pin_nome != "-- Selecione uma Sala --":
        active_pin = next(p for p in pins if p["sala"] == selected_pin_nome)
        active_pin_id = active_pin["id"]
        
    # Busca de sala (mantém como alternativa útil)
    search_query = st.sidebar.text_input("🔍 Buscar Sala ou Local (ex: TI, Protocolo)", "").strip()
    if search_query:
        matching_pins = [
            p for p in pins 
            if search_query.lower() in p.get("sala", "").lower() or search_query.lower() in p.get("descricao", "").lower()
        ]
        if matching_pins:
            st.sidebar.success(f"✨ Encontrado: {len(matching_pins)} correspondência(s)")
            active_pin_id = matching_pins[0].get("id")
        else:
            st.sidebar.warning("⚠️ Nenhum local encontrado.")
            
    # 4. Traçado de Rotas (Pathfinding)
    caminhos = selected_predio.get("caminhos", {})
    route_coords = []
    route_distance_meters = 0.0
    
    if caminhos and caminhos.get("nós") and caminhos.get("arestas"):
        st.sidebar.markdown("---")
        st.sidebar.subheader("🚶 Traçar Rota Interna")
        
        # Pega pins de todos os andares para origem/destino
        todos_pins = get_map_pins(predio_id)
        pin_origem_nomes = [f"{p['sala']} ({p['pavimento_id']}º Andar)" if p['pavimento_id'] > 0 else f"{p['sala']} (Térreo)" for p in todos_pins]
        
        selected_origem_display = st.sidebar.selectbox("Ponto de Origem", ["-- Selecione a Origem --"] + pin_origem_nomes)
        selected_destino_display = st.sidebar.selectbox("Ponto de Destino", ["-- Selecione o Destino --"] + pin_origem_nomes)
        
        if selected_origem_display != "-- Selecione a Origem --" and selected_destino_display != "-- Selecione o Destino --":
            orig_idx = pin_origem_nomes.index(selected_origem_display)
            dest_idx = pin_origem_nomes.index(selected_destino_display)
            
            orig_pin = todos_pins[orig_idx]
            dest_pin = todos_pins[dest_idx]
            
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
      </style>
    </head>
    <body>
      <div id="map" style="height: 650px; width: 100%;"></div>
      <script>
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

        // Pins
        var pins = {pins_json_str};
        var activePinId = "{active_pin_id}";

        pins.forEach(function(pin) {{
          var isActive = (pin.id === activePinId);
          
          // Estilo de marcador customizado usando HTML DivIcon do Leaflet para ficar bem premium
          var color = isActive ? "#ff4b4b" : "#4b9cff";
          var size = isActive ? "24px" : "16px";
          var border = isActive ? "3px solid white" : "2px solid white";
          
          var customIcon = L.divIcon({{
            className: 'custom-pin',
            html: '<div style="background-color: ' + color + '; width: ' + size + '; height: ' + size + '; border-radius: 50%; border: ' + border + '; box-shadow: 0 0 10px rgba(0,0,0,0.5);"></div>',
            iconSize: isActive ? [24, 24] : [16, 16],
            iconAnchor: isActive ? [12, 12] : [8, 8]
          }});

          var marker = L.marker([pin.y, pin.x], {{icon: customIcon}}).addTo(map);
          marker.bindPopup("<b>📌 " + pin.sala + "</b><br>" + pin.descricao);

          if (isActive) {{
            marker.openPopup();
            map.setView([pin.y, pin.x], 1);
          }}
        }});

        // =====================================================================
        // [DESENVOLVIMENTO] Renderização dos nós e arestas do pavimento ativo para apoio visual
        // =====================================================================
        var debugLayer = L.layerGroup().addTo(map);

        var activeNodes = {active_nodes_json_str};
        activeNodes.forEach(function(node) {{
          var nodeIcon = L.divIcon({{
            className: 'debug-node',
            html: '<div style="background-color: #8a2be2; width: 8px; height: 8px; border-radius: 50%; opacity: 0.5; box-shadow: 0 0 3px rgba(0,0,0,0.5);"></div>',
            iconSize: [8, 8],
            iconAnchor: [4, 4]
          }});
          var marker = L.marker([node.y, node.x], {{icon: nodeIcon}}).addTo(debugLayer);
          marker.bindTooltip("<b>Nó:</b> " + node.id + "<br>" + node.nome, {{sticky: true}});
        }});

        var activeArestas = {active_arestas_json_str};
        activeArestas.forEach(function(edge) {{
          var polyline = L.polyline([edge.de_coords, edge.para_coords], {{
            color: '#2ecc71', // Verde
            weight: 3,
            opacity: 0.4,
            dashArray: '5, 5'
          }}).addTo(debugLayer);
          polyline.bindTooltip("<b>Aresta:</b> " + edge.de_id + " ➔ " + edge.para_id + (edge.tipo !== 'caminho' ? " (" + edge.tipo + ")" : ""), {{sticky: true}});
        }});

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
            container.title = "Exibir Malha de Caminhos";

            // Ícone do Olho para Visibilidade
            container.innerHTML = '<span style="font-size: 16px; line-height: 1; filter: grayscale(100%);">👁️</span>';

            var isVisible = true;
            container.onclick = function(e) {{
              L.DomEvent.stopPropagation(e); // Previne clique de propagar para o mapa
              if (isVisible) {{
                map.removeLayer(debugLayer);
                container.style.opacity = '0.5';
              }} else {{
                map.addLayer(debugLayer);
                container.style.opacity = '1.0';
              }}
              isVisible = !isVisible;
            }};
            
            // Efeito hover
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
        // =====================================================================

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
          polyline.bindPopup("<b>🚶 Rota Interna Calculada</b><br>Distância total estimada: <b>" + {route_distance_meters:.1f} + " m</b>");
          
          // Enquadra a visão do mapa para englobar toda a rota percorrida
          map.fitBounds(polyline.getBounds());
        }}

        // Envia coordenadas de clique silenciosamente para o console.log (F12) se estiver dentro da planta
        map.on('click', function(e) {{
          var coord = e.latlng;
          var x = Math.round(coord.lng);
          var y = Math.round(coord.lat);
          
          // Filtra coordenadas válidas dentro das dimensões da foto
          if (x >= 0 && x <= w && y >= 0 && y <= h) {{
            var floor = {pavimento_id};
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
        }});
      </script>
    </body>
    </html>
    """
    
    st.iframe(leaflet_html, height=670)


# Navegação por abas/páginas no Topo (Flutua no Header via CSS Fixed)
page = st.radio(
    "Navegação",
    ["📋 Painel de Chamados", "📍 Mapa & Localização"],
    horizontal=True,
    label_visibility="collapsed"
)

if page == "📍 Mapa & Localização":
    render_mapa_page()
    st.stop()  # Interrompe a execução para não carregar a página padrão de chamados


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

    def get_filtered_options(col_name: str) -> list:
        """
        Calcula as opções únicas disponíveis para uma coluna específica,
        aplicando todos os filtros ativos, exceto o filtro da própria coluna.
        """
        temp_df = df.copy()
        
        # 1. Filtro de data
        dr = st.session_state.get("f_date_range", (min_date, max_date))
        if isinstance(dr, tuple) and len(dr) == 2:
            start_date, end_date = dr
            temp_df = temp_df[
                (temp_df['datetime_obj'].dt.date >= start_date) & 
                (temp_df['datetime_obj'].dt.date <= end_date)
            ]
            
        # 2. Aplica demais filtros (exceto o próprio)
        if col_name != 'status' and st.session_state.get("f_status"):
            temp_df = temp_df[temp_df['status'].isin(st.session_state["f_status"])]
            
        if col_name != 'tag' and st.session_state.get("f_tags"):
            temp_df = temp_df[temp_df['tag'].isin(st.session_state["f_tags"])]
            
        if col_name != 'localidade_fisica' and st.session_state.get("custom_loc_selection"):
            temp_df = temp_df[temp_df['localidade_fisica'].isin(st.session_state["custom_loc_selection"])]
            
        if col_name != 'cidade_predio' and st.session_state.get("f_cities"):
            temp_df = temp_df[temp_df['cidade_predio'].isin(st.session_state["f_cities"])]
            
        if col_name != 'unidade' and st.session_state.get("f_units"):
            temp_df = temp_df[temp_df['unidade'].isin(st.session_state["f_units"])]
            
        if col_name != 'base' and st.session_state.get("f_bases"):
            temp_df = temp_df[temp_df['base'].isin(st.session_state["f_bases"])]
            
        if col_name != 'usuario' and st.session_state.get("f_user"):
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

    # Layout do título de Filtros e do Botão de Limpar
    col_h, col_b = st.sidebar.columns([5, 4])
    with col_h:
        st.markdown("### Filtros")
    with col_b:
        # Botão menor e elegante alinhado horizontalmente
        if st.button("🧹 Limpar", help="Limpa todos os filtros ativos de uma vez"):
            st.session_state["f_date_range"] = (min_date, max_date)
            st.session_state["f_status"] = []
            st.session_state["f_tags"] = []
            st.session_state["custom_loc_selection"] = []
            st.session_state["f_cities"] = []
            st.session_state["f_units"] = []
            st.session_state["f_bases"] = []
            st.session_state["f_user"] = ""
            st.rerun()
            
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
    status_options = get_filtered_options('status')
    selected_status = st.sidebar.multiselect(
        "Status", 
        options=status_options, 
        key="f_status", 
        placeholder="Escolha as opções..."
    )
    
    tag_options = get_filtered_options('tag')
    selected_tags = st.sidebar.multiselect(
        "TAG", 
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
    
    city_options = get_filtered_options('cidade_predio')
    selected_cities = st.sidebar.multiselect(
        "Cidade - Prédio", 
        options=city_options, 
        key="f_cities", 
        placeholder="Escolha as opções..."
    )
    
    unit_options = get_filtered_options('unidade')
    selected_units = st.sidebar.multiselect(
        "Unidade", 
        options=unit_options, 
        key="f_units", 
        placeholder="Escolha as opções..."
    )
    
    # Filtro de Base (CitSmart/OTRS)
    base_options = get_filtered_options('base')
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
    
    st.sidebar.markdown("---")
    

    # Aplica os filtros
    filtered_df = df.copy()
    
    # Filtro de data
    if isinstance(date_range, tuple) and len(date_range) == 2:
        start_date, end_date = date_range
        filtered_df = filtered_df[
            (filtered_df['datetime_obj'].dt.date >= start_date) & 
            (filtered_df['datetime_obj'].dt.date <= end_date)
        ]
        
    if selected_status:
        filtered_df = filtered_df[filtered_df['status'].isin(selected_status)]
    if selected_tags:
        filtered_df = filtered_df[filtered_df['tag'].isin(selected_tags)]
    if selected_locs:
        filtered_df = filtered_df[filtered_df['localidade_fisica'].isin(selected_locs)]
    if selected_cities:
        filtered_df = filtered_df[filtered_df['cidade_predio'].isin(selected_cities)]
    if selected_units:
        filtered_df = filtered_df[filtered_df['unidade'].isin(selected_units)]
    if selected_bases:
        filtered_df = filtered_df[filtered_df['base'].isin(selected_bases)]
    if user_search:
        filtered_df = filtered_df[filtered_df['usuario'].str.contains(user_search, case=False, na=False)]
        
    # Exibe métricas
    col1, col2, col3 = st.columns(3)
    col1.metric("Total de Chamados", len(filtered_df))
    col2.metric("Abertos", len(filtered_df[filtered_df['status'] == 'Aberto']))
    col3.metric("Fechados", len(filtered_df[filtered_df['status'] == 'Fechado']))
    
    st.write("---")
    
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
    
    st.subheader("Lista de Chamados")
    st.write("Dica: Clique no **checkbox (caixinha de seleção)** no início de qualquer linha na tabela abaixo para abrir os Detalhes e Descrição no Modal.")
    
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
