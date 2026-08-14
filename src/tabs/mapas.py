import sys
import asyncio
import json
import heapq
import math
import threading
from pathlib import Path
from http.server import BaseHTTPRequestHandler, HTTPServer
import streamlit as st
import streamlit.components.v1 as components
from src.config import DEBUG_DIR_LEAFLET, setup_logging

logger = setup_logging(DEBUG_DIR_LEAFLET / "leaflet.log", "leaflet")

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


def calculate_dijkstra_route(caminhos: dict, start_pin: dict, end_pin: dict) -> list:
    """
    Calcula a rota mais curta entre o pin de origem e o pin de destino
    usando a malha de caminhos (nós e arestas) com o algoritmo de Dijkstra.
    """
    nos = caminhos.get("nós", [])
    if not nos:
        return []
        
    start_no = None
    min_start_dist = float("inf")
    for no in nos:
        if no["pavimento_id"] == start_pin["pavimento_id"]:
            dist = math.sqrt((no["x"] - start_pin["x"])**2 + (no["y"] - start_pin["y"])**2)
            if dist < min_start_dist:
                min_start_dist = dist
                start_no = no
                
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
        
    nodes_map = {n["id"]: n for n in nos}
    adj = {nid: [] for nid in nodes_map}
    
    for edge in caminhos.get("arestas", []):
        u = edge.get("de")
        v = edge.get("para")
        if u in nodes_map and v in nodes_map:
            n1 = nodes_map[u]
            n2 = nodes_map[v]
            
            if n1["pavimento_id"] != n2["pavimento_id"]:
                weight = 300.0
            else:
                weight = math.sqrt((n1["x"] - n2["x"])**2 + (n1["y"] - n2["y"])**2)
                
            adj[u].append((v, weight))
            adj[v].append((u, weight))
            
    queue = [(0.0, start_no["id"], [start_no["id"]])]
    visited = set()
    
    while queue:
        dist, curr, path = heapq.heappop(queue)
        if curr in visited:
            continue
        visited.add(curr)
        
        if curr == end_no["id"]:
            return [nodes_map[nid] for nid in path]
            
        for neighbor, weight in adj[curr]:
            if neighbor not in visited:
                heapq.heappush(queue, (dist + weight, neighbor, path + [neighbor]))
                
    return []


class SaveConfigHandler(BaseHTTPRequestHandler):
    def do_OPTIONS(self):
        self.send_response(200)
        self.send_header('Access-Control-Allow-Origin', '*')
        self.send_header('Access-Control-Allow-Methods', 'POST, GET, OPTIONS')
        self.send_header('Access-Control-Allow-Headers', 'Content-Type')
        self.end_headers()

    def do_GET(self):
        if not hasattr(st, "_global_route"):
            st._global_route = {"origem": "", "destino": ""}
            
        from urllib.parse import urlparse, parse_qs
        parsed = urlparse(self.path)
        logger.debug(f"🌐 GET request recebida no servidor Leaflet backend: {parsed.path}")
        if parsed.path == '/set_route':
            query = parse_qs(parsed.query)
            logger.debug(f"📍 Parâmetros query recebidos: {query}")
            
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
            
            try:
                from streamlit.runtime import get_instance
                runtime = get_instance()
                active_sessions = runtime._session_mgr.list_active_sessions()
                logger.debug(f"👥 Total de sessões Streamlit ativas encontradas: {len(active_sessions)}")
                for session_info in active_sessions:
                    session_state = session_info.session.session_state
                    
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
    
    if not hasattr(st, "_global_route"):
        st._global_route = {"origem": "", "destino": ""}
    url_origem = st._global_route.get("origem", "")
    url_destino = st._global_route.get("destino", "")
    
    st.title("📍 Mapa & Localização de Chamados")
    st.write("Visualize no mapa/planta baixa a localização exata das salas de atendimento.")
    
    config = get_map_config()
    predios = config.get("predios", [])
    
    if not predios:
        st.info("Nenhum prédio cadastrado no banco de dados. Faça o upload de um JSON de configurações na barra lateral.")
        return
        
    predio_nomes = [p.get("nome") for p in predios]
    selected_predio_nome = st.sidebar.selectbox("Selecione o Prédio", predio_nomes)
    selected_predio = next(p for p in predios if p.get("nome") == selected_predio_nome)
    predio_id = selected_predio.get("id")
    logger.debug(f"🏢 Prédio selecionado na UI: {selected_predio_nome} (ID: {predio_id})")
    
    todos_pins = get_map_pins(predio_id)
    if url_origem:
        orig_match = next((p for p in todos_pins if p["id"] == url_origem), None)
        if orig_match:
            display_name = f"{orig_match['sala']} ({orig_match['pavimento_id']}º Andar)" if orig_match['pavimento_id'] > 0 else f"{orig_match['sala']} (Térreo)"
            st.session_state.sb_origem = display_name
    elif "sb_origem" not in st.session_state:
        st.session_state.sb_origem = "-- Selecione a Origem --"
        
    if url_destino:
        dest_match = next((p for p in todos_pins if p["id"] == url_destino), None)
        if dest_match:
            display_name = f"{dest_match['sala']} ({dest_match['pavimento_id']}º Andar)" if dest_match['pavimento_id'] > 0 else f"{dest_match['sala']} (Térreo)"
            st.session_state.sb_destino = display_name
    elif "sb_destino" not in st.session_state:
        st.session_state.sb_destino = "-- Selecione o Destino --"
    
    pavimentos = selected_predio.get("pavimentos", [])
    if not pavimentos:
        st.sidebar.warning("Sem pavimentos.")
        return
        
    pavimento_nomes = [pav.get("nome") for pav in pavimentos]
    
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
    
    img_path_str = selected_pav.get("imagem")
    img_path = Path(img_path_str)
    
    if not img_path.exists():
        st.error(f"Imagem da planta baixa não encontrada no caminho: `{img_path_str}`")
        return
        
    w, h = get_image_dimensions(img_path)
    b64_image = get_image_base64(img_path)
    
    if not b64_image:
        st.error("Erro ao processar a imagem da planta baixa.")
        return
        
    pins = get_map_pins(predio_id, pavimento_id)
    
    st.sidebar.markdown("---")
    
    col_sub, col_clear = st.sidebar.columns([2, 1])
    with col_sub:
        st.subheader("🎯 Salas")
    with col_clear:
        st.markdown("<div style='height: 5px;'></div>", unsafe_allow_html=True)
        if st.button("🧹 Limpar", width='stretch'):
            st._global_route["origem"] = ""
            st._global_route["destino"] = ""
            st.session_state.sb_sala = "-- Selecione uma Sala --"
            st.session_state.txt_busca = ""
            st.rerun()
            
    pin_nomes = ["-- Selecione uma Sala --"] + [p["sala"] for p in pins]
    
    default_sb_index = 0
    if "sb_sala" in st.session_state and st.session_state.sb_sala in pin_nomes:
        default_sb_index = pin_nomes.index(st.session_state.sb_sala)
        
    selected_pin_nome = st.sidebar.selectbox("Ir para a Sala", pin_nomes, index=default_sb_index, key="sb_sala")
    active_pin_ids = []
    
    if selected_pin_nome != "-- Selecione uma Sala --":
        active_pin = next(p for p in pins if p["sala"] == selected_pin_nome)
        active_pin_ids.append(active_pin["id"])
        
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
            
            first_match = matching_pins[0]
            if first_match.get("pavimento_id") != pavimento_id and search_query != "":
                st.session_state.selected_pavimento_id = first_match.get("pavimento_id")
                st.rerun()
        else:
            st.sidebar.warning("⚠️ Nenhum local encontrado.")
            
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
            if st.button("🧹 Limpar", key="btn_limpar_rota", width='stretch'):
                st._global_route["origem"] = ""
                st._global_route["destino"] = ""
                st.session_state.sb_origem = "-- Selecione a Origem --"
                st.session_state.sb_destino = "-- Selecione o Destino --"
                st.rerun()
        
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
                    total_dist_pixels = 0.0
                    for idx_n in range(len(route_nodes) - 1):
                        n1 = route_nodes[idx_n]
                        n2 = route_nodes[idx_n+1]
                        if n1["pavimento_id"] != n2["pavimento_id"]:
                            total_dist_pixels += 100.0
                        else:
                            total_dist_pixels += math.sqrt((n1["x"] - n2["x"])**2 + (n1["y"] - n2["y"])**2)
                    
                    route_distance_meters = total_dist_pixels * 0.05
                    st.sidebar.success(f"🎉 Rota calculada com sucesso! ({route_distance_meters:.1f} m)")
                    
                    active_floor_nodes = [n for n in route_nodes if n["pavimento_id"] == pavimento_id]
                    
                    coords_to_draw = []
                    if orig_pin["pavimento_id"] == pavimento_id:
                        coords_to_draw.append([orig_pin["y"], orig_pin["x"]])
                        
                    for n in active_floor_nodes:
                        coords_to_draw.append([n["y"], n["x"]])
                        
                    if dest_pin["pavimento_id"] == pavimento_id:
                        coords_to_draw.append([dest_pin["y"], dest_pin["x"]])
                        
                    route_coords = coords_to_draw
                    
                    outros_andares = [n for n in route_nodes if n["pavimento_id"] != pavimento_id]
                    if outros_andares:
                        st.sidebar.warning("⚠️ Rota exige mudança de pavimento! Siga até a escada/elevador e alterne para o pavimento destino para ver a continuação.")
                else:
                    st.sidebar.error("Não foi possível calcular uma rota válida.")

    active_nodes = []
    if caminhos and caminhos.get("nós"):
        active_nodes = [n for n in caminhos.get("nós", []) if n.get("pavimento_id") == pavimento_id]
    active_nodes_json_str = json.dumps(active_nodes)

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
      
      <link rel="stylesheet" href="https://unpkg.com/leaflet@1.9.4/dist/leaflet.css" />
      <script src="https://unpkg.com/leaflet@1.9.4/dist/leaflet.js"></script>
      
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

        /* TEMA ESCURO (DARK MODE) */
        body, body.dark-mode {{
          background-color: #0e1117;
          color: #ffffff;
        }}
        body.dark-mode #map {{
          background: #0e1117;
          border: 1px solid #464855;
        }}
        body.dark-mode .leaflet-popup-content-wrapper,
        body.dark-mode .leaflet-popup-tip {{
          background: #1e1f25 !important;
          color: #ffffff !important;
          border: 1px solid #464855 !important;
          box-shadow: 0 4px 15px rgba(0,0,0,0.5) !important;
        }}

        /* TEMA CLARO (LIGHT MODE) */
        body.light-mode {{
          background-color: #ffffff;
          color: #0f172a;
        }}
        body.light-mode #map {{
          background: #f8fafc;
          border: 1px solid #cbd5e1;
        }}
        body.light-mode .leaflet-popup-content-wrapper,
        body.light-mode .leaflet-popup-tip {{
          background: #ffffff !important;
          color: #0f172a !important;
          border: 1px solid #cbd5e1 !important;
          box-shadow: 0 4px 15px rgba(0,0,0,0.1) !important;
        }}

        #map {{
          height: 100%;
          width: 100%;
          margin: 0;
          padding: 0;
          border-radius: 8px;
          box-sizing: border-box;
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
        function updateThemeFromParent() {{
          var isLight = false;
          try {{
            var parentBody = window.parent.document.body;
            var parentApp = window.parent.document.querySelector('.stApp');
            var themeAttr = (parentBody && parentBody.getAttribute('data-theme')) || 
                            (parentApp && parentApp.getAttribute('data-theme'));
            
            if (themeAttr === 'light') {{
              isLight = true;
            }} else if (themeAttr === 'dark') {{
              isLight = false;
            }} else {{
              isLight = window.parent.matchMedia('(prefers-color-scheme: light)').matches;
            }}
          }} catch(e) {{
            isLight = window.matchMedia('(prefers-color-scheme: light)').matches;
          }}

          if (isLight) {{
            document.body.className = 'light-mode';
          }} else {{
            document.body.className = 'dark-mode';
          }}
        }}

        document.addEventListener('DOMContentLoaded', function() {{
          updateThemeFromParent();
          setInterval(updateThemeFromParent, 1000);
        }});

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

        var map = L.map('map', {{
          crs: L.CRS.Simple,
          minZoom: -2,
          maxZoom: 3,
          attributionControl: false,
          maxBounds: bounds,
          maxBoundsViscosity: 1.0
        }});
        
        var image = L.imageOverlay('{b64_image}', bounds).addTo(map);
        map.fitBounds(bounds);

        map.addControl(new L.Control.Fullscreen({{
          position: 'topright',
          title: {{
            'false': 'Ver em Tela Cheia',
            'true': 'Sair da Tela Cheia'
          }}
        }}));

        var activePinIds = {active_pin_ids_json_str};
        var activeBuildingId = "{predio_id}";
        var floorId = {pavimento_id};
        var fullConfig = {config_json_str};

        var pinsLayer = L.layerGroup().addTo(map);
        var debugLayer = L.layerGroup();

        window.redrawAllLayers = function() {{
          pinsLayer.clearLayers();
          debugLayer.clearLayers();

          var predio = fullConfig.predios.find(function(p) {{ return p.id === activeBuildingId; }});
          if (!predio) return;

          if (!predio.caminhos) predio.caminhos = {{ "nós": [], "arestas": [] }};
          if (!predio.caminhos.nós) predio.caminhos.nós = [];
          if (!predio.caminhos.arestas) predio.caminhos.arestas = [];
          if (!predio.pins) predio.pins = [];

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

            marker.on('dragend', function(e) {{
              var latlng = marker.getLatLng();
              node.x = Math.round(latlng.lng);
              node.y = Math.round(latlng.lat);
              if (!devState.unsavedElements.includes(node.id)) {{
                devState.unsavedElements.push(node.id);
                sessionStorage.setItem('dev_unsavedElements', JSON.stringify(devState.unsavedElements));
              }}
              window.redrawAllLayers();
            }});

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

            marker.on('dragend', function(e) {{
              var latlng = marker.getLatLng();
              pin.x = Math.round(latlng.lng);
              pin.y = Math.round(latlng.lat);
              if (!devState.unsavedElements.includes(pin.id)) {{
                devState.unsavedElements.push(pin.id);
                sessionStorage.setItem('dev_unsavedElements', JSON.stringify(devState.unsavedElements));
              }}
              window.redrawAllLayers();
            }});
          }});
        }};

        window.redrawAllLayers();

        var routeCoords = {route_coords_json_str};
        if (routeCoords && routeCoords.length > 1) {{
          var routePolyline = L.polyline(routeCoords, {{
            color: '#ff4b4b',
            weight: 5,
            opacity: 0.9,
            lineCap: 'round',
            lineJoin: 'round'
          }}).addTo(map);

          routePolyline.bindTooltip("🚶 <b>Rota Sugerida</b> ({route_distance_meters:.1f}m)", {{sticky: true}});
        }}

        window.setRouteOrigin = function(pinId) {{
          fetch('http://localhost:8099/set_route?origem=' + encodeURIComponent(pinId))
            .then(res => res.json())
            .then(data => {{ console.log("Origem setada:", pinId); }});
        }};

        window.setRouteDestination = function(pinId) {{
          fetch('http://localhost:8099/set_route?destino=' + encodeURIComponent(pinId))
            .then(res => res.json())
            .then(data => {{ console.log("Destino setado:", pinId); }});
        }};
      </script>
    </body>
    </html>
    """
    
    leaflet_html += f"\n<!-- key: map_{url_origem}_{url_destino}_{pavimento_id} -->"
    st.components.v1.html(leaflet_html, height=670)
