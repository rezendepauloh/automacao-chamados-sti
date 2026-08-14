import json
from .connection import get_connection

def setup_map_tables():
    """Cria as tabelas do mapa se não existirem."""
    conn = get_connection()
    cursor = conn.cursor()
    
    cursor.execute("""
    CREATE TABLE IF NOT EXISTS mapa_config (
        id TEXT PRIMARY KEY,
        config_json TEXT
    )
    """)
    
    cursor.execute("""
    CREATE TABLE IF NOT EXISTS mapa_pins (
        id TEXT PRIMARY KEY,
        predio_id TEXT,
        pavimento_id INTEGER,
        sala TEXT,
        x INTEGER,
        y INTEGER,
        descricao TEXT
    )
    """)
    conn.commit()
    conn.close()

def save_map_config(config_data: dict):
    """Salva as configurações de prédios e os pins no banco de dados SQLite."""
    setup_map_tables()
    conn = get_connection()
    cursor = conn.cursor()
    
    predios = config_data.get("predios", [])
    config_json_str = json.dumps({"predios": predios})
    cursor.execute("INSERT OR REPLACE INTO mapa_config (id, config_json) VALUES ('config_atual', ?)", (config_json_str,))
    
    cursor.execute("DELETE FROM mapa_pins")
    
    for predio in predios:
        p_id = predio.get("id")
        p_pins = predio.get("pins", [])
        for pin in p_pins:
            cursor.execute("""
            INSERT OR REPLACE INTO mapa_pins (id, predio_id, pavimento_id, sala, x, y, descricao)
            VALUES (?, ?, ?, ?, ?, ?, ?)
            """, (
                str(pin.get("id")),
                str(pin.get("predio_id", p_id)),
                int(pin.get("pavimento_id")),
                str(pin.get("sala")),
                int(pin.get("x")),
                int(pin.get("y")),
                str(pin.get("descricao", ""))
            ))
            
    pins = config_data.get("pins", [])
    for pin in pins:
        cursor.execute("""
        INSERT OR REPLACE INTO mapa_pins (id, predio_id, pavimento_id, sala, x, y, descricao)
        VALUES (?, ?, ?, ?, ?, ?, ?)
        """, (
            str(pin.get("id")),
            str(pin.get("predio_id")),
            int(pin.get("pavimento_id")),
            str(pin.get("sala")),
            int(pin.get("x")),
            int(pin.get("y")),
            str(pin.get("descricao", ""))
        ))
        
    conn.commit()
    conn.close()

def get_map_config() -> dict:
    """Retorna a configuração atual de prédios e pavimentos."""
    setup_map_tables()
    conn = get_connection()
    cursor = conn.cursor()
    cursor.execute("SELECT config_json FROM mapa_config WHERE id = 'config_atual'")
    row = cursor.fetchone()
    conn.close()
    if row:
        return json.loads(row[0])
    return {"predios": []}

def get_map_pins(predio_id=None, pavimento_id=None) -> list:
    """Retorna os pins cadastrados, podendo filtrar por prédio e pavimento."""
    setup_map_tables()
    conn = get_connection()
    cursor = conn.cursor()
    
    if predio_id is not None and pavimento_id is not None:
        cursor.execute("""
        SELECT id, predio_id, pavimento_id, sala, x, y, descricao 
        FROM mapa_pins 
        WHERE predio_id = ? AND pavimento_id = ?
        """, (predio_id, pavimento_id))
    elif predio_id is not None:
        cursor.execute("""
        SELECT id, predio_id, pavimento_id, sala, x, y, descricao 
        FROM mapa_pins 
        WHERE predio_id = ?
        """, (predio_id,))
    elif pavimento_id is not None:
        cursor.execute("""
        SELECT id, predio_id, pavimento_id, sala, x, y, descricao 
        FROM mapa_pins 
        WHERE pavimento_id = ?
        """, (pavimento_id,))
    else:
        cursor.execute("""
        SELECT id, predio_id, pavimento_id, sala, x, y, descricao 
        FROM mapa_pins
        """)
        
    rows = cursor.fetchall()
    conn.close()
    
    return [
        {
            "id": r[0],
            "predio_id": r[1],
            "pavimento_id": r[2],
            "sala": r[3],
            "x": r[4],
            "y": r[5],
            "descricao": r[6]
        }
        for r in rows
    ]
