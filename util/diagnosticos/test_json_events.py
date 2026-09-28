import sys
from pathlib import Path
import json

ROOT_DIR = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(ROOT_DIR))
sys.path.insert(0, str(ROOT_DIR / "src"))

from src.database import load_data, get_comments_by_ticket
from src.tabs.calendario_geral import parse_ticket_date_iso_and_br

def test_events_json():
    df = load_data()
    events = []
    
    for _, row in df.iterrows():
        cid = str(row.get('id', '')).strip()
        titulo = str(row.get('titulo', '')).strip()
        if not titulo or titulo.lower() in ["none", "nan", "null"]:
            titulo = "Sem Titulo"
            
        status = str(row.get('status', 'Aberto')).strip()
        tag = str(row.get('tag', '')).strip()
        usuario = str(row.get('usuario', '')).strip()
        localidade = str(row.get('localidade_fisica', '')).strip()
        unidade = str(row.get('unidade', '')).strip()
        descricao = str(row.get('descricao', '')).strip()
        dt_criacao_raw = row.get('data_criacao')

        raw_comments = get_comments_by_ticket(cid)
        formatted_comments = []
        if raw_comments:
            for c in raw_comments:
                dt_c = str(c.get('data', '')).strip()
                aut_c = str(c.get('autor', '')).strip()
                txt_c = str(c.get('texto', '')).strip()
                if txt_c:
                    formatted_comments.append(f"{dt_c} - por {aut_c}: {txt_c}")

        iso_dt, br_dt = parse_ticket_date_iso_and_br(dt_criacao_raw)
        if not iso_dt:
            continue

        desc_resumo = (descricao[:350] + "...") if len(descricao) > 350 else descricao

        events.append({
            "id": f"chamado_{cid}",
            "title": f"#{cid}: {titulo[:35]}",
            "start": iso_dt,
            "backgroundColor": "#0ea5e9",
            "borderColor": "#0284c7",
            "extendedProps": {
                "categoria_evento": "chamado",
                "tipo": "Chamado OTRS",
                "id": cid,
                "base": "OTRS",
                "titulo": titulo,
                "tag": tag,
                "status": status,
                "usuario": usuario,
                "localidade": localidade,
                "unidade": unidade,
                "data_criacao": br_dt if br_dt else str(dt_criacao_raw),
                "descricao": desc_resumo if desc_resumo else "Sem descricao.",
                "comentarios": formatted_comments
            }
        })
        
    print(f"Total de eventos de chamados gerados: {len(events)}")
    try:
        json_str = json.dumps(events, ensure_ascii=False).replace("<", "\\u003c").replace(">", "\\u003e")
        print("JSON gerado com SUCESSO! Tamanho:", len(json_str), "bytes")
    except Exception as e:
        print("ERRO na geracao do JSON:", e)

if __name__ == "__main__":
    test_events_json()
