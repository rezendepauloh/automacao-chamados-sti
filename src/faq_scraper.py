import sys
import sqlite3
import logging
from pathlib import Path
from bs4 import BeautifulSoup
import time

root_dir = Path(__file__).parent.parent
sys.path.insert(0, str(root_dir))
sys.path.insert(0, str(root_dir / "src"))

db_path = root_dir / "chamados.db"

logging.basicConfig(level=logging.INFO, format="%(asctime)s - %(levelname)s - %(message)s")

def init_faq_schema():
    """Garante que a tabela faqs possui a estrutura completa no SQLite."""
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
    cursor.execute("PRAGMA table_info(faqs)")
    columns = [col[1] for col in cursor.fetchall()]
    if "conteudo" not in columns:
        cursor.execute("ALTER TABLE faqs ADD COLUMN conteudo TEXT")
        conn.commit()
    conn.close()

def clean_sharepoint_html(raw_html: str, base_url: str = "https://ministeriopublicoms.sharepoint.com") -> str:
    """Limpa o HTML bruto do SharePoint mantendo formatação, imagens, vídeos e estilos de leitura."""
    soup = BeautifulSoup(raw_html, "html.parser")
    
    # Remove elementos desnecessários da interface do SharePoint mantendo video e iframe
    for tag in soup.find_all(["button", "svg", "script", "style", "nav"]):
        tag.decompose()

    # Trata elementos de vídeo do SharePoint (ex: Stream/EmbedVideoPreview)
    for video_div in soup.find_all(attrs={"id": "EmbedVideoPreview"}):
        video_tag = video_div.find("video")
        if video_tag and video_tag.get("src"):
            video_src = video_tag["src"]
            if video_src.startswith("/"):
                video_src = f"{base_url}{video_src}"
            video_html = f'<video controls style="max-width: 100%; border-radius: 8px; margin: 16px 0; border: 1px solid #343541;" src="{video_src}"></video>'
            video_div.replace_with(BeautifulSoup(video_html, "html.parser"))
        else:
            # Se for apenas a imagem de preview do vídeo com link ou dataset de vídeo
            img_tag = video_div.find("img")
            if img_tag:
                img_src = img_tag.get("data-sp-originalimgsrc") or img_tag.get("src") or ""
                if img_src.startswith("/"):
                    img_src = f"{base_url}{img_src}"
                # Renderiza a capa do vídeo acompanhada de um aviso/botão de play
                video_preview_html = f'''
                <div style="position: relative; margin: 16px 0;">
                    <img src="{img_src}" style="max-width: 100%; border-radius: 8px; display: block; filter: brightness(0.8);" />
                    <div style="position: absolute; top: 50%; left: 50%; transform: translate(-50%, -50%); background: rgba(255, 75, 75, 0.9); color: white; padding: 10px 20px; border-radius: 20px; font-weight: bold; box-shadow: 0 4px 10px rgba(0,0,0,0.5);">▶ Vídeo do SharePoint</div>
                </div>
                '''
                video_div.replace_with(BeautifulSoup(video_preview_html, "html.parser"))
        
    # Preserva e ajusta links de imagens do SharePoint
    for img in soup.find_all("img"):
        src = img.get("data-sp-originalimgsrc") or img.get("src") or ""
        
        # Se for um caminho relativo do SharePoint, transforma em URL absoluta
        if src.startswith("/"):
            src = f"{base_url}{src}"
        
        if src and not src.startswith("blob:"):
            img["src"] = src
            
        img["style"] = "max-width: 100%; height: auto; border-radius: 8px; margin: 16px 0; display: block; box-shadow: 0 4px 12px rgba(0,0,0,0.3);"

    # Garante que os links externos tenham target="_blank"
    for a in soup.find_all("a"):
        a["target"] = "_blank"
        a["style"] = "color: #ff4b4b; text-decoration: underline; font-weight: 500;"

    return str(soup)

def perform_microsoft_login(page):
    """Realiza o login na conta Microsoft / SharePoint se a tela de login for exibida."""
    try:
        from config import CITSMART_EMAIL, PASSWORD
        email, password = CITSMART_EMAIL, PASSWORD
    except Exception as e:
        logging.warning(f"Não foi possível obter credenciais de config: {e}")
        return

    # Preenche o e-mail
    try:
        page.wait_for_selector('input[name="loginfmt"]', timeout=8000)
        logging.info("Preenchendo e-mail da conta Microsoft...")
        page.fill('input[name="loginfmt"]', email)
        page.click('input[id="idSIButton9"]')
        time.sleep(1.5)
    except Exception:
        pass

    # Preenche a senha
    try:
        page.wait_for_selector('input[name="passwd"]', timeout=8000)
        logging.info("Preenchendo senha...")
        page.fill('input[name="passwd"]', password)
        page.click('input[id="idSIButton9"]')
        time.sleep(1.5)
    except Exception:
        pass

    # Pula o KMSI (manter conectado)
    try:
        page.wait_for_selector('#KmsiCheckboxField, input[id="idSIButton9"]', timeout=8000)
        logging.info("Pulando confirmação KMSI...")
        page.click('input[id="idSIButton9"]')
        time.sleep(2)
    except Exception:
        pass


def scrape_all_faqs():
    """Conecta ao SharePoint via Playwright usando autenticação automática e atualiza a coluna conteudo no SQLite."""
    try:
        from playwright.sync_api import sync_playwright
    except ImportError:
        logging.error("Biblioteca 'playwright' não encontrada. Execute 'pip install playwright' e 'playwright install chromium'.")
        return

    init_faq_schema()

    conn = sqlite3.connect(db_path)
    cursor = conn.cursor()
    cursor.execute("SELECT id, titulo, url FROM faqs")
    pending_faqs = cursor.fetchall()
    
    if not pending_faqs:
        logging.info("Nenhum FAQ encontrado no banco para raspar.")
        conn.close()
        return

    logging.info(f"Iniciando processo de raspagem para {len(pending_faqs)} FAQs...")

    with sync_playwright() as p:
        browser = p.chromium.launch(headless=True)
        context = browser.new_context()
        page = context.new_page()

        # Efetua login inicial no SharePoint usando a primeira URL
        first_url = pending_faqs[0][2]
        logging.info(f"🔑 Acessando portal para autenticação inicial...")
        page.goto(first_url, wait_until="domcontentloaded", timeout=40000)
        perform_microsoft_login(page)
        
        sucessos = 0
        erros = 0

        for faq_id, titulo, url in pending_faqs:
            try:
                logging.info(f"⏳ Raspando: '{titulo}' ({url})")
                page.goto(url, wait_until="networkidle", timeout=35000)
                
                # Se for redirecionado para login no meio do caminho, realiza login novamente
                if "login.microsoftonline.com" in page.url:
                    perform_microsoft_login(page)
                    page.goto(url, wait_until="networkidle", timeout=35000)

                # Tenta localizar o contêiner de conteúdo do SharePoint
                try:
                    page.wait_for_selector('div[data-automation-id="CanvasLayout"], .ck-content, [data-automation-id="CanvasZone"]', timeout=15000)
                except Exception:
                    pass

                element = (
                    page.query_selector('div[data-automation-id="CanvasLayout"]') or
                    page.query_selector('.ck-content') or
                    page.query_selector('[data-automation-id="CanvasZone"]')
                )

                if element:
                    raw_html = element.inner_html()
                    cleaned_html = clean_sharepoint_html(raw_html)
                    
                    cursor.execute("""
                        UPDATE faqs 
                        SET conteudo = ?, data_atualizacao = CURRENT_TIMESTAMP 
                        WHERE id = ?
                    """, (cleaned_html, faq_id))
                    conn.commit()
                    sucessos += 1
                    logging.info(f"✅ Salvo com sucesso: '{titulo}'")
                else:
                    erros += 1
                    logging.warning(f"⚠️ Não foi possível isolar o contêiner de texto em: '{titulo}'")
            except Exception as e:
                erros += 1
                logging.error(f"❌ Erro ao processar '{titulo}': {e}")

        browser.close()

    conn.close()
    logging.info(f"✨ Raspagem concluída! Sucessos: {sucessos} | Erros/Pendentes: {erros}")

if __name__ == "__main__":
    scrape_all_faqs()
