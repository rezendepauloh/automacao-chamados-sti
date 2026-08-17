import sys
import sqlite3
import logging
from pathlib import Path
from bs4 import BeautifulSoup
import time

# Adiciona o diretório raiz e o diretório src ao sys.path para suportar importações diretas
root_dir = Path(__file__).resolve().parent.parent.parent
src_dir = Path(__file__).resolve().parent.parent
if str(root_dir) not in sys.path:
    sys.path.insert(0, str(root_dir))
if str(src_dir) not in sys.path:
    sys.path.insert(0, str(src_dir))

db_path = root_dir / "chamados.db"

from src.config import setup_logging, DEBUG_DIR_FAQ
from src.terminal import log, print_header, CYAN, GREEN, RED, YELLOW, WHITE

logger = setup_logging(DEBUG_DIR_FAQ / "faq_scraper.log", __name__)


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
    
    for tag in soup.find_all(["button", "svg", "script", "style", "nav"]):
        tag.decompose()

    for video_div in soup.find_all(attrs={"id": "EmbedVideoPreview"}):
        video_tag = video_div.find("video")
        if video_tag and video_tag.get("src"):
            video_src = video_tag["src"]
            if video_src.startswith("/"):
                video_src = f"{base_url}{video_src}"
            video_html = f'<video controls style="max-width: 100%; border-radius: 8px; margin: 16px 0; border: 1px solid #343541;" src="{video_src}"></video>'
            video_div.replace_with(BeautifulSoup(video_html, "html.parser"))
        else:
            img_tag = video_div.find("img")
            if img_tag:
                img_src = img_tag.get("data-sp-originalimgsrc") or img_tag.get("src") or ""
                if img_src.startswith("/"):
                    img_src = f"{base_url}{img_src}"
                video_preview_html = f'''
                <div style="position: relative; margin: 16px 0;">
                    <img src="{img_src}" style="max-width: 100%; border-radius: 8px; display: block; filter: brightness(0.8);" />
                    <div style="position: absolute; top: 50%; left: 50%; transform: translate(-50%, -50%); background: rgba(255, 75, 75, 0.9); color: white; padding: 10px 20px; border-radius: 20px; font-weight: bold; box-shadow: 0 4px 10px rgba(0,0,0,0.5);">▶ Vídeo do SharePoint</div>
                </div>
                '''
                video_div.replace_with(BeautifulSoup(video_preview_html, "html.parser"))
        
    for img in soup.find_all("img"):
        src = img.get("data-sp-originalimgsrc") or img.get("src") or ""
        
        if src.startswith("/"):
            src = f"{base_url}{src}"
        
        if src and not src.startswith("blob:"):
            img["src"] = src
            
        img["style"] = "max-width: 100%; height: auto; border-radius: 8px; margin: 16px 0; display: block; box-shadow: 0 4px 12px rgba(0,0,0,0.3);"

    for a in soup.find_all("a"):
        a["target"] = "_blank"
        a["style"] = "color: #ff4b4b; text-decoration: underline; font-weight: 500;"

    return str(soup)

def perform_microsoft_login(page):
    try:
        from config import CITSMART_EMAIL, PASSWORD
        email, password = CITSMART_EMAIL, PASSWORD
    except Exception as e:
        logging.warning(f"Não foi possível obter credenciais de config: {e}")
        return

    try:
        page.wait_for_selector('input[name="loginfmt"]', timeout=8000)
        logging.info("Preenchendo e-mail da conta Microsoft...")
        page.fill('input[name="loginfmt"]', email)
        page.click('input[id="idSIButton9"]')
        time.sleep(1.5)
    except Exception:
        pass

    try:
        page.wait_for_selector('input[name="passwd"]', timeout=8000)
        logging.info("Preenchendo senha...")
        page.fill('input[name="passwd"]', password)
        page.click('input[id="idSIButton9"]')
        time.sleep(1.5)
    except Exception:
        pass

    try:
        page.wait_for_selector('#KmsiCheckboxField, input[id="idSIButton9"]', timeout=8000)
        logging.info("Pulando confirmação KMSI...")
        page.click('input[id="idSIButton9"]')
        time.sleep(2)
    except Exception:
        pass

def scrape_all_faqs():
    try:
        from playwright.sync_api import sync_playwright
    except ImportError:
        logging.error("Biblioteca 'playwright' não encontrada. Execute 'pip install playwright' e 'playwright install chromium'.")
        return

    print_header("SCRAPER FAQ - BASE DE CONHECIMENTO", color=CYAN)
    logger.info("🤖 Iniciando raspagem de FAQs e Manuais do SharePoint...")
    init_faq_schema()
    conn = sqlite3.connect(db_path)
    cursor = conn.cursor()

    cursor.execute("""
        SELECT id, titulo, url 
        FROM faqs 
        WHERE conteudo IS NULL OR conteudo = ''
    """)
    pending_faqs = cursor.fetchall()
    
    if not pending_faqs:
        logger.info("🎉 Todos os FAQs já estão com conteúdo raspado no banco de dados!")
        conn.close()
        return

    logger.info(f"📄 Encontrados {len(pending_faqs)} FAQs pendentes de raspagem completa.")

    first_url = pending_faqs[0][2]

    with sync_playwright() as p:
        browser = p.chromium.launch(headless=True)
        context = browser.new_context()
        page = context.new_page()

        logger.info("🔑 Acessando portal para autenticação inicial...")
        page.goto(first_url, wait_until="domcontentloaded", timeout=40000)
        perform_microsoft_login(page)
        
        sucessos = 0
        erros = 0

        for faq_id, titulo, url in pending_faqs:
            try:
                logger.info(f"⏳ Raspando: '{titulo}' ({url})")
                page.goto(url, wait_until="networkidle", timeout=35000)
                
                if "login.microsoftonline.com" in page.url:
                    perform_microsoft_login(page)
                    page.goto(url, wait_until="networkidle", timeout=35000)

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
                    logger.info(f"✅ Salvo com sucesso: '{titulo}'")
                else:
                    erros += 1
                    logger.warning(f"⚠️ Não foi possível isolar o contêiner de texto em: '{titulo}'")
            except Exception as e:
                erros += 1
                logger.error(f"❌ Erro ao processar '{titulo}': {e}")

        browser.close()

    conn.close()
    logger.info(f"✨ Raspagem concluída! Sucessos: {sucessos} | Erros/Pendentes: {erros}")

if __name__ == "__main__":
    scrape_all_faqs()
