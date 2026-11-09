import sys
import time
from pathlib import Path

# Paths
root_dir = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(root_dir))
sys.path.insert(0, str(root_dir / "src"))

import pandas as pd
from src.scrapers.otrs_scraper_fast import OTRSClient, HEADERS as FAST_HEADERS
from src.terminal import print_header, CYAN, GREEN, YELLOW, RED, WHITE, log

def benchmark():
    print_header("TESTE COMPARATIVO: SCRAPER SELENIUM vs SCRAPER HTTP RÁPIDO", color=CYAN)
    
    # ==========================================
    # 1. TESTE DO SCRAPER RÁPIDO (HTTP Direto)
    # ==========================================
    print(f"\n{YELLOW}▶ Iniciando Scraper HTTP Rápido (1 página, 5 chamados com detalhes)...{WHITE}")
    t0_fast = time.time()
    
    client = OTRSClient()
    client.login()
    
    # Baixa 1 página
    all_fast_items = client.fetch_queue_tickets(max_pages=1, max_workers=5)
    t_fast_total = time.time() - t0_fast
    
    print(f"{GREEN}✔ Scraper HTTP concluído em: {t_fast_total:.2f} segundos!{WHITE}")
    print(f"  Total de chamados extraídos: {len(all_fast_items)}")
    
    # ==========================================
    # 2. TESTE DO SCRAPER SELENIUM (Clássico)
    # ==========================================
    print(f"\n{YELLOW}▶ Iniciando Scraper Selenium Tradicional (apenas 5 chamados para comparação)...{WHITE}")
    t0_sel = time.time()
    
    try:
        from config import get_chrome_driver, OTRS_URL
        from src.scrapers.otrs_scraper import login_page, navigation_queue, extract_row_data
        from selenium.webdriver.common.by import By
        from selenium.webdriver.support.ui import WebDriverWait
        from selenium.webdriver.support import expected_conditions as EC
        
        driver = get_chrome_driver()
        login_page(driver)
        navigation_queue(driver)
        
        WebDriverWait(driver, 15).until(EC.presence_of_element_located((By.CSS_SELECTOR, "table.TableSmall, table#OverviewControl")))
        rows = driver.find_elements(By.CSS_SELECTOR, "tr.MasterAction")
        
        # Pega apenas os primeiros 5 para não demorar demais
        sel_items = []
        limit_sel = min(5, len(rows))
        for idx in range(limit_sel):
            t_row_start = time.time()
            rows_now = driver.find_elements(By.CSS_SELECTOR, "tr.MasterAction")
            data = extract_row_data(driver, rows_now[idx])
            sel_items.append(data)
            t_row = time.time() - t_row_start
            print(f"    [Selenium] Chamado {idx+1}/{limit_sel} ({data.get('Chamado#')}) levou {t_row:.2f}s")
            
        driver.quit()
        t_sel_total = time.time() - t0_sel
        print(f"{GREEN}✔ Scraper Selenium (5 chamados) concluído em: {t_sel_total:.2f} segundos!{WHITE}")
    except Exception as e:
        print(f"{RED}Erro no teste do Selenium: {e}{WHITE}")
        t_sel_total = None
        sel_items = []

    # ==========================================
    # 3. TABELA COMPARATIVA
    # ==========================================
    print_header("RESUMO E RESULTADOS COMPARATIVOS", color=GREEN)
    
    speedup = f"{t_sel_total / (t_fast_total / len(all_fast_items) * 5):.1f}x mais rápido" if t_sel_total else "N/A"
    
    print(f"⚡ Tempo HTTP Rápido (27 chamados inteiros):  {t_fast_total:.2f} s  (~{t_fast_total/len(all_fast_items):.2f}s por chamado)")
    if t_sel_total:
        print(f"🐢 Tempo Selenium (apenas 5 chamados):        {t_sel_total:.2f} s  (~{t_sel_total/len(sel_items):.2f}s por chamado)")
        estimativa_selenium_27 = (t_sel_total / len(sel_items)) * len(all_fast_items)
        print(f"⏳ Estimativa do Selenium para 27 chamados:   {estimativa_selenium_27:.2f} s ({estimativa_selenium_27/60:.1f} minutos)")
        print(f"🚀 Ganho de Performance Real:                 {speedup}!")

    # Validação de integridade do primeiro chamado
    if all_fast_items and sel_items:
        f0 = all_fast_items[0]
        s0 = sel_items[0]
        print("\n🔍 Validação de Fidelidade dos Dados (Chamado #1):")
        print(f"  • Chamado#:    [HTTP] {f0.get('Chamado#')} | [Selenium] {s0.get('Chamado#')}")
        print(f"  • Solicitante: [HTTP] {f0.get('Nome do Usuário')} | [Selenium] {s0.get('Nome do Usuário')}")
        print(f"  • Título:      [HTTP] {f0.get('Título')[:35]}... | [Selenium] {s0.get('Título')[:35]}...")
        desc_match = "IGUAL" if f0.get('Descrição', '').strip()[:50] == s0.get('Descrição', '').strip()[:50] else "DIFERENTE (verificar)"
        print(f"  • Descrição:   [{desc_match}]")

if __name__ == "__main__":
    benchmark()

