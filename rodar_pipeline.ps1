# 1. Garante que o terminal está rodando na mesma pasta deste script (raiz do projeto)
Set-Location -Path $PSScriptRoot

# 2. Ativa o ambiente virtual (assumindo que sua pasta se chama 'venv')
if (Test-Path ".\venv\Scripts\Activate.ps1") {
    .\venv\Scripts\Activate.ps1
}

# 3. Executa a esteira de dados na ordem correta
Write-Host "[1/4] Iniciando OTRS..." -ForegroundColor Green -BackgroundColor Black
python otrs_scraper.py

Write-Host "[2/4] Iniciando CitSmart..." -ForegroundColor Green -BackgroundColor Black
python citsmart_scraper.py

Write-Host "[3/4] Iniciando Pré-processamento..." -ForegroundColor Green -BackgroundColor Black
python preprocess_chamados.py

Write-Host "[4/4] Iniciando Classificador de Tags..." -ForegroundColor Green -BackgroundColor Black
python tag_classifier.py

Write-Host "Pipeline finalizado com sucesso!" -ForegroundColor Green -BackgroundColor Black