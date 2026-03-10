# 1. Garante que o terminal está rodando na mesma pasta deste script (raiz do projeto)
Set-Location -Path $PSScriptRoot

# 2. Ativa o ambiente virtual (assumindo que sua pasta se chama 'venv')
if (Test-Path ".\venv\Scripts\Activate.ps1") {
    .\venv\Scripts\Activate.ps1
}

# 3. Executa a esteira de dados na ordem correta
Write-Host "Iniciando OTRS..."
python otrs_scraper.py

Write-Host "Iniciando CitSmart..."
python citsmart_scraper.py

Write-Host "Iniciando Pré-processamento..."
python preprocess_chamados.py

Write-Host "Iniciando Classificador de Tags..."
python tag_classifier.py

Write-Host "Pipeline finalizado com sucesso!"