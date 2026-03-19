import subprocess
import os

# Caminho para o executável "invisível" do Python dentro do seu ambiente virtual
python_exe = os.path.join("venv", "Scripts", "pythonw.exe")

# Lista dos seus robôs na ordem correta
scripts = [
    "otrs_scraper.py",
    "citsmart_scraper.py",
    "preprocess_chamados.py",
    "tag_classifier.py",
    "sync_master.py"
]

# CREATE_NO_WINDOW = 0x08000000 é um comando da API do Windows que 
# proíbe absolutamente a criação de qualquer janela de terminal
CREATE_NO_WINDOW = 0x08000000

for script in scripts:
    # Executa cada script na sequência, totalmente oculto
    subprocess.run([python_exe, script], creationflags=CREATE_NO_WINDOW)