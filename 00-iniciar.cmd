@echo off
cd /d "%~dp0"
echo Ativando ambiente virtual (venv)...
call venv\Scripts\activate.bat
echo Iniciando o Streamlit na porta 8501...
streamlit run dashboard.py --server.port 8501
pause
