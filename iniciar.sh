#!/bin/bash

# Direciona para a pasta onde o script está localizado
DIR="$( cd "$( dirname "${BASH_SOURCE[0]}" )" >/dev/null 2>&1 && pwd )"
cd "$DIR"

# Ativação do ambiente virtual
if [ -d "venv" ]; then
    source venv/bin/activate
elif [ -d ".venv" ]; then
    source .venv/bin/activate
else
    echo "[ERRO] Ambiente virtual (venv) não encontrado!"
    exit 1
fi

clear
echo "======================================================================="
echo "                       S I S T E M A   B A N C A D A"
echo "======================================================================="
echo ""

# Detectar IP Local IPv4 no Linux/WSL
LOCAL_IP=$(hostname -I 2>/dev/null | awk '{print $1}')

if [ -z "$LOCAL_IP" ]; then
    LOCAL_IP=$(ip route get 1.1.1.1 2>/dev/null | grep -oP 'src \K\S+')
fi

if [ -z "$LOCAL_IP" ]; then
    LOCAL_IP="localhost"
fi

PORT="8501"
FULL_URL="http://${LOCAL_IP}:${PORT}/"

mkdir -p src/js
echo "window.BANCADA_LOCAL_IP = '${LOCAL_IP}';" > src/js/server-info.js

echo " [OK] Banco de dados SQLite centralizado ativo em chamados.db"
echo ""
echo " -----------------------------------------------------------------------"
echo "  ACESSO NO COMPUTADOR : http://localhost:${PORT}/"
echo "  ACESSO NO CELULAR / TABLET: ${FULL_URL}"
echo " -----------------------------------------------------------------------"
echo ""
echo " Aponte a câmera do Celular para o QR Code abaixo:"
echo ""

curl -s "https://qrenco.de/${FULL_URL}" 2>/dev/null

echo ""
echo "======================================================================="
echo " Pressione CTRL+C para encerrar o servidor."
echo "======================================================================="
echo ""

python -m streamlit run dashboard.py --server.port "${PORT}" --server.headless=true --logger.level=error
