#!/bin/bash

# Direciona para a pasta onde o script está localizado
DIR="$( cd "$( dirname "${BASH_SOURCE[0]}" )" >/dev/null 2>&1 && pwd )"
cd "$DIR"

# Função de encerramento gracioso ao pressionar CTRL+C
cleanup() {
    echo ""
    echo "======================================================================="
    echo " Encerrando containers do Sistema Bancada..."
    echo "======================================================================="
    if command -v docker >/dev/null 2>&1; then
        docker compose down
    fi
    echo " [OK] Servidor e containers finalizados com sucesso."
    exit 0
}

# Captura sinal de interrupção (CTRL+C) e terminação
trap cleanup INT TERM

# Detectar IP Local IPv4 no Linux/WSL
LOCAL_IP=$(hostname -I 2>/dev/null | awk '{print $1}')

if [ -z "$LOCAL_IP" ]; then
    LOCAL_IP=$(ip route get 1.1.1.1 2>/dev/null | grep -oP 'src \K\S+')
fi

if [ -z "$LOCAL_IP" ]; then
    LOCAL_IP="localhost"
fi

export HOST_IP="${LOCAL_IP}"

# Garantir existência da pasta do keyring e do banco SQLite local
mkdir -p ~/.local/share/python_keyring
touch chamados.db

if command -v docker >/dev/null 2>&1; then
    # Verifica se o usuário solicitou reconstrução forçada (--build ou -b)
    FORCE_BUILD=false
    if [ "$1" == "--build" ] || [ "$1" == "-b" ]; then
        FORCE_BUILD=true
    fi

    # Verifica se a imagem Docker já foi construída
    IMAGE_ID=$(docker compose images -q web 2>/dev/null)
    if [ -z "$IMAGE_ID" ]; then
        IMAGE_ID=$(docker images -q automacao-chamados-sti-web:latest 2>/dev/null)
    fi

    if [ -z "$IMAGE_ID" ] || [ "$FORCE_BUILD" = true ]; then
        echo " [INFO] Construindo imagem Docker..."
        if ! docker compose build web; then
            echo ""
            echo " [ERRO] Falha ao construir a imagem Docker. Verifique os logs acima."
            exit 1
        fi
        echo " [OK] Imagem Docker construída com sucesso!"
    fi

    if ! docker compose up -d --force-recreate web; then
        echo ""
        echo " [ERRO] Falha ao iniciar o container Docker."
        exit 1
    fi

    # Executa o assistente de senhas no container interativo com terminal tty se solicitado ou inicial
    if [ "$1" == "--config-senhas" ] || [ "$1" == "-s" ]; then
        echo " [INFO] Abrindo configurador de senhas no container..."
        docker compose run --rm web python src/salvar_senha.py
    fi

    # Transmite logs em tempo real (exibindo o banner colorido e QR code do init.py)
    docker compose logs -f web
else
    echo " [ERRO] Docker não encontrado no PATH. Instale o Docker e o Docker Compose para executar a aplicação."
    exit 1
fi
