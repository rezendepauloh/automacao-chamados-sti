# ==============================================================================
# Dockerfile - Sistema Bancada (Ambiente WSL / Linux Padrão)
# ==============================================================================

FROM python:3.11-slim

# Variáveis de ambiente para comportamento limpo do Python e bypass de inspeção SSL corporativa
ENV PYTHONDONTWRITEBYTECODE=1 \
    PYTHONUNBUFFERED=1 \
    DEBIAN_FRONTEND=noninteractive \
    TZ=America/Campo_Grande \
    PIP_DEFAULT_TIMEOUT=120 \
    PIP_TRUSTED_HOST="pypi.org pypi.python.org files.pythonhosted.org github.com github-releases.githubusercontent.com objects.githubusercontent.com release-assets.githubusercontent.com raw.githubusercontent.com github-production-release-asset-2e65be.s3.amazonaws.com github-cloud.s3.amazonaws.com" \
    NODE_TLS_REJECT_UNAUTHORIZED=0

# Diretório base de trabalho da aplicação
WORKDIR /app

# Instalação de dependências do sistema operacional e certificados
RUN apt-get update && apt-get install -y --no-install-recommends \
    build-essential \
    curl \
    ca-certificates \
    tzdata \
    libpq-dev \
    && update-ca-certificates \
    && ln -snf /usr/share/zoneinfo/$TZ /etc/localtime && echo $TZ > /etc/timezone \
    && rm -rf /var/lib/apt/lists/*

# Configura arquivo global do pip para confiar nos domínios corporativos e CDNs
RUN mkdir -p /etc /root/.pip \
    && printf "[global]\ntrusted-host =\n    pypi.org\n    pypi.python.org\n    files.pythonhosted.org\n    github.com\n    github-releases.githubusercontent.com\n    objects.githubusercontent.com\n    release-assets.githubusercontent.com\n    raw.githubusercontent.com\n    github-production-release-asset-2e65be.s3.amazonaws.com\n    github-cloud.s3.amazonaws.com\n" > /etc/pip.conf \
    && cp /etc/pip.conf /root/.pip/pip.conf

# Copia arquivo de dependências
COPY requirements.txt .

# Atualiza pip e instala dependências Python com bypass de SSL corporativo
RUN pip install --no-cache-dir \
    --trusted-host pypi.org \
    --trusted-host pypi.python.org \
    --trusted-host files.pythonhosted.org \
    --trusted-host github.com \
    --trusted-host github-releases.githubusercontent.com \
    --trusted-host objects.githubusercontent.com \
    --trusted-host release-assets.githubusercontent.com \
    --trusted-host raw.githubusercontent.com \
    --trusted-host github-production-release-asset-2e65be.s3.amazonaws.com \
    --trusted-host github-cloud.s3.amazonaws.com \
    --upgrade pip \
    && pip install --no-cache-dir \
    --trusted-host pypi.org \
    --trusted-host pypi.python.org \
    --trusted-host files.pythonhosted.org \
    --trusted-host github.com \
    --trusted-host github-releases.githubusercontent.com \
    --trusted-host objects.githubusercontent.com \
    --trusted-host release-assets.githubusercontent.com \
    --trusted-host raw.githubusercontent.com \
    --trusted-host github-production-release-asset-2e65be.s3.amazonaws.com \
    --trusted-host github-cloud.s3.amazonaws.com \
    -r requirements.txt

# Instala dependências do sistema e navegador Chromium para o Playwright (faq_scraper)
RUN playwright install-deps chromium \
    && playwright install chromium

# Copia todo o código-fonte para o container
COPY . .

# Expõe a porta padrão do Streamlit
EXPOSE 8501

# Variáveis padrão de execução do Streamlit
ENV STREAMLIT_SERVER_PORT=8501 \
    STREAMLIT_SERVER_ADDRESS=0.0.0.0 \
    STREAMLIT_SERVER_HEADLESS=true

# Checagem de integridade do serviço Streamlit
HEALTHCHECK --interval=10s --timeout=5s --start-period=15s --retries=3 \
    CMD curl --fail http://localhost:8501/_stcore/health || exit 1

# Comando padrão de inicialização (executa o banner interativo e sobe o Streamlit)
CMD ["python", "init.py"]
