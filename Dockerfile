# 1. Utiliza a imagem base oficial do Red Hat UBI 9 com Python 3.12 solicitada pelo Victor
FROM registry.access.redhat.com/ubi9/python-312:9.8-1786510252

# 2. Define o diretório de trabalho dentro do container
WORKDIR /app

# 3. Copia os arquivos de dependências do projeto
COPY requirements.txt .

# 4. Instala as dependências Python no ambiente do container
RUN pip install --no-cache-dir -r requirements.txt

# 5. Copia todo o código-fonte restante para dentro do container
COPY . .

# 6. Expõe a porta padrão onde o Streamlit roda
EXPOSE 8501

# 7. Configura variáveis de ambiente para o Streamlit rodar em modo container limpo
ENV STREAMLIT_SERVER_PORT=8501
ENV STREAMLIT_SERVER_ADDRESS=0.0.0.0
ENV STREAMLIT_SERVER_HEADLESS=true

# 8. Comando padrão para iniciar a aplicação Streamlit
CMD ["streamlit", "run", "dashboard.py"]