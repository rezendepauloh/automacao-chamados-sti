# Guia de Instalação e Deploy via Docker Compose

> **Ambiente:** Linux / WSL2 (Ubuntu 22.04+) / Windows 10/11  
> **Orquestrador:** `docker-compose.yml`

---

## 🎯 1. Pré-Requisitos

- **Docker Desktop** (no Windows) ou **Docker Engine + Docker Compose v2** (no Linux nativo/WSL).
- Portas `8501` (Streamlit) e `5432` (PostgreSQL) liberadas no host.

---

## 🚀 2. Inicialização Rápida

### No Linux ou WSL (Ubuntu):
O projeto conta com um inicializador automatizado que valida as permissões, sobe os contêineres e exibe um QR Code no terminal para acesso via smartphone ou tablet na mesma rede:

```bash
chmod +x 00-iniciar.sh
./00-iniciar.sh
```

### No Windows:
Basta clicar duas vezes em:
```cmd
00-iniciar.cmd
```

---

## 🐳 3. Comandos Docker do Dia-a-Dia

Todos os comandos devem ser executados a partir da raiz do repositório:

### Subir os contêineres em segundo plano:
```bash
docker compose up -d
```

### Acompanhar logs da aplicação em tempo real:
```bash
docker compose logs -f app
```

### Acompanhar logs do banco de dados:
```bash
docker compose logs -f db
```

### Parar os contêineres:
```bash
docker compose down
```

### Reiniciar a aplicação após alterações de dependências:
```bash
docker compose restart app
```

---

## 🌐 4. Acessando a Aplicação

Abra o navegador em:
- **Painel Principal:** [http://localhost:8501](http://localhost:8501)
- **Acesso na Rede Local:** `http://<IP-DO-SEU-COMPUTADOR>:8501`

