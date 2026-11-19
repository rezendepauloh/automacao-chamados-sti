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
O projeto conta com um inicializador automatizado que, mesmo em um computador recém-clonado:
1. Cria o `.env` automaticamente a partir do `.env.example`, caso ainda não exista.
2. Cria toda a estrutura de diretórios locais (`data/`, `models/`, `debug_logs/`, `uploads/`, `src/js/`).
3. Sobe os contêineres (`db` PostgreSQL, `evolution-api` e `web` Streamlit).
4. Inicializa automaticamente todas as tabelas, configurações iniciais, rotinas de cron e índices no PostgreSQL (`init.py` -> `initialize_database()`) e exibe um QR Code no terminal para acesso via rede:

```bash
chmod +x 00-iniciar.sh
./00-iniciar.sh
```

### No Windows:
Basta clicar duas vezes em (também cria `.env` e pastas automaticamente):
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

---

## 🚀 5. Migração Inicial e Deploy Contínuo (WSL → WSL Produção)

Para ambientes onde o desenvolvimento ocorre em uma máquina WSL e a produção temporária em outra máquina WSL dedicada:

1. Configure no `.env` de desenvolvimento as variáveis `DEPLOY_HOST`, `DEPLOY_USER`, `DEPLOY_PORT` (ex.: `2222`), `DEPLOY_DEST_DIR` e `DEPLOY_SSH_KEY`.
2. **Migração inicial completa** (código, `.env`, `.secret.key`, `python_keyring`, `data/`, `models/`, `uploads/` e volumes `postgres_data`, `evolution_instances`, `evolution_store`):
   ```bash
   ./util/deploy/migrar_para_producao.sh
   ```
3. **Deploy contínuo de atualizações de código** (preserva `.env`, `.secret.key`, banco e dados de produção):
   ```bash
   ./00-iniciar.sh --deploy
   # ou
   ./00-iniciar.sh -dp
   ```

---

## 🛠️ 6. Troubleshooting (WSL2 → WSL2 & Rede)

### 1. API REST (`8510`) retorna `Connection reset by peer` para clientes externos
- **Causa:** Por padrão, `BANCADA_API_BIND` no `.env` é `127.0.0.1` (loopback). Quando o tráfego chega de outra máquina via `netsh interface portproxy` do Windows para o IP interno do WSL2, o `wslrelay` derruba o encaminhamento para `127.0.0.1`.
- **Solução:**
  1. No `.env` da máquina de produção, defina `BANCADA_API_BIND="0.0.0.0"` (mantendo `127.0.0.1` como padrão seguro no `docker-compose.yml`) e recrie o container: `docker compose up -d api`.
  2. No PowerShell (Admin) do Windows de produção, adicione o `portproxy` e a regra de Firewall para a porta `8510` apontando para o IP interno atual do WSL2 (`hostname -I`).

### 2. SSH (`2222`) ou portas `8501`/`8510` param de responder após reinício
- **Causa:** O serviço `ssh` no WSL2 não inicia automaticamente após reboot ou o IP interno do WSL2 mudou.
- **Solução:** No WSL2 de produção, rode `sudo service ssh start`, confira o IP com `hostname -I | awk '{print $1}'` e atualize o `portproxy` no PowerShell (Admin) para as portas `2222`, `8501` e `8510`.

### 3. Erro `x509: certificate signed by unknown authority` no `docker compose`
- **Causa:** O daemon Docker em um WSL2 recém-instalado ainda não confia no certificado raiz da inspeção SSL corporativa.
- **Solução:** Copie o certificado `.crt` da CA para `/usr/local/share/ca-certificates/`, execute `sudo update-ca-certificates` e reinicie o serviço com `sudo service docker restart`.

