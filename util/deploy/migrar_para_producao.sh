#!/usr/bin/env bash
# ==============================================================================
# MIGRAÇÃO INICIAL COMPLETA — DEV (WSL) -> PRODUÇÃO (WSL)
# ==============================================================================
# Copia código, segredos (.env, .secret.key, python_keyring), diretórios de
# dados (data/, models/, uploads/) e volumes Docker (postgres_data,
# evolution_instances, evolution_store) para a máquina de produção e sobe a stack.
# ==============================================================================

set -u

cd "$(dirname "$0")/../.." || exit 1

C_RESET="\033[0m"
C_BOLD="\033[1m"
C_CYAN="\033[36m"
C_GREEN="\033[32m"
C_YELLOW="\033[33m"
C_RED="\033[31m"

get_env_var() {
    local key="$1"
    local default_val="${2:-}"
    local val="${!key:-}"
    if [ -z "$val" ] && [ -f .env ]; then
        val=$(grep -E "^[[:space:]]*${key}[[:space:]]*=" .env | tail -n 1 | cut -d '=' -f 2- | tr -d ' "\r\n' | tr -d "'")
    fi
    echo "${val:-$default_val}"
}

HOST=$(get_env_var "DEPLOY_HOST")
USER=$(get_env_var "DEPLOY_USER")
PORT_SSH=$(get_env_var "DEPLOY_PORT" "2222")
DEST_DIR=$(get_env_var "DEPLOY_DEST_DIR" "~/PythonProjects/automated-OTRS-and-CitSmart")
SSH_KEY=$(get_env_var "DEPLOY_SSH_KEY" "~/.ssh/id_ed25519")
SSH_KEY="${SSH_KEY/#\~/$HOME}"

if [ -z "$HOST" ] || [ -z "$USER" ]; then
    echo -e "${C_RED}[ERRO] Defina DEPLOY_HOST e DEPLOY_USER no arquivo .env antes de executar a migração.${C_RESET}"
    exit 1
fi

SSH_OPTS="-p ${PORT_SSH} -o StrictHostKeyChecking=accept-new -o ConnectTimeout=10"
SCP_OPTS="-P ${PORT_SSH} -o StrictHostKeyChecking=accept-new -o ConnectTimeout=10"
if [ -f "$SSH_KEY" ]; then
    SSH_OPTS="$SSH_OPTS -i $SSH_KEY"
    SCP_OPTS="$SCP_OPTS -i $SSH_KEY"
fi

echo -e "${C_CYAN}${C_BOLD}================================================================${C_RESET}"
echo -e "${C_CYAN}${C_BOLD}   MIGRAÇÃO INICIAL COMPLETA — BANCADA STI (${USER}@${HOST}:${PORT_SSH})${C_RESET}"
echo -e "${C_CYAN}${C_BOLD}================================================================${C_RESET}"
echo ""

# 1. Pré-checagem SSH e Docker remoto
echo -e "${C_YELLOW}[1/7] Validando acesso SSH e Docker no destino...${C_RESET}"
if ! ssh $SSH_OPTS "${USER}@${HOST}" "command -v docker >/dev/null 2>&1 && docker compose version >/dev/null 2>&1"; then
    echo -e "${C_RED}[ERRO] Falha no SSH ou Docker/Docker Compose não instalado em ${USER}@${HOST}.${C_RESET}"
    exit 1
fi
echo -e "${C_GREEN}[OK] SSH e Docker Compose disponíveis no destino.${C_RESET}"

# Trava de segurança caso destino já possua stack
REMOTE_EXISTS=$(ssh $SSH_OPTS "${USER}@${HOST}" "[ -f ${DEST_DIR}/.env ] && echo 'yes' || echo 'no'")
if [ "$REMOTE_EXISTS" = "yes" ]; then
    echo -e "${C_YELLOW}[AVISO] O diretório ${DEST_DIR} já possui .env na máquina de destino.${C_RESET}"
    read -p "Deseja sobrescrever dados e volumes em ${HOST}? (digite SIM para continuar): " confirm
    if [ "$confirm" != "SIM" ]; then
        echo -e "${C_RED}Migração cancelada pelo usuário.${C_RESET}"
        exit 0
    fi
fi

# 2. Preparar diretórios remotos
echo -e "${C_YELLOW}[2/7] Preparando estrutura de diretórios no destino...${C_RESET}"
ssh $SSH_OPTS "${USER}@${HOST}" "mkdir -p ${DEST_DIR}/data ${DEST_DIR}/models ${DEST_DIR}/uploads ${DEST_DIR}/debug_logs ${DEST_DIR}/src/js ~/.local/share/python_keyring" || exit 1

# 3. Sincronizar código + dados iniciais (dump completo, sem --delete destrutivo)
echo -e "${C_YELLOW}[3/7] Transferindo código, data/, models/ e uploads/ via rsync...${C_RESET}"
rsync -avz --progress \
    -e "ssh $SSH_OPTS" \
    --exclude '.git' \
    --exclude '.venv' \
    --exclude '__pycache__' \
    --exclude '.pytest_cache' \
    --exclude '*.pyc' \
    --exclude 'temp' \
    --exclude 'src/js/server-info.js' \
    ./ "${USER}@${HOST}:${DEST_DIR}/" || exit 1

# 4. Transferir segredos (.env, .secret.key, python_keyring)
echo -e "${C_YELLOW}[4/7] Transferindo .env, .secret.key e cofre python_keyring...${C_RESET}"
if [ -f ".env" ]; then
    scp $SCP_OPTS ".env" "${USER}@${HOST}:${DEST_DIR}/.env" || exit 1
    ssh $SSH_OPTS "${USER}@${HOST}" "chmod 600 ${DEST_DIR}/.env"
fi
if [ -f ".secret.key" ]; then
    scp $SCP_OPTS ".secret.key" "${USER}@${HOST}:${DEST_DIR}/.secret.key" || exit 1
    ssh $SSH_OPTS "${USER}@${HOST}" "chmod 600 ${DEST_DIR}/.secret.key"
fi
if [ -d "$HOME/.local/share/python_keyring" ]; then
    tar -C "$HOME/.local/share" -czf - python_keyring 2>/dev/null | \
        ssh $SSH_OPTS "${USER}@${HOST}" "tar -C ~/.local/share -xzf - && chmod 700 ~/.local/share/python_keyring && chmod 600 ~/.local/share/python_keyring/* 2>/dev/null || true"
fi

# 5. Exportar e migrar PostgreSQL (dump lógico)
echo -e "${C_YELLOW}[5/7] Exportando e migrando banco PostgreSQL...${C_RESET}"
PG_USER=$(get_env_var "POSTGRES_USER")
PG_DB=$(get_env_var "POSTGRES_DB")
DUMP_FILE="/tmp/bancada_pg_migration_$$.dump"

if docker ps --format '{{.Names}}' | grep -q '^bancada_postgres_db$'; then
    docker exec bancada_postgres_db pg_dump -U "$PG_USER" -Fc "$PG_DB" > "$DUMP_FILE" || exit 1
    scp $SCP_OPTS "$DUMP_FILE" "${USER}@${HOST}:/tmp/bancada_pg_migration.dump" || exit 1
    rm -f "$DUMP_FILE"

    ssh $SSH_OPTS "${USER}@${HOST}" "cd ${DEST_DIR} && docker compose up -d db" || exit 1
    echo "Aguardando PostgreSQL ficar saudável no destino..."
    ssh $SSH_OPTS "${USER}@${HOST}" "for i in \$(seq 1 30); do docker exec bancada_postgres_db pg_isready -U '${PG_USER}' -d '${PG_DB}' >/dev/null 2>&1 && break; sleep 2; done"
    ssh $SSH_OPTS "${USER}@${HOST}" "docker exec -i bancada_postgres_db pg_restore -U '${PG_USER}' -d '${PG_DB}' --clean --if-exists --no-owner --no-privileges < /tmp/bancada_pg_migration.dump; rm -f /tmp/bancada_pg_migration.dump"
else
    echo -e "${C_YELLOW}[AVISO] Container bancada_postgres_db não está rodando localmente; pulando dump do PostgreSQL.${C_RESET}"
fi

# 6. Migrar volumes da Evolution API (evolution_instances e evolution_store)
echo -e "${C_YELLOW}[6/7] Migrando volumes da Evolution API...${C_RESET}"
PROJECT_NAME=$(basename "$(pwd)" | tr '[:upper:]' '[:lower:]' | sed 's/[^a-z0-9_-]//g')
REMOTE_PROJECT=$(ssh $SSH_OPTS "${USER}@${HOST}" "cd ${DEST_DIR} && basename \"\$(pwd)\" | tr '[:upper:]' '[:lower:]' | sed 's/[^a-z0-9_-]//g'")

for VOL_SUFFIX in evolution_instances evolution_store; do
    SRC_VOL="${PROJECT_NAME}_${VOL_SUFFIX}"
    DST_VOL="${REMOTE_PROJECT}_${VOL_SUFFIX}"
    if docker volume inspect "$SRC_VOL" >/dev/null 2>&1; then
        echo "  -> Migrando volume ${SRC_VOL} -> ${DST_VOL}..."
        ssh $SSH_OPTS "${USER}@${HOST}" "docker volume create ${DST_VOL} >/dev/null"
        docker run --rm -v "${SRC_VOL}:/from:ro" alpine tar -C /from -czf - . | \
            ssh $SSH_OPTS "${USER}@${HOST}" "docker run --rm -i -v '${DST_VOL}:/to' alpine tar -C /to -xzf -"
    fi
done

# 7. Subir toda a stack Docker no destino
echo -e "${C_YELLOW}[7/7] Construindo e iniciando toda a stack Docker na produção...${C_RESET}"
ssh -t $SSH_OPTS "${USER}@${HOST}" "cd ${DEST_DIR} && docker compose up -d --build && docker compose ps" || exit 1

echo ""
echo -e "${C_GREEN}${C_BOLD}🎉 MIGRAÇÃO INICIAL CONCLUÍDA COM SUCESSO!${C_RESET}"
echo -e "Lembre-se de rodar no PowerShell (Admin) da máquina ${HOST} o script:"
echo -e "  ${C_CYAN}util/rede_firewall/configurar_portproxy_firewall.ps1${C_RESET} (para expor a porta 8501 na rede)"
