import os
import pandas as pd
from datetime import datetime
from src.database.connection import get_connection, DB_TYPE

DEFAULT_TASKS = [
    {
        "task_id": "whatsapp_d1",
        "nome": "📱 Alertas WhatsApp (Plantões, Viagens e Portarias)",
        "categoria": "WhatsApp & Notificações",
        "ativo": 1,
        "tipo_agendamento": "horario_fixo",
        "intervalo_valor": 1,
        "intervalo_unidade": "dias",
        "horario_fixo": "12:00",
        "apenas_dias_uteis": 1,
        "descricao": "Dispara alertas D-1 de plantões e viagens, e portarias do mesmo dia às 12:00 em dias úteis."
    },
    {
        "task_id": "sync_portarias",
        "nome": "📜 Varredura de Novas Portarias MPMS",
        "categoria": "Diário Oficial",
        "ativo": 1,
        "tipo_agendamento": "intervalo",
        "intervalo_valor": 2,
        "intervalo_unidade": "horas",
        "horario_fixo": "12:00",
        "apenas_dias_uteis": 0,
        "descricao": "Consulta a API pública do MPMS em busca de publicações envolvendo servidores da bancada."
    },
    {
        "task_id": "sync_viagens",
        "nome": "🚗 Sincronização de Viagens (SharePoint)",
        "categoria": "Planilhas",
        "ativo": 1,
        "tipo_agendamento": "intervalo",
        "intervalo_valor": 4,
        "intervalo_unidade": "horas",
        "horario_fixo": "08:00",
        "apenas_dias_uteis": 0,
        "descricao": "Baixa a planilha oficial de viagens da bancada do SharePoint e atualiza o calendário e lista."
    },
    {
        "task_id": "sync_plantoes",
        "nome": "📅 Escalas de Plantão Matutino e Semanal",
        "categoria": "Planilhas",
        "ativo": 1,
        "tipo_agendamento": "intervalo",
        "intervalo_valor": 6,
        "intervalo_unidade": "horas",
        "horario_fixo": "08:00",
        "apenas_dias_uteis": 0,
        "descricao": "Verifica escalas de plantão e gera alertas de notificação no painel do sistema."
    },
    {
        "task_id": "sync_fiscalizacao",
        "nome": "📑 Fiscalização de Contratos & Garantias",
        "categoria": "Planilhas",
        "ativo": 1,
        "tipo_agendamento": "intervalo",
        "intervalo_valor": 12,
        "intervalo_unidade": "horas",
        "horario_fixo": "07:30",
        "apenas_dias_uteis": 0,
        "descricao": "Sincroniza as planilhas corporativas de fiscalização de contratos e controle de garantias."
    },
    {
        "task_id": "orquestrador_chamados",
        "nome": "🎫 Coleta & Classificação de Chamados TI",
        "categoria": "Chamados TI",
        "ativo": 0,
        "tipo_agendamento": "intervalo",
        "intervalo_valor": 30,
        "intervalo_unidade": "minutos",
        "horario_fixo": "08:00",
        "apenas_dias_uteis": 1,
        "descricao": "Executa o robô de extração de chamados OTRS/CitSmart e classificação de tags por IA."
    }
]

def setup_cron_tables():
    """Cria tabelas de agendamentos (cron_schedules) e logs de execução (cron_logs)."""
    conn = get_connection()
    cursor = conn.cursor()

    if DB_TYPE in ["postgres", "postgresql"]:
        cursor.execute("""
        CREATE TABLE IF NOT EXISTS cron_schedules (
            task_id VARCHAR(80) PRIMARY KEY,
            nome VARCHAR(150) NOT NULL,
            categoria VARCHAR(50),
            ativo INTEGER DEFAULT 1,
            tipo_agendamento VARCHAR(30) DEFAULT 'intervalo',
            intervalo_valor INTEGER DEFAULT 2,
            intervalo_unidade VARCHAR(20) DEFAULT 'horas',
            horario_fixo VARCHAR(10) DEFAULT '12:00',
            apenas_dias_uteis INTEGER DEFAULT 0,
            descricao TEXT,
            ultima_execucao TIMESTAMP,
            proxima_execucao TIMESTAMP,
            ultimo_status VARCHAR(30) DEFAULT 'pendente',
            ultimo_log TEXT,
            updated_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP
        );
        """)
        cursor.execute("""
        CREATE TABLE IF NOT EXISTS cron_logs (
            id SERIAL PRIMARY KEY,
            task_id VARCHAR(80) NOT NULL,
            inicio TIMESTAMP NOT NULL,
            fim TIMESTAMP,
            duracao_segundos NUMERIC(10, 2),
            status VARCHAR(30) NOT NULL,
            mensagem TEXT,
            created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP
        );
        """)
    else:
        cursor.execute("""
        CREATE TABLE IF NOT EXISTS cron_schedules (
            task_id TEXT PRIMARY KEY,
            nome TEXT NOT NULL,
            categoria TEXT,
            ativo INTEGER DEFAULT 1,
            tipo_agendamento TEXT DEFAULT 'intervalo',
            intervalo_valor INTEGER DEFAULT 2,
            intervalo_unidade TEXT DEFAULT 'horas',
            horario_fixo TEXT DEFAULT '12:00',
            apenas_dias_uteis INTEGER DEFAULT 0,
            descricao TEXT,
            ultima_execucao DATETIME,
            proxima_execucao DATETIME,
            ultimo_status TEXT DEFAULT 'pendente',
            ultimo_log TEXT,
            updated_at DATETIME DEFAULT CURRENT_TIMESTAMP
        );
        """)
        cursor.execute("""
        CREATE TABLE IF NOT EXISTS cron_logs (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            task_id TEXT NOT NULL,
            inicio DATETIME NOT NULL,
            fim DATETIME,
            duracao_segundos REAL,
            status TEXT NOT NULL,
            mensagem TEXT,
            created_at DATETIME DEFAULT CURRENT_TIMESTAMP
        );
        """)

    conn.commit()

    # Previne que tarefas fiquem presas em 'executando' caso o container tenha reiniciado ou ocorrido queda
    try:
        cursor.execute("UPDATE cron_schedules SET ultimo_status = 'erro', ultimo_log = 'Execução interrompida por reinício do sistema.' WHERE ultimo_status = 'executando';")
        cursor.execute("UPDATE cron_logs SET status = 'erro', mensagem = 'Interrompido por reinício do sistema' WHERE status = 'executando';")
        conn.commit()
    except Exception:
        pass

    cursor.close()
    conn.close()

    seed_cron_tasks_if_empty()

def seed_cron_tasks_if_empty():
    """Insere as rotinas padrão caso não existam."""
    conn = get_connection()
    cursor = conn.cursor()

    for task in DEFAULT_TASKS:
        try:
            if DB_TYPE in ["postgres", "postgresql"]:
                cursor.execute("""
                INSERT INTO cron_schedules (
                    task_id, nome, categoria, ativo, tipo_agendamento,
                    intervalo_valor, intervalo_unidade, horario_fixo,
                    apenas_dias_uteis, descricao
                ) VALUES (%s, %s, %s, %s, %s, %s, %s, %s, %s, %s)
                ON CONFLICT (task_id) DO NOTHING;
                """, (
                    task["task_id"], task["nome"], task["categoria"], task["ativo"],
                    task["tipo_agendamento"], task["intervalo_valor"], task["intervalo_unidade"],
                    task["horario_fixo"], task["apenas_dias_uteis"], task["descricao"]
                ))
            else:
                cursor.execute("""
                INSERT OR IGNORE INTO cron_schedules (
                    task_id, nome, categoria, ativo, tipo_agendamento,
                    intervalo_valor, intervalo_unidade, horario_fixo,
                    apenas_dias_uteis, descricao
                ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?);
                """, (
                    task["task_id"], task["nome"], task["categoria"], task["ativo"],
                    task["tipo_agendamento"], task["intervalo_valor"], task["intervalo_unidade"],
                    task["horario_fixo"], task["apenas_dias_uteis"], task["descricao"]
                ))
        except Exception:
            pass

    conn.commit()
    cursor.close()
    conn.close()

def get_cron_schedules() -> pd.DataFrame:
    """Retorna todas as rotinas agendadas do sistema."""
    setup_cron_tables()
    conn = get_connection()
    df = pd.read_sql_query("SELECT * FROM cron_schedules ORDER BY categoria, nome ASC", conn)
    conn.close()
    return df

def get_cron_schedule_by_id(task_id: str) -> dict | None:
    """Retorna as configurações de uma rotina específica."""
    setup_cron_tables()
    conn = get_connection()
    cursor = conn.cursor()
    if DB_TYPE in ["postgres", "postgresql"]:
        cursor.execute("SELECT * FROM cron_schedules WHERE task_id = %s", (task_id,))
    else:
        cursor.execute("SELECT * FROM cron_schedules WHERE task_id = ?", (task_id,))
    row = cursor.fetchone()
    if not row:
        cursor.close()
        conn.close()
        return None

    columns = [desc[0] for desc in cursor.description]
    cursor.close()
    conn.close()
    return dict(zip(columns, row))

def update_cron_schedule(task_id: str, ativo: bool, tipo_agendamento: str,
                         intervalo_valor: int, intervalo_unidade: str,
                         horario_fixo: str, apenas_dias_uteis: bool) -> bool:
    """Atualiza a configuração de uma rotina no banco de dados."""
    setup_cron_tables()
    conn = get_connection()
    cursor = conn.cursor()
    try:
        sql = """
        UPDATE cron_schedules SET
            ativo = %s,
            tipo_agendamento = %s,
            intervalo_valor = %s,
            intervalo_unidade = %s,
            horario_fixo = %s,
            apenas_dias_uteis = %s,
            updated_at = CURRENT_TIMESTAMP
        WHERE task_id = %s
        """
        params = (
            1 if ativo else 0,
            tipo_agendamento,
            int(intervalo_valor),
            intervalo_unidade,
            horario_fixo.strip(),
            1 if apenas_dias_uteis else 0,
            task_id
        )
        if DB_TYPE not in ["postgres", "postgresql"]:
            sql = sql.replace("%s", "?")

        cursor.execute(sql, params)
        conn.commit()
        return True
    except Exception as e:
        print(f"Erro ao atualizar cron schedule: {e}")
        return False
    finally:
        cursor.close()
        conn.close()

def log_cron_execution_start(task_id: str) -> int:
    """Registra o início de uma execução e atualiza o status na tabela principal."""
    setup_cron_tables()
    conn = get_connection()
    cursor = conn.cursor()
    now_str = datetime.now().strftime("%Y-%m-%d %H:%M:%S")

    log_id = 0
    try:
        if DB_TYPE in ["postgres", "postgresql"]:
            cursor.execute("""
            INSERT INTO cron_logs (task_id, inicio, status, mensagem)
            VALUES (%s, %s, 'executando', 'Execução iniciada pelo agendador...')
            RETURNING id;
            """, (task_id, now_str))
            log_id = cursor.fetchone()[0]
            cursor.execute("""
            UPDATE cron_schedules SET
                ultimo_status = 'executando',
                ultima_execucao = %s
            WHERE task_id = %s;
            """, (now_str, task_id))
        else:
            cursor.execute("""
            INSERT INTO cron_logs (task_id, inicio, status, mensagem)
            VALUES (?, ?, 'executando', 'Execução iniciada pelo agendador...');
            """, (task_id, now_str))
            log_id = cursor.lastrowid
            cursor.execute("""
            UPDATE cron_schedules SET
                ultimo_status = 'executando',
                ultima_execucao = ?
            WHERE task_id = ?;
            """, (now_str, task_id))
        conn.commit()
    except Exception as e:
        print(f"Erro ao iniciar log de cron: {e}")
    finally:
        cursor.close()
        conn.close()
    return log_id

def log_cron_execution_end(log_id: int, task_id: str, status: str, mensagem: str, duracao_seg: float = 0.0):
    """Registra a conclusão de uma execução."""
    setup_cron_tables()
    conn = get_connection()
    cursor = conn.cursor()
    now_str = datetime.now().strftime("%Y-%m-%d %H:%M:%S")

    try:
        if DB_TYPE in ["postgres", "postgresql"]:
            cursor.execute("""
            UPDATE cron_logs SET
                fim = %s,
                duracao_segundos = %s,
                status = %s,
                mensagem = %s
            WHERE id = %s;
            """, (now_str, duracao_seg, status, mensagem, log_id))
            cursor.execute("""
            UPDATE cron_schedules SET
                ultimo_status = %s,
                ultimo_log = %s,
                updated_at = CURRENT_TIMESTAMP
            WHERE task_id = %s;
            """, (status, mensagem, task_id))
        else:
            cursor.execute("""
            UPDATE cron_logs SET
                fim = ?,
                duracao_segundos = ?,
                status = ?,
                mensagem = ?
            WHERE id = ?;
            """, (now_str, duracao_seg, status, mensagem, log_id))
            cursor.execute("""
            UPDATE cron_schedules SET
                ultimo_status = ?,
                ultimo_log = ?,
                updated_at = CURRENT_TIMESTAMP
            WHERE task_id = ?;
            """, (status, mensagem, task_id))
        conn.commit()
    except Exception as e:
        print(f"Erro ao finalizar log de cron: {e}")
    finally:
        cursor.close()
        conn.close()

def get_recent_cron_logs(limit: int = 30) -> pd.DataFrame:
    """Retorna o histórico recente de execuções do agendador."""
    setup_cron_tables()
    conn = get_connection()
    query = """
    SELECT l.id, s.nome as rotina, l.inicio, l.fim, l.duracao_segundos as duracao_s, l.status, l.mensagem
    FROM cron_logs l
    LEFT JOIN cron_schedules s ON l.task_id = s.task_id
    ORDER BY l.id DESC
    LIMIT ?
    """
    if DB_TYPE in ["postgres", "postgresql"]:
        query = query.replace("?", "%s")
    df = pd.read_sql_query(query, conn, params=(limit,))
    conn.close()
    return df
