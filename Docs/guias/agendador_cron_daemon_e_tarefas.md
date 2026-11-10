# Guia do Agendador de Tarefas em Segundo Plano (BancadaCronDaemon)

> **Módulo do Serviço:** [`src/services/cron_scheduler.py`](../../src/services/cron_scheduler.py)  
> **DAO:** [`src/database/cron_db.py`](../../src/database/cron_db.py)  
> **Interface de Controle:** [`src/tabs/configuracoes.py`](../../src/tabs/configuracoes.py)

---

## 🎯 1. O que é o BancadaCronDaemon?

O `BancadaCronDaemon` é uma thread em background executada continuamente dentro do contêiner Docker da aplicação. Ele é responsável por disparar rotinas periódicas de manutenção, sincronização e inteligência sem travar a interface do usuário nem exigir intervenção manual.

---

## 📅 2. Tarefas Agendadas Padrão

| ID da Tarefa | Nome Amigável | Categoria | Intervalo Padrão | Descrição |
| :--- | :--- | :--- | :--- | :--- |
| **`retrain_ia_gridsearch`** | Recalibração Mensal da IA | Inteligência Artificial | 30 dias (às 02:00) | Executa busca exaustiva de hiperparâmetros (GridSearchCV) para recalibrar o modelo campeão |
| **`sync_sccm`** | Sincronização SCCM / WMI | Inventário | 6 horas | Atualiza a base de estações ativas e seus IPs mais recentes |
| **`sync_ramais`** | Atualização de Ramais MPMS | Comunicação | 24 horas | Varre a lista institucional de ramais e titulares |
| **`sync_plantoes`** | Sincronização de Plantões | Escalas | 12 horas | Coleta escalas atualizadas de sobreaviso do SIMP |
| **`limpeza_logs`** | Rotação de Logs Antigos | Manutenção | 7 dias | Expurga arquivos de log legados e compacta históricos |

---

## 🎛️ 3. Gerenciamento via Interface Gráfica

Na aba **⚙️ Configurações** do dashboard:
1. Acesse o sub-painel **Cron & Tarefas Agendadas**.
2. É possível:
   - Ativar ou desativar qualquer rotina com um clique (*toggle*).
   - Ajustar o intervalo de disparo (minutos, horas ou dias).
   - Visualizar o horário exato da última execução e da próxima execução programada.
   - Disparar uma tarefa manualmente de forma avulsa clicando em **Executar Agora**.

---

## 🔒 4. Arquitetura de Processo Único & Isolamento de Concorrência

Para garantir que tarefas em segundo plano executem de forma fluida, sem colisão e sem interferir umas nas outras:

1. **Lock Global de Daemon (`DAEMON_LOCK_FILE`):** O `BancadaCronDaemon` registra o PID do processo primário. Instâncias secundárias entram automaticamente em modo cliente (apenas leitura e disparos manuais sob demanda), prevenindo duplicação de loops de agendamento.
2. **Execução Concorrente em Threads Isoladas:** Cada tarefa disparada pelo cron é executada em sua própria thread assíncrona (`threading.Thread`). Uma rotina de longa duração (como a recalibração da IA ou raspagem do SIMP) **não bloqueia** o daemon de executar outras tarefas em seus horários programados.
3. **Locks Dedicados por Domínio (`.lock`):** Cada tarefa possui seu próprio arquivo de lock temporário (ex: `portarias_sync.lock`, `donations_sync.lock`, `garantia_sync.lock`, `plantoes_matutino_sync.lock`, `plantoes_semanal_sync.lock`). Dessa forma, tarefas distintas nunca bloqueiam umas às outras.
4. **Isolamento de Concorrência da Mesma Tarefa (`_executing_tasks`):** O daemon monitora o conjunto `_executing_tasks`. Se uma tarefa já estiver em andamento, o agendador não dispara uma segunda execução simultânea da mesma tarefa.
5. **Orquestrador de Chamados sem Auto-Bloqueio:** O orquestrador opera como subprocesso desacoplado. O despachante do cron verifica exclusivamente se há outro processo ativo (`check_process_running`) no arquivo de lock do SO (`automated_otrs_citsmart.lock`), impedindo falsos positivos de auto-bloqueio entre a thread do daemon e o processo executor.
6. **Proteção Contra Tarefas Interrompidas:** A limpeza de registros marcados como `'executando'` ocorre apenas no boot inicial (`cleanup_interrupted_cron_tasks`), recuperando graciosamente qualquer tarefa interrompida por desligamento do contêiner.

