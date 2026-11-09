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

## 🔒 4. Arquitetura de Processo Único & Lock de Concorrência

Para evitar que múltiplos processos em segundo plano (como o processo mestre `init.py` e o servidor web `dashboard.py`) disparem as mesmas rotinas concorrentemente:
1. **Lock Global de Daemon (`DAEMON_LOCK_FILE`):** O `BancadaCronDaemon` registra o PID do processo primário. Instâncias secundárias entram automaticamente em modo cliente (apenas leitura e disparos manuais sob demanda), prevenindo duplicação de loops.
2. **Proteção de Tarefas em Execução:** A limpeza de tarefas marcadas como `'executando'` é restrita ao boot do sistema (`cleanup_interrupted_cron_tasks`), impedindo que consultas corriqueiras da interface alterem indevidamente o status de tarefas ativas para erro.
3. **Detecção Inteligente do Orquestrador:** Se uma instância do orquestrador já estiver em andamento, novas tentativas do cron reconhecem a atividade concorrente e registram aviso informativo sem gerar falhas espúrias no histórico.

