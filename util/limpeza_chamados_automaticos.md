# Remoção de Chamados Automáticos do Banco de Dados SQLite

Este documento orienta como remover do banco de dados `chamados.db` os chamados gerados automaticamente pelos usuários:
- **Monitoramento Adm Mpms**
- **Adm Ticket Por Email**

Como o executável de linha de comando `sqlite3` pode não estar instalado no Windows, a forma mais prática e nativa é executar via Python.

---

## 1. Executando via Script Python (Recomendado)

Criei um script utilitário em `util/limpar_chamados.py`. Basta executar no terminal:

```powershell
python util/limpar_chamados.py
```

---

## 2. Executando via Comando de Linha Única no Python

Caso não queira rodar o arquivo, você pode executar diretamente no PowerShell/Terminal:

```powershell
python -c "import sqlite3; conn = sqlite3.connect('chamados.db'); cursor = conn.cursor(); users = ('Monitoramento Adm Mpms', 'Adm Ticket Por Email'); cursor.execute('DELETE FROM comentarios WHERE chamado_id IN (SELECT id FROM chamados WHERE usuario IN (?, ?))', users); cursor.execute('DELETE FROM chamados WHERE usuario IN (?, ?)', users); conn.commit(); conn.close(); print('Limpeza concluída!')"
```

---

## 3. Comandos SQL (Referência)

Se você utilizar algum SGBD visual (como SQLite Browser / DBeaver) ou tiver o `sqlite3` CLI instalado:

```sql
-- Remove os comentários associados aos chamados
DELETE FROM comentarios 
WHERE chamado_id IN (
    SELECT id FROM chamados 
    WHERE usuario IN ('Monitoramento Adm Mpms', 'Adm Ticket Por Email')
);

-- Remove os chamados
DELETE FROM chamados 
WHERE usuario IN ('Monitoramento Adm Mpms', 'Adm Ticket Por Email');
```
