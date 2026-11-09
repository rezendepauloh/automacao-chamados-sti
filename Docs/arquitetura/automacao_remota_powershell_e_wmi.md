# Automação Remota, PowerShell e Integração SCCM/WMI

> **Diretório:** [`src/scripts_powershell/`](../../src/scripts_powershell/)  
> **Testes de Integridade:** [`tests/scripts/test_scripts.py`](../../tests/scripts/test_scripts.py)

---

## 🎯 1. Visão Geral

O Sistema Bancada STI possui um ecossistema integrado para diagnósticos avançados, auditoria e manutenção de computadores corporativos do MPMS diretamente a partir do navegador web, combinando o protocolo personalizado do Windows `bancada://` com scripts PowerShell remotos.

---

## 🪟 2. Protocolo Personalizado (`bancada://`)

Para que o analista clique em um botão na interface do Streamlit e a estação abra uma janela nativa do PowerShell executando a tarefa selecionada:

1. O instalador de protocolo do Windows registra a chave de registro:
   ```reg
   [HKEY_CLASSES_ROOT\bancada]
   @="URL:Protocolo Bancada STI"
   "URL Protocol"=""
   [HKEY_CLASSES_ROOT\bancada\shell\open\command]
   @="powershell.exe -ExecutionPolicy Bypass -File \"C:\\Bancada\\Launcher.ps1\" \"%1\""
   ```
2. O launcher faz o parsing do URI (ex: `bancada://analisador?ip=10.111.64.50`) e dispara o script correspondente com privilégios adequados.

---

## 📜 3. Scripts PowerShell Disponíveis

| Script | Função Principal |
| :--- | :--- |
| **`Analisador.ps1`** | Coleta diagnóstico completo da estação de trabalho remota (hardware, SMART de discos, uso de CPU/RAM, portas abertas, processos travados) |
| **`GeradorHtml.ps1`** | Converte os logs e relatórios de auditoria coletados em um painel HTML interativo |
| **`LimpezaEstacao.ps1`** | Limpeza remota de diretórios temporários (`%temp%`, `C:\Windows\Temp`), caches de navegadores e logs órfãos |
| **`RemoverUsuariosLocais.ps1`** | Expurgo de perfis de usuários locais antigos ou órfãos para liberação de espaço em disco |
| **`sccm_sync.ps1`** | Consulta o banco de inventário do SCCM/MECM via WMI e sincroniza workstations ativas |

---

## ⚠️ 4. Requisito Obrigatório: Encoding UTF-8 com BOM (`utf-8-sig`)

O Windows PowerShell legado (versão 5.1 integrada ao Windows 10 e 11) interpreta arquivos `.ps1` salvos em UTF-8 puro (sem BOM) como código ANSI/Windows-1252. Isso causa quebra sistemática de acentos, mensagens de console corrompidas e erros sintáticos em caracteres especiais como `ã`, `é`, `ç`.

> [!IMPORTANT]
> **Regra Obrigatória:** Todo arquivo `.ps1` criado ou editado no repositório **DEVE ser salvo com UTF-8 com BOM** (`utf-8-sig`).
> A suíte de testes automatizados (`tests/scripts/test_scripts.py::test_powershell_scripts_have_utf8_bom`) valida estritamente a presença dos 3 bytes `\xef\xbb\xbf` no início de cada arquivo `.ps1`.

---

## 📡 5. Integração com SCCM / WMI

Para localização física hiperprecisa de chamados abertos:
- O sistema consulta o servidor SCCM (System Center Configuration Manager) via WMI.
- Cruza o nome da máquina ou login do usuário com a tabela de inventário.
- Obtém o endereço IPv4 ativo e calcula a máscara CIDR para identificar a promotoria/sala física exata.

