# Guia de Gestão de Bancada, Redistribuição e Termos

> **Módulo:** [`src/tabs/redistribuicao.py`](../../src/tabs/redistribuicao.py)  
> **DAO:** [`src/database/donations_db.py`](../../src/database/donations_db.py)

---

## 🎯 1. Visão Geral do Módulo

O módulo de **Redistribuição & Bancada** gerencia o ciclo de vida físico de microcomputadores, notebooks e periféricos do MPMS após a substituição ou desativação:

1. **Equipamentos em Bancada:** Máquinas recebidas do interior ou sede para testes, formatação ou troca de peças.
2. **Preparo para Doação:** Máquinas revisadas e higienizadas destinadas a instituições beneficentes, escolas ou prefeituras conveniadas.
3. **Preparo para Baixa Patrimonial:** Equipamentos obsoletos ou irrecuperáveis que passarão por descarte ecológico ou leilão.
4. **Cálculo Logístico de Rotas:** Otimização do itinerário para transporte e entrega em comarcas do interior.

---

## 📋 2. Modal de Preparação & Geração de HTML para o DMP (OTRS)

Ao finalizar a preparação de um lote de máquinas, o analista clica em **"📝 Gerar Texto do Chamado"**. O sistema gera automaticamente um assunto padronizado e um bloco HTML estilizado pronto para colagem no chamado:

### Campos e Colunas do Relatório:
- **`#`**: Numeração sequencial ordinal.
- **`Patrimônio`**: Número do patrimônio do equipamento tombado.
- **`Modelo`**: Modelo do equipamento (ex: HP EliteDesk 800 G5, Dell OptiPlex 7050).
- **`Serial Number PC`**: Número de série do gabinete/computador.
- **`Equipamento`**: Tipo do equipamento (Computador, Notebook, etc.).
- **`Memória RAM`**: Exibido dinamicamente quando preenchido na planilha oficial (ex: `16 GB (2x 8GB)`, `8 GB (1x 8GB)`), atendendo aos questionamentos da administração superior e do DMP sobre a quantidade e capacidade dos pentes entregues.
- **`SSD / HD`**: Capacidade e tipo de armazenamento instalado (ex: `256 GB SSD`, `500 GB HD`).
- **`Motivo/Obs`**: Observações técnicas ou motivo de baixa patrimonial.

### Resguardo Institucional Universal (Doação, Baixa, Redistribuição e Garantia):
Para todos os lotes que envolvem computadores/notebooks, o relatório insere automaticamente observações individuais por tipo de movimentação e a **Nota de Conferência e Entrega (DMP)** no rodapé:
> *"📌 **Nota de Conferência e Entrega:** Registramos que todas as máquinas destinadas a [movimentações] relacionadas neste chamado saíram da bancada técnica e foram entregues completas ao setor responsável, devidamente montadas com seus respectivos discos de armazenamento (HD/SSD) e memórias RAM."*

### Como Utilizar:
1. Na aba **Redistribuição**, acesse o painel **Gerador de Texto (Preparo)** na barra lateral.
2. Selecione a **Data de Preparo** do lote.
3. Clique em **"📝 Gerar Texto do Chamado"**.
4. O modal exibirá o assunto do chamado, o texto renderizado e o código-fonte HTML.
5. No chamado correspondente no **OTRS**:
   - Abra a caixa de resposta ou nota interna.
   - Clique no ícone de visualização de código HTML do editor do OTRS (`<>` ou botão HTML).
   - Cole o código copiado e salve.

---

## 🚚 3. Otimização de Rotas de Entrega

O sistema calcula distâncias entre a sede (Campo Grande) e as comarcas do interior cadastradas:
- Utiliza algoritmo de caminho mais curto para agrupar polos regionais (ex: Polo Dourados, Polo Três Lagoas, Polo Corumbá).
- Minimiza custos de deslocamento da equipe técnica em viagens de entrega de novos parques de TI.

---

## 📊 4. Sincronização da Planilha de Doações e Baixas

A planilha de doações pode ser sincronizada via link oficial do SharePoint ou por arquivo local/OneDrive:
- **Prioridade 1:** Link do SharePoint ou caminho configurado na aba **Configurações > 📂 SharePoint & Planilhas** (`DONATIONS_EXCEL_RELATIVE_PATH`).
- **Prioridade 2 (Fallback):** Variável `DONATIONS_EXCEL_RELATIVE_PATH` no arquivo `.env`.
- **Download Automatizado:** Ao informar uma URL do SharePoint (`https://...`), o robô (`src/syncs/sync_donations.py`) efetua o download para `uploads/doacoes_redistribuicao.xlsx` e sincroniza a base de dados mantendo os dados protegidos contra caminhos corrompidos.

