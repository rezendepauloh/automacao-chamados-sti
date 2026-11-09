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

## 📋 2. Modal de Preparação & Geração de HTML para o OTRS

Ao finalizar a preparação de um lote de máquinas, o analista clica em **"Gerar Relatório de Preparação"**. O sistema gera automaticamente um bloco HTML estilizado pronto para colagem:

### Como Utilizar:
1. Na aba **Redistribuição**, acesse o painel **Preparo de Máquinas**.
2. Selecione os computadores que foram preparados (marcando destino como *Doação* ou *Baixa*).
3. O modal exibirá uma caixa com o código HTML gerado e um botão **Copiar HTML**.
4. No chamado correspondente no **OTRS**:
   - Abra a caixa de resposta ou nota interna.
   - Clique no ícone de visualização de código HTML do editor do OTRS (`<>` ou botão HTML).
   - Cole o código copiado.
   - O OTRS renderizará a tabela formatada com cabeçalhos azuis institucionais, quantidades e listagem por patrimônio e número de série.

> [!NOTE]
> Por diretrizes de segurança contra XSS do OTRS, tags `<script>` ou campos de input interativos em JavaScript são sanitizados pelo editor do OTRS. Por isso, as tabelas utilizam formatação semântica CSS pura.

---

## 🚚 3. Otimização de Rotas de Entrega

O sistema calcula distâncias entre a sede (Campo Grande) e as comarcas do interior cadastradas:
- Utiliza algoritmo de caminho mais curto para agrupar polos regionais (ex: Polo Dourados, Polo Três Lagoas, Polo Corumbá).
- Minimiza custos de deslocamento da equipe técnica em viagens de entrega de novos parques de TI.

