# Sessions — diário do projeto P × Q do faturamento (Enel CE e Enel RJ)

Diário do projeto, no estilo scrum: milestones, pendências e decisões importantes. Não registra cada
conversa. Alterações neste arquivo só com autorização do usuário.

Tem quatro partes:

1. **Milestones.** Entregas e marcos do projeto.
2. **Pendências.** Pontos que dependem de decisão ou de confirmação de alguma área, com o status.
3. **Informações a levantar.** Dados que, se obtidos, eliminariam uma aproximação do processo.
4. **Registro de decisões.** O que foi decidido, quando e por quê.

Uma pendência resolvida continua na parte 2, com status Resolvida ou Encerrada e a data. A regra
resultante passa a constar no [README](README.md) ou no
[procedimento](1-Entrega/procedimentos/PROCEDIMENTO_PXQ.md), e não volta a ser descrita aqui.

---

## 1. Milestones

| Data | Milestone | Observação |
|---|---|---|
| 03/11/2026 | Especificação técnica da base de faturamento (v1.0) enviada | `docs/especificacao/` |
| 26/11/2026 | Estrutura e templates da documentação criados (README, procedimento, sessions) | A partir da documentação de SP |

## 2. Pendências

| # | Tema | O que falta | Com quem | Aberta em | Status | Data status | Observações |
|---|---|---|---|---|---|---|---|
| P1 | Atendimento da especificação | Confirmar se a base virá no formato da especificação v1.0 (4 arquivos + códigos tarifários + controle) e se o AT virá cliente a cliente. | TI / Faturamento | 03/11/2026 | Aberta | | |
| P2 | Optantes pela tarifa B | A especificação os coloca no BT com subclasse "Optante B"; tratamento definitivo (BT ou AT) a reavaliar. | Interno | 03/11/2026 | Aberta | | Mesmo tema da RNF de CE e RJ |
| P3 | Classificação GD I, GD II e GD III na base SAMP | A separação entre GD I, GD II e GD III nos `tipo_arquivo` da base SAMP não está correta; só o total do GD está certo. A separação é mantida. Não altera o P × Q, só a classificação dessas linhas. A área vai corrigir, e a correção será replicada automaticamente na base. Confirmar quando estiver corrigido. | Faturamento | 26/11/2026 | Aberta | | |

## 3. Informações a levantar

| # | Informação | O que melhora | Status |
|---|---|---|---|
| I1 | PCAT 2024 de CE e RJ | Tarifa do REFAT e do CNR com referência em 2024 | Aberta |
| I2 | Calendário de leitura por lote (BT) | Tarifa proporcional nos meses de reajuste | Aberta |
| I3 | Informe de gestão mensal, DMR da Tarifa Social, contas contábeis | Validação e fechamento (especificação, seção 7) | Aberta |

## 4. Registro de decisões

| Data | Decisão | Motivo | Onde está a regra |
|---|---|---|---|
| 26/11/2026 | CE e RJ num só projeto (pasta `CE_RJ/FAT`), com README, procedimento e sessions próprios; diferenças entre as duas marcadas no texto. | Mesmo sistema de faturamento nas duas distribuidoras. | README |
| 26/11/2026 | Método, regras e convenções de SP valem para CE e RJ, salvo decisão em contrário registrada aqui. | Replicação do método de SP. | README, seção 0 |
| 27/11/2026 | Desconto Social da vigência 2025 (RJ e CE): mantida a tarifa derivada da Residencial sem os itens isentos (CDE, CDE Covid, CDE GD), mesma regra de SP. Estudado e comprovado: a tarifa derivada é igual à tarifa de aplicação do Mercado Receita do reajuste de 2026 (RJ jan/2026: faixa 01 810,44 e faixa 02 925,39 R$/MWh; CE jan e fev/2026: 645,10 e 710,07 R$/MWh). A diferença no item `TUSD Subsidio` entre a regra e o PCAT de 2026 não afeta o cálculo, pois nas vigências de 2026 vale a tarifa publicada. Assunto encerrado. | Os PCATs de 2025 não publicam o Desconto Social, faturado desde jan/2026. | Procedimento, Passo 4 |
| 27/11/2026 | Cativos, só na fase de desenvolvimento: o P × Q roda sobre a base SAMP de documentos completa (cerca de 3 milhões de registros por mês no RJ, documento a documento) e sobre a extração agrupada (SQL agregado), e os dois resultados são comparados. Não é etapa do procedimento mensal. | Verificar se o agrupamento altera o resultado; a base documento a documento também permite conferir, cliente a cliente, descontos e tarifas aplicadas (por exemplo, desconto concedido a mais, que deixaria item tarifário negativo). | — |
