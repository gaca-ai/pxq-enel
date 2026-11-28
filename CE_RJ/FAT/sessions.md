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
| 27/11/2026 | P × Q dos livres e geradoras do RJ, out/2026, concluído e resultado enviado por e-mail | Receita explicada por componente 98,6%; não explicado 1,5 MR$ |

## 2. Pendências

| # | Tema | O que falta | Com quem | Aberta em | Status | Data status | Observações |
|---|---|---|---|---|---|---|---|
| P1 | Atendimento da especificação | Confirmar se a base virá no formato da especificação v1.0 (4 arquivos + códigos tarifários + controle) e se o AT virá cliente a cliente. | TI / Faturamento | 03/11/2026 | Encerrada | 28/11/2026 | Todos os arquivos já vêm no layout SAMP |
| P2 | Optantes pela tarifa B | A especificação os coloca no BT com subclasse "Optante B"; tratamento definitivo (BT ou AT) a reavaliar. | Interno | 03/11/2026 | Aberta | | Mesmo tema da RNF de CE e RJ |
| P3 | Classificação GD I, GD II e GD III na base SAMP | A separação entre GD I, GD II e GD III nos `tipo_arquivo` da base SAMP não está correta; só o total do GD está certo. A separação é mantida. Não altera o P × Q, só a classificação dessas linhas. A área vai corrigir, e a correção será replicada automaticamente na base. Confirmar quando estiver corrigido. | Faturamento | 26/11/2026 | Aberta | | |
| P4 | Livres: distribuidoras com receita e sem quantidade | 3 contas REVENDA com receita de energia fora ponta sem energia nem demanda (+1,78 MR$, principal parcela do não explicado de 1,5 MR$). O que é faturado? lista de exceções (`lista_excecoes_RJ_2026-10.xlsx`), aba `Por_conta`, motivo "RECEITA SEM QUANTIDADE" e `ZnTemp.` REVENDA. | Faturamento | 27/11/2026 | Aberta | | Enviada por e-mail em 27/11. O filtro enviado foi só pelo motivo, que traz também 1 geradora A3A_GER (R$ 320,11) |
| P5 | Livres: contas faturadas pela tarifa APE | 10 contas PPFED com chave de livre faturadas pela tarifa APE (−0,32 MR$ no não explicado). Confirmar se são APE; a base não discrimina o APE. lista de exceções (`lista_excecoes_RJ_2026-10.xlsx`), aba `APE`. | Faturamento | 27/11/2026 | Aberta |  | Enviada por e-mail em 27/11. Se confirmado, valorizar pela tarifa APE (linha (e) Isenção APE) |
| P6 | Livres: ICMS DEVEC | A base LV traz o ICMS total (nosso e DEVEC). Confirmar que a receita de fornecimento dos livres do Informe exclui o ICMS DEVEC e inclui os demais tributos (fecha com a base a 0,4 MR$). Rural e Uso da Distribuição diferem da base por classe (+4,62 e −5,05 MR$). | Faturamento / Controladoria | 27/11/2026 | Aberta | | Inferido pela conciliação |
| P7 | Livres: documentos sem energia e sem demanda | 2.333 documentos de 29/10 sem energia e sem demanda (21,91 MR$ de total faturado), de contas que também têm documento de energia no mês. São ERCAP? | Faturamento | 27/11/2026 | Aberta |  | Os documentos não têm valor em nenhuma coluna de receita nem de tributos, só demanda contratada (`DEM_CONTRAT_*`); os 21,9 MR$ caem inteiros na diferença do `TOTAL_FAT` (não tarifável, UD/ER) |
| P8 | Livres: composição do `TOTAL_FAT` | 7,17 MR$ sem coluna própria (24 documentos com receita e `TOTAL_FAT` zero; nos demais, `TOTAL_FAT` 14,41 MR$ abaixo da soma das colunas). Como o `TOTAL_FAT` é distribuído entre documentos da mesma conta? É possível trazer a energia reativa excedente em coluna própria? | Faturamento | 27/11/2026 | Aberta |  | Por ora, tratado como UD/ER (decisão de 28/11). Os 7,17 MR$ são líquidos: +21,9 MR$ dos documentos da P7 e −14,7 MR$ nos documentos de energia das mesmas contas |
| P9 | Livres: acessante das contas REVENDA | Distribuidora de cada uma das 25 contas faturadas pela tarifa "Distribuição". lista de exceções (`lista_excecoes_RJ_2026-10.xlsx`), aba `Por_conta`, motivo "RECLASSIFICADA COMO DISTRIBUIÇÃO", coluna "Acessante correto". | Faturamento | 27/11/2026 | Aberta | |  |
| P10 | Livres: geradoras | Nome da usina de cada geradora (aba `Geradoras`); desconto de 50% na TUSDg com `SUBSGERFI` zerado (desconto de 0,59 MR$ inferido pela receita); R$ 320,11 sem quantidade na conta A3A_GER com motivo "CHAVE SEM TARIFA". | Faturamento | 27/11/2026 | Aberta | |  |
| P11 | Livres: refaturamento de geradoras | 3 contas: estorno com 50% de desconto e conta nova sem desconto (−1,00 MR$ de desconto líquido). Por que o desconto foi retirado? Aba `Geradoras`, "Tipos de faturamento" = "FAT | REFAT". | Faturamento | 27/11/2026 | Aberta | |  |
| P12 | Livres: desconto diferente do `SUBSCARFI` | 41 contas com desconto pela tarifa diferente do informado (diferença líquida de 0,04 MR$). Aba `Divergencia_desconto`. | Faturamento | 27/11/2026 | Encerrada | 28/11/2026 | Diferença não justifica o esforço |
| P13 | Livres: tarifa fora da faixa da chave | 14 contas cuja receita não é explicada pela tarifa da chave; conferir categoria, subclasse e detalhe. Aba `Por_conta`, motivos "TARIFA APLICADA FORA DA FAIXA" e "DESCONTO FORA DE 0%, 50% OU 100%". | Faturamento | 27/11/2026 | Aberta | |  |
| P14 | Livres: Código Arquivo 06 | 52 documentos (−0,10 MR$): confirmar que são refaturamento e se é possível trazer o mês de referência da conta original. | Faturamento | 27/11/2026 | Encerrada | 28/11/2026 | Sem informação adicional neste primeiro momento. Código Arquivo: 05 = faturamento livres; 06 = refaturamento livres; 07 = outros |
| P15 | Livres: energia fora da base | 2.471 MWh entre a base (359.390 MWh) e o "Consumo Faturado (Livres)" do Informe (361.861 MWh). Confirmar que é o suprimento à Elektro. | Faturamento | 27/11/2026 | Aberta | |  |
| P16 | Cativos: composição do `total_fat` | 1,08 MR$ sem coluna própria na amostra (cerca de 11%): CIP, serviços, reativo excedente? É possível trazer o reativo em coluna própria? | Faturamento | 27/11/2026 | Aberta | | Ver Project doc `status_FAT_CE_RJ_cativo.md`. Por ora, tratado como UD/ER (decisão de 28/11) |
| P17 | Cativos: bandeira na GD | Nos `tipo_arquivo` 8, 9 e 10, `receita_bandeira` é menor que a soma das bandeiras por posto. O total é líquido da energia compensada? | Faturamento | 27/11/2026 | Aberta | |  |
| P18 | Cativos: consumo na GD | Os campos `cons_*` parecem trazer o consumo líquido da compensação (exemplo: opbel 310709790307). É possível informar a energia faturada, incluindo o custo de disponibilidade? | Faturamento | 27/11/2026 | Aberta | |  |
| P19 | Cativos: tabela do CE | Nome da tabela equivalente do CE (`bi_brce_bill.bt_brce_zisu_t_samp_m_dt`?). | TI | 27/11/2026 | Aberta | |  |
| P20 | Base: campos a incluir pela TI | Identificador de CNR (Consumo Não Registrado), lote de leitura (etapa posterior) e data da fatura original dos refaturamentos. | TI | 27/11/2026 | Aberta | | Sem a data da fatura original, vale a regra de vigência do refaturamento (decisão de 28/11) |

## 3. Informações a levantar

| # | Informação | O que melhora | Status |
|---|---|---|---|
| I1 | PCAT 2024 de CE e RJ | Tarifa do REFAT e do CNR com referência em 2024 | Resolvida (28/11/2026): vigências 2022, 2022_extraordinaria (13/07/2022), 2023 e 2024 de CE e RJ incluídas na `tarifa_long.csv` a partir dos dados abertos da ANEEL (componentes tarifárias; FIN = TA − BE − CVA), conferidas contra os PCATs; 2025 em diante segue pelo PCAT. Nas vigências da ANEEL, a TE de suprimento às distribuidoras (CERCI e outras) fica como publicada (zero) |
| I2 | Calendário de leitura por lote (BT) | Tarifa proporcional nos meses de reajuste | Aberta |
| I3 | Informe de gestão mensal, DMR da Tarifa Social, contas contábeis | Validação e fechamento (especificação, seção 7) | Aberta |

## 4. Registro de decisões

| Data | Decisão | Motivo | Onde está a regra |
|---|---|---|---|
| 26/11/2026 | CE e RJ num só projeto (pasta `CE_RJ/FAT`), com README, procedimento e sessions próprios; diferenças entre as duas marcadas no texto. | Mesmo sistema de faturamento nas duas distribuidoras. | README |
| 26/11/2026 | Método, regras e convenções de SP valem para CE e RJ, salvo decisão em contrário registrada aqui. | Replicação do método de SP. | README, seção 0 |
| 27/11/2026 | Desconto Social da vigência 2025 (RJ e CE): mantida a tarifa derivada da Residencial sem os itens isentos (CDE, CDE Covid, CDE GD), mesma regra de SP. Estudado e comprovado: a tarifa derivada é igual à tarifa de aplicação do Mercado Receita do reajuste de 2026 (RJ jan/2026: faixa 01 810,44 e faixa 02 925,39 R$/MWh; CE jan e fev/2026: 645,10 e 710,07 R$/MWh). A diferença no item `TUSD Subsidio` entre a regra e o PCAT de 2026 não afeta o cálculo, pois nas vigências de 2026 vale a tarifa publicada. Assunto encerrado. | Os PCATs de 2025 não publicam o Desconto Social, faturado desde jan/2026. | Procedimento, Passo 4 |
| 27/11/2026 | Cativos, só na fase de desenvolvimento: o P × Q roda sobre a base SAMP de documentos completa (cerca de 3 milhões de registros por mês no RJ, documento a documento) e sobre a extração agrupada (SQL agregado), e os dois resultados são comparados. Não é etapa do procedimento mensal. | Verificar se o agrupamento altera o resultado; a base documento a documento também permite conferir, cliente a cliente, descontos e tarifas aplicadas (por exemplo, desconto concedido a mais, que deixaria item tarifário negativo). | — |
| 28/11/2026 | Livres: a diferença entre o `TOTAL_FAT` do documento e as colunas de receita, bandeira e tributos é tratada como UD/ER (grupo `UD_ER_REATIVO`); a coluna Outros fica vazia. | Primeiro momento, até a área explicar a composição do `TOTAL_FAT` (P8). | Notebook, Passo 2 |
| 28/11/2026 | Cativos: a diferença entre o `total_fat` do documento e as colunas de receita, bandeira e tributos é tratada como UD/ER, mesma regra dos livres. | Primeiro momento, até a área explicar a composição do `total_fat` (P16). | Notebook, Parte II |
| 28/11/2026 | Cativos, vigência tarifária do refaturamento (definido; paliativo até a TI informar a data da fatura original, P20): a vigência de cada documento é identificada pela tarifa implícita, comparada com as vigências de 2022 a 2026 da tabela de tarifas (tolerância de 3%; havendo mais de uma vigência na faixa, vale a mais próxima, sem exceção, pois tarifas próximas não distorcem o resultado). Quando nenhuma vigência é identificada, vale a regra default (fatura original tratada como do mês anterior, com a data de faturamento equivalente) e o documento vai para a lista de exceções do mês. Proporcionalização pela receita descartada (explicaria a receita por construção). Ordem de grandeza, RJ out/2026: 32.054 documentos de refaturamento (20,81 MR$ em valor absoluto); exceções: 8.859 documentos, 4,03 MR$ em valor absoluto (0,50% do total faturado dos cativos, 813,13 MR$), −0,95 MR$ líquido; impacto máximo estimado na 2ª passada bruta: cerca de 0,6 MR$ (se todas as exceções fossem da vigência 2025 valorizadas pela 2026). Acompanhar a 2ª passada bruta dos documentos em exceção. | A base não traz a data da fatura original (P20); a base completa de out/2026 mostrou refaturamentos do RJ em tarifas de vigências anteriores. | Procedimento (a incluir) |
