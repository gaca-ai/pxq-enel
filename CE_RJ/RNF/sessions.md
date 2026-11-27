# Sessions — ata do projeto RNF (Enel CE e Enel RJ)

Registro das reuniões e conversas de trabalho da RNF de CE e RJ. Tem três partes:

1. **Pendências abertas.** Pontos que dependem de decisão ou de confirmação de alguma área.
2. **Informações a levantar.** Dados que, se obtidos, eliminariam uma aproximação do processo.
3. **Registro de decisões.** O que foi decidido, quando e por quê.

Quando uma pendência é resolvida, ela sai da parte 1 e entra na parte 3. A regra resultante passa a
constar no [README](README.md) ou no [procedimento](1-Entrega/procedimentos/PROCEDIMENTO_RNF.md), e
não volta a ser descrita aqui.

---

## 1. Pendências abertas

| # | Tema | O que falta | Com quem | Aberta em |
|---|---|---|---|---|
| P2 | Clientes do grupo A com demanda e sem consumo | 69 clientes, com 0,24 MR$ de demanda e energia zero. Hipótese: geração, provavelmente minigeração distribuída com a energia toda compensada (as classes são de consumo: comercial, poder público, industrial). Confirmar. Para o cálculo, só a demanda é valorizada. | Faturamento | 24/11/2026 |
| P3 | Optantes B com preço médio zero | 94 clientes B3 optantes com 13,9 MWh de energia não faturada e preço médio zero: a energia fica sem valor. Confirmar se é correto. | Faturamento | 24/11/2026 |
| P7 | Iluminação pública: B4a e B4b | O de-para de classes só tem o código `B4A_IP`. Confirmar se há IP na B4b e como é identificada (mesmo ponto da P18 de SP). | Faturamento | 24/11/2026 |
| P10 | TOI em Consumo Próprio (RJ) | 146 TOIs, 3.871 MWh e 4,28 MR$ em Consumo Próprio do grupo B. A energia não entra na linha "RNF (kWh)" do informe, mas o valor entra na linha "RNF (R$)". Verificar do que se trata. | Usuário | 25/11/2026 |

## 2. Informações a levantar

Em ordem de prioridade.

| # | Informação | O que melhora |
|---|---|---|
| I1 | Demanda em kW no arquivo do sistema | A demanda passa a ser valorizada pela quantidade, sem depender da composição do faturado |
| I2 | Energia por posto horário (ponta e fora ponta) | A energia de ponta passa ao bloco 1, sem depender do faturado do mês |
| I3 | Valor separado entre TE e TUSD | A divisão do valor deixa de depender da composição do faturado |

## 3. Registro de decisões

| Data | Decisão | Motivo | Onde está a regra |
|---|---|---|---|
| 24/11/2026 | CE e RJ são tratados no mesmo processo e na mesma pasta (`CE_RJ`), com a estrutura de SP. | As duas distribuidoras usam o mesmo sistema comercial. | README |
| 24/11/2026 | A fonte da RNF de sistema são os arquivos do sistema comercial, um por nível de tensão (AT e BT), uma linha por cliente, com a mesma estrutura nos dois níveis. | Arquivo de exemplo recebido. | Procedimento, seção 1 |
| 24/11/2026 | Os arquivos do sistema contêm só clientes cativos: energia em TE e TUSD na mesma quantidade. | Os livres virão em arquivo separado (P4). | README 4 |
| 24/11/2026 | Valores com PIS/COFINS e sem ICMS, como em SP; valor líquido = bruto / (1 + alíquota). | Confirmado; coerente com o arquivo (PIS/COFINS da demanda = 4,29% da base). | Procedimento, seção 2 |
| 24/11/2026 | O arquivo BT tem exatamente o mesmo layout do AT, com cerca de 1 GB por mês. A leitura será feita em blocos e o arquivo convertido para parquet antes do processamento. | Informado pela área. Fecha a antiga pendência de exemplo do BT. | Procedimento, seção 1 |
| 24/11/2026 | Combinações de classe e código de tarifa incoerentes (ex.: subclasse comercial com `B1_RESID`; `REPLN` com `A4_VD`) são tratadas como optante B. | Definido pelo usuário. | Procedimento, seção 4 |
| 24/11/2026 | Subclasses da Residencial Baixa Renda: `REBXR` geral (Tarifa Social), `REBRBPC` BPC, `REBRIND` indígena, `REBRQUI` quilombola (faixas estimadas como em SP) e `REBXDS` Desconto Social (chave própria). | Regras da ANEEL em `docs/tarifas_baixa_renda_aneel.xlsx`. | Procedimento, seção 4.2 |
| 24/11/2026 | A classe vem da `SUBCLASSE`, pelo de-para `docs/CL_SUB_TAR.xlsx` (45 subclasses, 9 classes). O campo `CLASSE` do arquivo não é usado. | O código numérico de `CLASSE` não corresponde à classe de consumo (a classe 10, por exemplo, mistura residencial, comercial, rural e serviço público). | Procedimento, seção 4 |
| 24/11/2026 | Posto horário da energia do grupo A, abertura da demanda por item tarifário e divisão do valor entre TE e TUSD pela composição do faturado do mês (mesmo subgrupo e modalidade). | O arquivo não traz posto, traz a demanda só em R$ e um preço médio único. Mesmo critério já usado em SP. | README 4; procedimento, seção 4 |
| 25/11/2026 | O arquivo de exemplo `CLIENTES_ALTA_AG_28092026_082440.txt` é do RJ. As bases recebidas até aqui são amostras; as definitivas virão por link do SharePoint. | Confirmado pelo usuário. Fecha a P1. | — |
| 25/11/2026 | Todo o valor da RNF (sistema, TOI, Pendentes e Livres) é aberto proporcionalmente a partir do valor do fechamento. | Nenhum informe traz posto horário nem separação entre energia e demanda: todos partem de energia por grupo e classe valorizada por tarifa média. | Procedimento |
| 25/11/2026 | TOI e Pendentes: energia por grupo (A/B) e classe × tarifa média da RNF de cativos do SAP do mês, com PIS/COFINS. A alíquota é obtida em cada arquivo, a cada mês, pelo quociente entre a tarifa aplicada e a tarifa "sem PIS e COFINS"; o quociente deve ser igual em todas as classes, dentro de uma tolerância parametrizável. Subgrupo, modalidade e subclasse vêm da mistura da RNF de sistema da mesma classe. | A alíquota muda a cada mês. Mesmo critério dos ajustes manuais de SP. | Procedimento |
| 25/11/2026 | Livres: energia medida do mês inteiro × tarifa média de TUSD (SAMP de uso do mês anterior), aberta por classe em percentuais fixos. O valor inclui PIS/COFINS, apesar do título "Sem Tributos". Posto, demanda e desconto de fonte incentivada vêm da aba `Tarifa Média` do próprio informe (SAMP cliente a cliente). | Confirmado pelo usuário. A aba `Tarifa Média` é o faturado dos livres no mesmo critério que gerou a tarifa média. | Procedimento |
