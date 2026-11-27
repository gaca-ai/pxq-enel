# RNF — Abertura da Renda Não Faturada por item tarifário (Enel CE e Enel RJ)

Ponto de entrada da RNF de Ceará e Rio de Janeiro. As duas distribuidoras usam o mesmo sistema comercial e
recebem o mesmo tratamento; diferenças entre elas, quando houver, são indicadas no texto. O detalhamento
técnico está no [procedimento](1-Entrega/procedimentos/PROCEDIMENTO_RNF.md). Decisões e pendências com as
áreas ficam em [`sessions.md`](sessions.md). O contexto, o objetivo e o embasamento regulatório são os da
RNF de SP ([README de SP](../../SP/RNF/README.md), seções 0 a 3) e não são repetidos aqui.

Valores monetários em milhões de reais (MR$), energia em MWh e demanda em MW.

---

## 4. Implementação

Situação: em especificação. O processo segue o da RNF de SP, com a fonte trocada: no lugar do ZUBIL_146 do
SAP, os arquivos de RNF do sistema comercial, um por nível de tensão (AT e BT), uma linha por cliente.

**Preparação**

- **Arquivos do sistema (AT e BT).** Energia não faturada (com fator), valor da energia, demanda em R$,
  bandeira e PIS/COFINS por cliente. Somente clientes cativos; os livres virão em arquivo separado.
- **Ajustes manuais.** A definir (mesmo papel da base de ajustes de SP).
- **PIS/COFINS.** Os valores do arquivo contêm PIS/COFINS e não contêm ICMS, como em SP. O valor líquido é
  o bruto / (1 + alíquota).

**Valorização**

- **Bloco 1.** Energia × (TE + TUSD) homologadas, na mesma quantidade para as duas componentes (cativo).
- **Bloco 2.** A diferença para o valor do arquivo é aberta pela composição do faturado do mesmo mês
  (mesmo subgrupo e modalidade), que também dá:
  - a separação da energia do grupo A entre ponta e fora ponta (o arquivo não traz posto);
  - a abertura da demanda por item tarifário (o arquivo traz a demanda só em R$);
  - a divisão do valor entre TE e TUSD (o arquivo traz um preço médio único).

**Saída e conferência**

- Mesmas saídas de SP: relatório detalhado e consolidado `receita_nao_faturada`, no layout Receita
  Faturada, com provisão, estorno, saldo e saldo YTD.

## 5. Resultados obtidos

Ainda não há resultados. O exemplo recebido (AT, referência set/2026, 5.323 clientes) foi usado para
validar as regras de cálculo do arquivo; os totais estão no procedimento, seção 3.
