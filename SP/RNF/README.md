# RNF — Abertura da Renda Não Faturada por item tarifário (Enel Distribuição São Paulo)

Este documento é o ponto de entrada da fase II do projeto P × Q. Ele explica o problema, o que foi
construído, onde está cada peça e quais resultados foram obtidos. As decisões e as pendências com as
áreas ficam em [`sessions.md`](sessions.md). As regras tarifárias comuns às duas fases estão no
[README do Faturamento](../BW/README.md), seção 3, e não são repetidas aqui.

---

## 0. Contexto geral

A RNF (Renda Não Faturada) é a receita da energia já entregue e ainda não faturada no fechamento do
mês: o consumo entre a última leitura de cada cliente e o último dia do mês. Ela é reconhecida pelo
regime de competência e revertida no mês seguinte, quando a conta é emitida.

A contabilidade apura a RNF em um relatório mensal (`RNF CCS <mmaaaa> Fator K Nova
Metodologia.xlsb`). A maior parte vem do SAP (transação ZUBIL_146, que calcula a RNF de cada lote,
classe e categoria tarifária); o restante vem de ajustes manuais (faturas pendentes, art. 323, TOI,
iluminação pública, ajuste de carga), valorizados pela tarifa média do ZUBIL de cada classe. O
resultado é um valor total, sem a composição tarifária.

Na fase I, o P × Q passou a abrir a receita **faturada** nos 37 itens da tarifa. Sem a RNF, a receita
do mês por item fica incompleta: o que entra numa competência e é revertido na seguinte não tem
abertura.

## 1. Objetivo

Produzir, todo mês, a **abertura da RNF contabilizada por item tarifário, família e camada**, com três
características:

- **Fecha com o relatório contábil.** O valor aberto, somado ao PIS/COFINS, ao reativo e à bandeira,
  é igual ao total do relatório.
- **Mesma régua do Faturamento.** Mesma tabela de tarifas homologadas, mesmos 37 itens, mesmas sete
  famílias e as quatro camadas (TA = BE + CVA + FIN). Faturamento e RNF somam item a item.
- **Mesma regra de ajuste do Faturamento.** O bloco 1 (quantidade × tarifa) é levado ao valor
  registrado pela mesma lógica do bloco 2 do BW: regime fio para o livre do grupo A, rateio
  proporcional para o restante.

## 2. Estrutura do projeto

```
SP/RNF/
├── README.md                     este documento
├── sessions.md                   ata: decisões tomadas e pendências com as áreas
├── 1-Entrega/
│   ├── Abertura_RNF.ipynb        o cálculo, passo a passo
│   └── rnf/                      saída mensal do notebook: rnf_<aaaa-mm>.xlsx
├── src/
│   └── RNF CCS <mmaaaa> ... .xlsb   relatório contábil da RNF do mês
└── arquivo_morto_2026-10-04.zip  análise inicial e versão anterior (base BIP)
```

| Pasta / arquivo | Para que serve |
|---|---|
| [`1-Entrega/Abertura_RNF.ipynb`](1-Entrega/Abertura_RNF.ipynb) | O processo completo. Lê o relatório contábil, monta as linhas do SAP e dos ajustes manuais, valoriza pela tarifa homologada, leva ao valor do relatório e concilia. |
| `1-Entrega/rnf/rnf_<ref>.xlsx` | Resultado mensal. `Componentes` traz os blocos 1 e 2 na tarifa de aplicação; `Componentes_BE` o bloco 2 na base econômica; `Familias` as sete famílias nas quatro camadas; `Resumo` por origem, mercado e classe; `Conciliacao` o fechamento com o relatório; `Nao_tarifaveis` reativo e bandeira. |
| `src/RNF CCS ... .xlsb` | Relatório contábil da RNF: é a fonte de volume e valor. |
| `../BW/tbl/tarifas/tarifa_long.csv` | Tabela de tarifas homologadas, compartilhada com a fase I. |
| `../BW/1-Entrega/pxq/pxq_<ref>.xlsx` | P × Q do Faturamento do mesmo mês: fornece a divisão por posto horário. |

**Rotina mensal:**

1. Rodar o P × Q do Faturamento do mês (`SP/BW`).
2. Colocar o relatório contábil da RNF em `src/`.
3. Ajustar `REFERENCIA` no Passo 1 do notebook e executar.

## 3. Embasamento regulatório

As regras de estrutura tarifária (TUSD, TE, famílias e camadas), cliente cativo e livre, fonte
incentivada, APE e benefícios sociais são as mesmas do Faturamento e estão no
[README do BW](../BW/README.md), seção 3. Aqui estão só as diferenças.

### 3.1 O que é RNF e por que ela existe

**Normas:** CPC 47 (receita de contrato com cliente: reconhecida quando a obrigação de desempenho é
cumprida, isto é, quando a energia é entregue) e Manual de Contabilidade do Setor Elétrico (MCSE),
conta *Consumidores – Fornecimento não Faturado*.

**Impacto no cálculo:** não há fatura. O valor de referência é o que a contabilidade registrou; o
processo o decompõe por item tarifário.

### 3.2 Regras herdadas do Faturamento

| Regra | README do BW | Como se aplica na RNF |
|---|---|---|
| Estrutura da tarifa e camadas | 3.1 | Igual. Mesma tabela, mesmos itens e famílias. |
| Cativo e livre | 3.2 | Cativo: energia em TUSD e TE, demanda em TUSD. Livre: só TUSD. |
| Fonte incentivada e APE | 3.3 e 3.4 | O ZUBIL não identifica o cliente. O efeito vai para Fio A e Fio B pelo regime fio do livre do grupo A, como na 1ª passada do Faturamento. |
| Benefícios sociais | 3.5 | Mesma regra: a participação de cada faixa de consumo é estimada pelo valor registrado de cada linha. |
| Vigência | 3.6 | No mês do reajuste, o peso de cada vigência por subgrupo é estimado pelo valor registrado, como no Faturamento. |
| O que não é tarifa | 3.8 | Os valores do ZUBIL contêm PIS/COFINS; o tributo é retirado antes da abertura, pela alíquota informada no próprio relatório. Tributo, reativo e bandeira voltam na conciliação. |

## 4. Implementação

O notebook percorre as etapas abaixo. Cada passo tem, no próprio notebook, uma célula de texto com o
porquê.

**Preparação**

- **Leitura do relatório** (Passo 2). Quatro abas, localizadas pelo rótulo das linhas: o ZUBIL_146
  detalhado, o volume dos ajustes por classe, o valor dos ajustes e o comparativo do mês.
- **Linhas do SAP** (Passo 3). O ZUBIL é somado por classe, subclasse e categoria tarifária. Energia,
  demanda do cativo do grupo A e valor saem de colunas próprias; reativo e bandeira vão para os não
  tarifáveis. Categorias e classes são traduzidas por tabelas explícitas.
  O mercado segue a convenção do relatório: colunas de TE e "TUSD TE" são cativo, coluna "TUSD" é
  livre, qualquer que seja a categoria.
- **Ajustes manuais** (Passo 4). O valor de cada classe é volume × tarifa média do ZUBIL da classe,
  ajustado ao total do ajuste. Dentro da classe, o ajuste herda o mix das linhas do SAP. A
  iluminação pública recebe linha própria (B4a).
- **PIS/COFINS** (Passo 4b). Retirado do valor do relatório pela alíquota do próprio relatório (aba
  `RNF TOI SP`, células "Alíquota de COFINS" e "Alíquota de PIS").
- **Chave tarifária** (Passo 5) e **posto horário** (Passo 6). O posto vem do P × Q do Faturamento do
  mesmo mês: a participação de ponta, fora ponta e intermediário no consumo e na demanda faturados
  do mesmo subgrupo, modalidade, mercado e classe.

**Valorização**

- **Bloco 1** (Passos 7 e 8). Energia e demanda × tarifa homologada, nos 37 itens e nas quatro
  camadas. No mês do reajuste, o peso de cada vigência por subgrupo é estimado pelo valor registrado,
  como na tarifa proporcional do Faturamento.
- **Benefícios sociais** (Passo 8). O ZUBIL traz, por subclasse, o total de energia e de valor dos
  clientes de Tarifa Social e de Desconto Social, mas não a parte do consumo que cai em cada faixa. A
  tarifa muda com a faixa:

  | Benefício | Faixa 01 | Faixa 02 |
  |---|---|---|
  | Tarifa Social | até 80 kWh: gratuidade (tarifa zero) | acima de 80 kWh: tarifa `Baixa Renda` (sem CDE, PROINFA, CDE COVID e CDE GD) |
  | Desconto Social | até 120 kWh: chave `Desconto Social – faixa 01` (sem CDE, CDE COVID e CDE GD) | acima de 120 kWh: chave `Desconto Social – faixa 02` (tarifa residencial) |

  Cada linha é valorizada nas duas faixas (A = toda a energia na faixa 01; B = toda na faixa 02), e a
  participação da faixa 02 é a que reproduz o valor registrado: `s2 = (R − A) / (B − A)`, limitada a
  [0, 1]. O bloco 1 da linha é `(1 − s2) × A + s2 × B`, item a item. É a mesma aproximação do
  Faturamento (README do BW, 3.5).

  Em agosto/2026, nas linhas do SAP: Tarifa Social com 63.816 MWh e R$ 612,14/MWh registrados, contra
  R$ 623,09/MWh da tarifa `Baixa Renda` → 98,2% da energia na faixa 02; Desconto Social com 12.867 MWh
  e R$ 698,69/MWh, entre R$ 633,66 (faixa 01) e R$ 789,38 (faixa 02) → 41,8% na faixa 02. A
  participação quase integral da faixa 02 na Tarifa Social indica que a RNF registra a parte gratuita
  pela tarifa `Baixa Renda` (sessions, P16).
- **Bloco 2** (Passo 9). Livre do grupo A: a diferença vai para Fio A e Fio B, na composição da
  tarifa de demanda. Demais linhas: rateio proporcional por todos os itens.

**Saída e conferência**

- **Conferências** (Passo 10): bloco 2 igual ao valor do relatório linha a linha; TA = BE + CVA +
  FIN; famílias = componentes; volume aberto igual ao volume do relatório.
- **Conciliação** (Passo 11): aberto + PIS/COFINS + reativo + bandeira = valor contabilizado na conta
  1120110910 (`Conciliação contábil Novo`).
- **Gravação** (Passo 12): `rnf_<ref>.xlsx`, no desenho do `pxq_<ref>.xlsx`.

## 5. Resultados obtidos

Janeiro a agosto de 2026, um relatório contábil por mês. Setembro ainda não roda: depende do P × Q do
Faturamento de setembro (sessions, P14).

### 5.1 Do valor pela tarifa ao valor registrado

A conta de chegada é a mesma do Faturamento: `(a) valor registrado = (b) bloco 1 + (d) APE + (e) regime
fio + (f) 2ª passada + (g) resíduo`.

| Item | Unidade | jan | fev | mar | abr | mai | jun | jul | ago |
|---|---|---:|---:|---:|---:|---:|---:|---:|---:|
| (a) RNF contabilizada, líquida de PIS/COFINS (base tarifável) | R$ milhões | 1.512,2 | 1.400,0 | 1.566,3 | 1.415,1 | 1.418,6 | 1.446,2 | 1.501,0 | 1.711,3 |
| (b) Bloco 1 (tarifa cheia) | R$ milhões | 1.511,8 | 1.405,1 | 1.564,2 | 1.421,6 | 1.428,4 | 1.447,4 | 1.513,5 | 1.730,3 |
| (c) Diferença a explicar, (b) − (a) | R$ milhões | −0,4 | 5,2 | −2,1 | 6,5 | 9,8 | 1,2 | 12,6 | 18,9 |
| (d) Isenção APE | R$ milhões | — | — | — | — | — | — | — | — |
| (e) Regime fio, livre do grupo A (1ª passada) | R$ milhões | 14,8 | 6,8 | 14,2 | 4,3 | 4,9 | 12,0 | 5,8 | 4,9 |
| (f) 2ª passada, líquida | R$ milhões | −14,3 | −11,9 | −12,1 | −10,8 | −14,7 | −13,1 | −18,3 | −23,8 |
| (g) Resíduo não ajustável | R$ milhões | 0,0 | 0,0 | 0,0 | 0,0 | 0,0 | 0,0 | 0,0 | 0,0 |
| (h) 2ª passada, bruta | R$ milhões | 26,0 | 24,4 | 25,7 | 27,3 | 28,3 | 29,1 | 34,5 | 35,8 |

**O que é cada linha:**

- **(a)** Valor da RNF registrado no relatório (SAP ZUBIL_146 + ajustes manuais), sem reativo, bandeira
  e PIS/COFINS. É o que a abertura precisa reproduzir.
- **(b)** Energia e demanda não faturadas × tarifa homologada, com os benefícios sociais. No mês do
  reajuste (julho), cada subgrupo combina as duas vigências pelo peso estimado no Passo 8.
- **(c)** O quanto a tarifa sozinha não explica. Fica entre −0,1% e +1,1% de (a) em todos os meses.
- **(d)** Não se aplica: o ZUBIL não identifica o APE. O desconto, quando existe, fica em (e).
- **(e)** Diferença do livre do grupo A, levada para Fio A e Fio B: demanda não valorizada do livre e
  descontos de fonte incentivada e APE. Positiva: o valor registrado é maior que a energia pela tarifa.
- **(f)** Ajuste proporcional das demais linhas, somado com sinal. Entre −0,8% e −1,4% de (a).
- **(g)** Linhas com valor no relatório e bloco 1 zero. Nenhuma no período.
- **(h)** A mesma 2ª passada em valor absoluto, linha a linha: o quanto ela mexe, sem compensação entre
  sinais. Entre 1,6% e 2,3% de (a).

**Leitura:** a tarifa homologada explica a RNF contabilizada com diferença inferior a 1,5% em todos os
meses. O ajuste que o bloco 2 faz é pequeno, então a composição por item vem da tarifa, não do rateio.

### 5.2 Conciliação com o valor contabilizado

| Item | Unidade | jan | fev | mar | abr | mai | jun | jul | ago |
|---|---|---:|---:|---:|---:|---:|---:|---:|---:|
| (a) Valor líquido aberto por item | R$ milhões | 1.512,2 | 1.400,0 | 1.566,3 | 1.415,1 | 1.418,6 | 1.446,2 | 1.501,0 | 1.711,3 |
| (i) PIS/COFINS retirado | R$ milhões | 77,9 | 90,3 | 81,7 | 72,6 | 95,4 | 65,9 | 81,5 | 96,7 |
| (j) Reativo | R$ milhões | 7,4 | 7,0 | 7,2 | 7,6 | 7,0 | 6,6 | 6,7 | 7,6 |
| (k) Bandeira | R$ milhões | 0,0 | 0,0 | 0,0 | 0,0 | 31,1 | 31,8 | 32,1 | 33,8 |
| **(a) + (i) + (j) + (k)** | R$ milhões | **1.597,6** | **1.497,3** | **1.655,2** | **1.495,3** | **1.552,0** | **1.550,6** | **1.621,4** | **1.849,5** |
| Contabilizado (`Conciliação contábil Novo`, conta 1120110910) | R$ milhões | 1.597,6 | 1.497,3 | 1.655,2 | erro no relatório | 1.552,0 | 1.550,6 | 1.621,4 | 1.849,5 |
| `Comparativo Novo`, Não Faturado (informativo) | R$ milhões | 1.597,6 | 1.497,3 | 1.655,2 | 1.495,3 | 1.552,0 | 1.550,6 | 1.702,3 | 1.849,5 |

A diferença contra o contabilizado é menor que R$ 0,05 em todos os meses. Em abril a célula do
contabilizado está com erro de fórmula no relatório, e a conciliação é feita contra o `Comparativo
Novo`. Em julho, o `Comparativo Novo` fica R$ 81,0 milhões acima do contabilizado, e o próprio
relatório acusa a diferença no "Check" (sessions, P13).

### 5.3 Abertura por família e camada, agosto/2026

| Família | Cativo | Livre | Total | % do total |
|---|---:|---:|---:|---:|
| FIO A | 143,8 | 21,5 | 165,3 | 9,7% |
| FIO B | 384,6 | 32,7 | 417,2 | 24,4% |
| ENCARGOS | 357,3 | 178,1 | 535,5 | 31,3% |
| PERDAS | 86,0 | 18,4 | 104,4 | 6,1% |
| ENERGIA | 506,7 | 0,0 | 506,7 | 29,6% |
| OUTROS | −15,3 | −2,4 | −17,8 | −1,0% |
| **Total** | **1.463,1** | **248,3** | **1.711,3** | **100,0%** |

| Camada | R$ milhões | % do total |
|---|---:|---:|
| BE | 1.642,1 | 96,0% |
| CVA | 104,4 | 6,1% |
| FIN | −35,1 | −2,1% |

Bloco 2, tarifa de aplicação, sem tributos, R$ milhões.

### 5.4 Validações

Todas as conferências passaram nos oito meses: bloco 2 igual ao valor líquido em todas as linhas;
TA = BE + CVA + FIN; famílias = componentes; conciliação com o contabilizado (5.2). A conferência de
volume contra a aba `RNF_POR_CLASSE` é informativa: de janeiro a maio o pivô dessa aba não foi
atualizado no relatório (sessions, P12).
