# RNF — Abertura da Renda Não Faturada por item tarifário (Enel Distribuição São Paulo)

Este documento é o ponto de entrada da fase II do projeto P × Q. Ele explica o problema, o que foi
entregue, onde está cada peça e quais resultados foram obtidos. O detalhamento de cada transformação
está no [procedimento](1-Entrega/procedimentos/PROCEDIMENTO_RNF.md). As decisões e as pendências com
as áreas ficam em [`sessions.md`](sessions.md). As regras tarifárias comuns às duas fases estão no
[README do Faturamento](../BW/README.md), seção 3, e não são repetidas aqui.

Valores monetários em milhões de reais (MR$), energia em MWh e demanda em MW.

---

## 0. Contexto geral

A RNF (Renda Não Faturada) é a receita da energia já entregue e ainda não faturada no fechamento do
mês. Ela é reconhecida pelo regime de competência: provisionada no mês e estornada no mês seguinte,
quando a conta é emitida. O efeito na receita do mês é a provisão do mês menos o estorno da provisão
do mês anterior; no acumulado do ano, a provisão do mês menos a provisão de dezembro do ano anterior.

A RNF do sistema é calculada pelo SAP na transação ZUBIL_146, por lote, classe/subclasse e categoria
tarifária. A contabilidade a completa com ajustes calculados fora do SAP: TOI, faturas pendentes,
art. 323, iluminação pública, ajuste de carga e ajuste dos livres. O resultado é um valor total, sem
a composição tarifária.

Na fase I, o P × Q abriu a receita **faturada** nos 37 itens da tarifa. Sem a RNF, a receita do mês
por item fica incompleta: o que entra numa competência e é estornado na seguinte não tem abertura.

## 1. Objetivo

Produzir, todo mês, a **abertura da RNF por item tarifário, família e camada**, com três
características:

- **Fecha com a RNF registrada.** O valor aberto, somado ao PIS/COFINS, ao reativo e à bandeira, é
  igual ao total da RNF do mês.
- **Mesma régua do Faturamento.** Mesma tabela de tarifas homologadas, mesmos 37 itens, mesmas sete
  famílias e as quatro camadas (TA = BE + CVA + FIN). Faturamento e RNF somam item a item.
- **Mesma lógica de entrega.** Um relatório detalhado, com toda a abertura tarifária, e um
  consolidado no layout **Receita Faturada** da área (`receita_nao_faturada`), com provisão, estorno
  e saldo.

## 2. Estrutura do projeto

```
SP/RNF/
├── README.md                     este documento
├── sessions.md                   ata: decisões tomadas e pendências com as áreas
├── 1-Entrega/
│   ├── Abertura_RNF.ipynb        o cálculo, passo a passo (só o Passo 1 muda por mês)
│   ├── rnf/                      relatório detalhado: rnf_<aaaa-mm>.xlsx
│   ├── receita_nao_faturada/     consolidado no layout da área: receita_nao_faturada_<aaaa-mm>.xlsx
│   └── procedimentos/
│       └── PROCEDIMENTO_RNF.md   manual gerencial do notebook
├── 2-Checks/
│   └── zubil/                    ZUBIL_146 do mês só com as colunas usadas no cálculo
├── 3-Py/
│   └── 2-Checks/
│       └── extrai_zubil_colunas.py
├── src/                          entradas do mês (não versionadas)
└── 9-Historico/                  versões anteriores do processo
```

| Pasta / arquivo | Para que serve |
|---|---|
| [`1-Entrega/Abertura_RNF.ipynb`](1-Entrega/Abertura_RNF.ipynb) | O processo completo: lê o ZUBIL e os ajustes manuais, valoriza pela tarifa homologada, leva ao valor registrado, concilia e grava as duas saídas. |
| `1-Entrega/rnf/rnf_<ref>.xlsx` | Relatório detalhado. `Provisao`, `Estorno`, `Saldo` e `Saldo_YTD` por origem, mercado, classe, subgrupo e modalidade, com os 37 itens e as famílias; `Componentes`, `Componentes_BE` e `Familias` linha a linha; `Resumo`; `Nao_tarifaveis`; conciliações; `Layout_base` (fonte do estorno do mês seguinte). |
| `1-Entrega/receita_nao_faturada/receita_nao_faturada_<ref>.xlsx` | Consolidado no layout Receita Faturada (R$ mil): `Provisao`, `Estorno`, `Saldo`, `Saldo_YTD` e `Conferencia`. |
| [`1-Entrega/procedimentos/PROCEDIMENTO_RNF.md`](1-Entrega/procedimentos/PROCEDIMENTO_RNF.md) | O que cada passo do notebook faz e por quê. |
| [`3-Py/2-Checks/extrai_zubil_colunas.py`](3-Py/2-Checks/extrai_zubil_colunas.py) | Grava o ZUBIL do mês só com as colunas usadas (`2-Checks/zubil/`): `python 3-Py/2-Checks/extrai_zubil_colunas.py 2026-08`. |
| `src/` | Relatório mensal da RNF (`RNF CCS <mmaaaa> ... .xlsb`, fonte da aba do ZUBIL_146), base de ajustes manuais (`BASE_RNF_MANUAIS_<aaaamm>.xlsx`) e consolidação da margem (`SP_Consolidacao Margem_<aaaa>.<mm>.xlsx`, aba `Base RNF`; opcional: sem ela, a abertura sai sem a conciliação). |
| `../BW/tbl/tarifas/tarifa_long.csv` | Tabela de tarifas homologadas, compartilhada com o Faturamento. |
| `../BW/1-Entrega/pxq/pxq_<ref>.xlsx` | P × Q do Faturamento do mesmo mês: composição do faturado do grupo A (ponta e demanda). |
| `../BW/tbl/modelo/De-Para_Modelo_PyC.xlsx` e `../BW/3-Py/1-Entrega/gera_receita_faturada.py` | Modelo e de-para do layout Receita Faturada, os mesmos do Faturamento. |
| `9-Historico/` | Versão anterior, que partia do relatório contábil da RNF, e o contexto da especificação da base de ajustes. |

**Rotina mensal**

1. Rodar o P × Q do Faturamento do mês (`SP/BW`).
2. Colocar em `src/` o relatório da RNF do mês, a base de ajustes manuais e a consolidação da margem.
3. Ajustar `REFERENCIA` no Passo 1 do notebook e executar. Os meses rodam em ordem cronológica, porque
   o estorno de cada mês é lido da saída do mês anterior.

## 3. Embasamento regulatório

As regras de estrutura tarifária, cliente cativo e livre, fonte incentivada, APE e benefícios sociais
são as do [README do BW](../BW/README.md), seção 3. Aqui estão só as particularidades da RNF.

### 3.1 O que é RNF e por que ela existe

**Normas:** CPC 47 (receita de contrato com cliente: reconhecida quando a obrigação de desempenho é
cumprida, isto é, quando a energia é entregue) e Manual de Contabilidade do Setor Elétrico (MCSE),
conta *Consumidores – Fornecimento não Faturado*.

**Impacto no cálculo:** não há fatura. O valor de referência é o que o SAP e os ajustes manuais
registram; o processo o decompõe por item tarifário.

### 3.2 Regras herdadas do Faturamento

| Regra | README do BW | Como se aplica na RNF |
|---|---|---|
| Estrutura da tarifa e camadas | 3.1 | Igual. Mesma tabela, mesmos itens e famílias. |
| Cativo e livre | 3.2 | Cativo: energia em TUSD e TE, demanda em TUSD. Livre: só TUSD. |
| Fonte incentivada e APE | 3.3 e 3.4 | O ZUBIL não identifica o cliente. O desconto é estimado pelo valor registrado e aplicado sobre Fio A e Fio B. |
| Benefícios sociais | 3.5 | A participação de cada faixa de consumo é estimada pelo valor registrado de cada linha. |
| Vigência | 3.6 | Vale a vigência do último dia do mês. |
| O que não é tarifa | 3.8 | PIS/COFINS, reativo e bandeira ficam fora da abertura e voltam na conciliação e no layout. |

### 3.3 O que o ZUBIL informa e o que o valor contém

O ZUBIL informa a energia sem posto horário e a demanda de uma única coluna. O valor registrado,
porém, contém a energia de ponta, a demanda nos dois postos e os descontos de fonte incentivada e
APE, como no faturamento do mês. Por isso:

- a demanda informada entra no bloco 1, no cativo e no livre do grupo A;
- no livre, o SAP registra a receita de demanda na coluna "TUSD TE" e a de energia na coluna "TUSD":
  cada linha livre gera uma linha de energia e uma de demanda;
- o que a quantidade informada não explica (energia de ponta, demanda não informada no Azul) é aberto
  pela composição do faturado do mesmo mês.

## 4. Implementação

O notebook percorre as etapas abaixo. O porquê de cada uma está no
[procedimento](1-Entrega/procedimentos/PROCEDIMENTO_RNF.md), na seção de mesmo número.

**Preparação**

- **ZUBIL_146** (Passos 2 e 3). Linhas somadas por classe, subclasse e categoria tarifária, com
  tabelas explícitas de categoria e classe. Energia, demanda e valor saem de colunas próprias; reativo
  e bandeira vão para os não tarifáveis. Cada linha livre é separada em energia e demanda.
- **Ajustes manuais** (Passo 4). Lidos da base de ajustes, linha a linha, com os atributos da base.
  O ajuste de carga e o ajuste dos livres, sem atributos, herdam o mix do SAP do mesmo mercado.
- **PIS/COFINS** (Passo 5). Retirado do valor do ZUBIL pelas alíquotas da base de ajustes.
- **Chave tarifária e posto** (Passos 6 e 7). Mesma chave de 8 atributos da tabela de tarifas;
  energia e demanda do grupo A na tarifa fora ponta.

**Valorização**

- **Bloco 1** (Passo 8). Energia e demanda × tarifa homologada, nos 37 itens e nas quatro camadas,
  com energia e demanda separadas. Benefícios sociais pela estimativa das faixas.
- **Bloco 2** (Passo 9). A diferença para o valor registrado é devolvida aos itens a que pertence:

  | Linha | Regra |
  |---|---|
  | Demanda do livre | Aplicação do desconto de fonte incentivada e APE sobre Fio A e Fio B |
  | Energia do grupo A (cativo e livre) | Composição do faturado do mesmo mês, só nos postos que o bloco 1 não valoriza (energia de ponta e demanda) |
  | Demais linhas | 2ª passada: rateio proporcional pelos itens |

**Saída e conferência**

- **Conferências** (Passo 10): bloco 2 igual ao valor líquido linha a linha; TA = BE + CVA + FIN;
  famílias = componentes; ZUBIL igual ao cluster `Sistema` da base de ajustes.
- **Conciliação** (Passo 11): aberto + PIS/COFINS + reativo + bandeira contra a aba `Base RNF` da
  consolidação da margem, bloco a bloco.
- **Layout** (Passo 12): provisão convertida no layout Receita Faturada, com o modelo e o de-para do
  Faturamento.
- **Gravação** (Passos 13 e 14): relatório detalhado e consolidado, com provisão, estorno e saldo.

## 5. Resultados obtidos

Agosto e setembro de 2026, rodados em sequência (setembro estorna agosto).

### 5.1 Do valor pela tarifa ao valor registrado

```
(a) RNF registrada = (b) bloco 1 + (d) desconto na demanda do livre + (e) composição do faturado
                     + (f) 2ª passada + (g) resíduo
```

| Item | ago | set |
|---|---:|---:|
| (a) RNF registrada, líquida de PIS/COFINS (base tarifável) | 1.716,31 | 1.661,72 |
| (b) Bloco 1: quantidade × tarifa cheia | 1.654,82 | 1.599,42 |
| (c) Diferença a explicar, (b) − (a) | −61,50 | −62,29 |
| (d) Aplicação do desconto na demanda do livre | −10,35 | −8,12 |
| (e) Composição do faturado: energia de ponta e demanda do grupo A | 72,04 | 70,15 |
| (e1) dela, livre | 55,44 | 52,97 |
| (e2) dela, cativo | 16,60 | 17,18 |
| (f) 2ª passada, líquida | −0,19 | 0,27 |
| (g) Resíduo sem item tarifário | 0,00 | 0,00 |
| (h) 2ª passada, bruta | 8,39 | 8,62 |

MR$.

**O que é cada linha**

- **(a)** Valor da RNF do ZUBIL e dos ajustes manuais, sem PIS/COFINS, reativo e bandeira. É o que a
  abertura precisa reproduzir.
- **(b)** Energia e demanda não faturadas × tarifa homologada cheia, com os benefícios sociais.
- **(c)** O quanto a quantidade informada não explica. Negativa: o valor registrado fica acima do
  bloco 1, porque contém a energia de ponta e demanda que a quantidade não traz.
- **(d)** A demanda do livre valorizada pela tarifa cheia menos o valor registrado: o desconto de
  fonte incentivada e APE.
- **(e)** Energia de ponta (relevante no Verde) e demanda não informada (relevante no Azul), abertas
  pela composição do faturado do mês.
- **(f) e (h)** Rateio proporcional das demais linhas, com sinal e em módulo. Vem dos ajustes manuais,
  valorizados pela tarifa média de cada classe; da iluminação pública (B4a e B4b não separadas); e da
  Branca (sem posto).

**Leitura:** depois de tratados a demanda e os postos horários, a 2ª passada líquida é praticamente
zero nos dois meses. A composição por item vem da tarifa e do faturado, não do rateio.

### 5.2 Conciliação com a `Base RNF`

| Bloco | ago `Base RNF` | ago PxQ RNF | ago dif. | set `Base RNF` | set PxQ RNF | set dif. |
|---|---:|---:|---:|---:|---:|---:|
| Consumo Cativo | 1.065,51 | 1.065,51 | 0,00 | 1.003,62 | 1.003,62 | 0,00 |
| TOI | 72,07 | 72,31 | 0,24 | 73,64 | 73,84 | 0,20 |
| Pendentes | 293,38 | 293,29 | −0,10 | 305,32 | 305,29 | −0,03 |
| Artigo 113 | 95,95 | 95,86 | −0,08 | 106,96 | 106,85 | −0,10 |
| Carga | 26,45 | 26,45 | 0,00 | 21,58 | 21,58 | 0,00 |
| Bandeira | 33,85 | 33,85 | 0,00 | 32,91 | 32,91 | 0,00 |
| Total Livres | 262,29 | 262,29 | 0,00 | 249,34 | 249,34 | 0,00 |
| **Total** | **1.849,50** | **1.849,56** | **0,06** | **1.793,38** | **1.793,44** | **0,07** |

MR$, com PIS/COFINS. Composição do total do PxQ RNF:

| Item | ago | set |
|---|---:|---:|
| Valor aberto por item (líquido) | 1.716,31 | 1.661,72 |
| PIS/COFINS | 91,82 | 91,73 |
| Reativo | 7,58 | 7,09 |
| Bandeira | 33,85 | 32,91 |
| **Total** | **1.849,56** | **1.793,44** |

A diferença que sobra vem da distribuição por classe do TOI, das faturas pendentes e do art. 323 na
base de ajustes, diferente da usada na `Base RNF`. O SAP, a carga, a bandeira e os livres fecham
exatamente.

### 5.3 Abertura por família e camada

| Família | ago cativo | ago livre | ago total | set cativo | set livre | set total |
|---|---:|---:|---:|---:|---:|---:|
| FIO A | 129,33 | 34,37 | 163,70 | 125,36 | 30,97 | 156,32 |
| FIO B | 382,40 | 30,86 | 413,27 | 371,98 | 28,51 | 400,49 |
| ENCARGOS | 352,65 | 185,92 | 538,56 | 343,91 | 175,89 | 519,81 |
| PERDAS | 88,73 | 19,32 | 108,05 | 86,45 | 18,29 | 104,74 |
| ENERGIA | 511,09 | 0,00 | 511,09 | 498,12 | 0,00 | 498,12 |
| OUTROS | −15,18 | −3,18 | −18,36 | −14,78 | −2,98 | −17,76 |
| **Total** | **1.449,03** | **267,28** | **1.716,31** | **1.411,04** | **250,68** | **1.661,72** |

| Camada | ago | set |
|---|---:|---:|
| BE | 1.647,25 | 1.594,86 |
| CVA | 105,08 | 101,85 |
| FIN | −36,02 | −34,99 |

Bloco 2, tarifa de aplicação, líquido de PIS/COFINS, MR$.

### 5.4 Efeito na receita: provisão, estorno e saldo (setembro)

| Item | MR$ |
|---|---:|
| Provisão de setembro | 1.793,44 |
| Estorno da provisão de agosto | −1.849,56 |
| **Saldo: efeito da RNF na receita de setembro** | **−56,12** |

| Família | Saldo (MR$) |
|---|---:|
| FIO A | −7,38 |
| FIO B | −12,78 |
| ENCARGOS | −18,76 |
| PERDAS | −3,31 |
| ENERGIA | −12,97 |
| OUTROS | 0,60 |
| Não tarifáveis (PIS/COFINS, reativo, bandeira) | −1,53 |
| **Total** | **−56,12** |

O saldo é −38,75 no cativo e −17,37 no livre. O `Saldo_YTD` (contra dezembro de 2025) será gerado
quando a RNF de dezembro de 2025 for processada.

### 5.5 Validações

Todas as conferências internas passaram nos dois meses: bloco 2 igual ao valor líquido em todas as
linhas; TA = BE + CVA + FIN; famílias = componentes. O layout `receita_nao_faturada` grava o total da
RNF sem valor fora do layout, e o estorno de setembro é igual à provisão de agosto com sinal trocado.
