# P × Q — Abertura do faturamento de energia por item tarifário (Enel Distribuição São Paulo)

Este documento é o ponto de entrada do projeto. Ele explica o problema, o que foi entregue, onde
está cada peça e quais resultados foram obtidos. O detalhamento de cada transformação está no
[procedimento](1-Entrega/procedimentos/PROCEDIMENTO_PXQ.md). As decisões e as pendências com as
áreas ficam registradas em [`sessions.md`](sessions.md).

---

## 0. Contexto geral

A receita de fornecimento de energia é contabilizada no SAP pelo valor faturado de cada
lançamento. O valor total é conhecido, mas não a sua composição: quanto daquele valor remunera a
rede da própria distribuidora (Parcela B), quanto repassa custos de transmissão, quanto é encargo
setorial, quanto é perda e quanto é compra de energia.

Essa composição é necessária para o controle da margem bruta e para os relatórios gerenciais e
regulatórios. Hoje ela é obtida pelo simulador tarifário (Sparta), que aplica as tarifas a um
mercado simulado. O simulador trabalha com tarifas arredondadas, não reproduz a tarifa
proporcional dos meses de reajuste e não fecha com o faturamento contabilizado.

O P × Q (preço × quantidade) parte do caminho inverso. Ele começa pelo que foi efetivamente
faturado, extraído do SAP BW, e multiplica cada quantidade faturada pela tarifa homologada pela
ANEEL. O resultado é aberto nos 37 itens da tarifa. Sobre ele são aplicados os descontos previstos
na regulação e, por fim, o ajuste de fechamento (2ª passada), para que o total feche exatamente com
a receita contabilizada.

## 1. Objetivo

Produzir, todo mês, a **abertura da receita de fornecimento por item tarifário e por família**,
com três características:

- **Fecha com a receita contabilizada.** O total aberto é igual, ao centavo, à receita do BW, que
  é a mesma base do informe de gestão.
- **Respeita a regra tarifária.** Cada desconto é aplicado sobre os itens em que a regulação o faz
  incidir. Ele não é espalhado de forma proporcional por todos os itens.
- **Separa as quatro camadas da tarifa.** Base econômica (BE), CVA e componentes financeiros
  (FIN), que somam a tarifa de aplicação (TA).

A saída final é entregue no layout **Receita Faturada** definido pela área de Planejamento
(`tbl/modelo/De-Para_Modelo_PyC.xlsx`).

O método foi construído para ser replicado. Outras distribuidoras do grupo (CE, RJ) e outras
fontes de receita, como a Renda Não Faturada (RNF) e a CNR, mudam os arquivos de entrada, não a
lógica do cálculo.

## 2. Estrutura do projeto

```
SP/BW/
├── README.md                     este documento
├── sessions.md                   ata: decisões tomadas e pendências com as áreas
├── 1-Entrega/
│   ├── Abertura_PxQ.ipynb        o cálculo, passo a passo (só o Passo 0 muda por mês)
│   ├── pxq/                      saída mensal do notebook: pxq_<aaaa-mm>.xlsx
│   ├── pxq_regulacao/            a entrega aberta em BE, CVA e FIN, para a Regulação
│   ├── receita_faturada/         a mesma saída no layout Receita Faturada
│   └── procedimentos/
│       └── PROCEDIMENTO_PXQ.md   manual gerencial do notebook
├── 2-Checks/                     saídas das conferências
│   ├── resultados/               um arquivo de conferência por mês
│   ├── receita_faturada_MR/      layout Receita Faturada do MR e check PxQ − MR, por mês
│   ├── receita_faturada_Bl1_vs_Bl2/  tarifa homologada × tarifa realizada e descontos por origem
│   ├── tabelas_readme/           tabelas da seção 5 deste documento (.md e .xlsx)
│   └── consolidado_*.xlsx        consolidados mês a mês
├── 3-Py/                         todos os scripts Python, separados pela pasta que alimentam
│   ├── 1-Entrega/
│   │   ├── prepara_base_bw.py        csv do BW -> parquet limpo, conferido e compactado
│   │   ├── extrai_pcat.py            gera a tabela de tarifas a partir dos PCATs
│   │   ├── gera_calendario_leitura.py  converte o calendário de leitura em tabela
│   │   └── gera_receita_faturada.py  converte o pxq no layout Receita Faturada
│   ├── 2-Checks/
│   │   ├── le_mr.py                  lê o Mercado Receita do mês
│   │   ├── check_mr_bloco2.py        check por opção de compra e subgrupo
│   │   ├── check_mr_camadas.py       check por família e camada
│   │   ├── consolida_check_mr_bloco2.py  consolida o check por subgrupo
│   │   ├── gera_receita_faturada_mr.py   layout Receita Faturada do MR e check célula a célula
│   │   └── gera_receita_faturada_bl1_vs_bl2.py  tarifa homologada × tarifa realizada (descontos)
│   └── readme/
│       └── tabelas_readme.ipynb      gera as tabelas de resultados da seção 5
├── tbl/
│   ├── tarifas/tarifa_long.csv   tarifas homologadas, todas as vigências e camadas
│   ├── modelo/De-Para_Modelo_PyC.xlsx  modelo da área para o layout Receita Faturada
│   ├── modelo/De-Para_Modelo_PyC_sugerido.xlsx  o mesmo, com a linha GERAÇÃO / E (P27)
│   ├── calendario/calendario_leitura.csv  data de leitura e ciclo de cada lote, por mês
│   └── calendario/de_para_lote.csv  lote do BW -> lote do calendário de leitura
└── src/                          arquivos de origem (não editados pelo processo)
    ├── csv_original/             extração mensal do BW, csv compactado (MM-AAAA_ENEL SP_RECEITA.zip)
    └── parquet/                  a mesma extração em parquet limpo, lida pelo notebook
```

| Pasta / arquivo | Para que serve |
|---|---|
| [`1-Entrega/Abertura_PxQ.ipynb`](1-Entrega/Abertura_PxQ.ipynb) | O processo completo. Lê o BW e a planilha de geradoras, suprimentos e uso da distribuição (SD), valoriza pela tarifa homologada, aplica os descontos regulatórios e o ajuste de fechamento (2ª passada) e grava `pxq_<ref>.xlsx`. |
| `1-Entrega/pxq/pxq_<ref>.xlsx` | Resultado mensal. A aba `Leia-me` descreve cada aba e os scripts que a usam. A aba `Componentes` traz a tarifa de aplicação e `Componentes_BE` a base econômica, ambas abertas nos 37 itens. `Familias` agrupa os itens nas sete famílias. `Nao_tarifaveis_linha` traz os lançamentos que não entram no cálculo, um a um. `Lotes_sem_calendario` lista os lotes do BW com faturamento normal que não estão no de-para do calendário de leitura. |
| `1-Entrega/pxq_regulacao/PxQ_regulacao_<ref>.xlsx` | Para a área de Regulação: as mesmas linhas da aba `Componentes`, numa aba só, com a identificação do BW, as quantidades (`Energia com TUSD`, `Energia com TE` e `Demanda`), a receita e a entrega (bloco 2) item a item na BE, na CVA e na FIN. Sem bloco 1, sem colunas de ajuste e sem painéis congelados. |
| `1-Entrega/receita_faturada/receita_faturada_<ref>.xlsx` | O pxq convertido no layout Receita Faturada, com a aba `Conferencia`, que concilia o total com a receita do BW, e a aba `Lotes sem calendário`. |
| [`1-Entrega/procedimentos/PROCEDIMENTO_PXQ.md`](1-Entrega/procedimentos/PROCEDIMENTO_PXQ.md) | Cada transformação do notebook e dos scripts, e o motivo dela. |
| `2-Checks/resultados/` | Conferências mensais contra o Mercado Receita e a conferência contra o informe de gestão. |
| `2-Checks/receita_faturada_MR/` | Um arquivo por mês com as abas `PxQ`, `MR`, `check` (fórmulas PxQ − MR) e `Conciliacao`, por opção de compra × unidade, e o consolidado com uma linha por mês. |
| `2-Checks/receita_faturada_Bl1_vs_Bl2/` | Um arquivo por mês com o layout pela tarifa homologada (bloco 1) e pela tarifa realizada (bloco 2), a diferença por fórmula e a diferença separada em Tarifa Social – DMR, fonte incentivada e 2ª passada (isenção de encargos do APE e SCEE - Baixa Renda só como memória na conciliação), e o consolidado com uma linha por mês (seção 5.4). |
| `2-Checks/tabelas_readme/` | `tabelas_readme.md` e `tabelas_readme.xlsx`: as tabelas de resultados da seção 5 (5.1, 5.3 e 5.4), geradas por `3-Py/readme/tabelas_readme.ipynb`. No Excel, uma aba por tabela com a explicação e a seção correspondente. Fora do Git. |
| `2-Checks/consolidado_*.xlsx` | Checks por família e camada e por opção de compra e subgrupo, consolidados mês a mês (seção 5.3). |
| [`3-Py/1-Entrega/prepara_base_bw.py`](3-Py/1-Entrega/prepara_base_bw.py) | Converte a extração mensal do BW (csv de mais de 200 MB) em parquet limpo em `src/parquet/`, confere o total contra o próprio csv, compacta o csv em `src/csv_original/` e o apaga. Roda antes do notebook. |
| [`3-Py/1-Entrega/extrai_pcat.py`](3-Py/1-Entrega/extrai_pcat.py) | Lê os PCATs (planilhas de cálculo tarifário da ANEEL) e grava `tbl/tarifas/tarifa_long.csv`. Roda uma vez a cada nova vigência tarifária. |
| [`3-Py/1-Entrega/gera_calendario_leitura.py`](3-Py/1-Entrega/gera_calendario_leitura.py) | Lê o calendário de leitura (`src/calendario_leitura_<ano>.xlsx`, uma aba por mês) e grava `tbl/calendario/calendario_leitura.csv`: `python 3-Py/1-Entrega/gera_calendario_leitura.py 2026`. Carrega só as abas presentes e substitui apenas esses meses na tabela. |
| [`3-Py/1-Entrega/gera_receita_faturada.py`](3-Py/1-Entrega/gera_receita_faturada.py) | Converte o pxq no layout Receita Faturada: `python 3-Py/1-Entrega/gera_receita_faturada.py 2026-06` ou `--todos`. |
| [`3-Py/2-Checks/`](3-Py/2-Checks/) | `check_mr_bloco2.py` compara o mês por opção de compra e subgrupo. `check_mr_camadas.py` compara por família e camada, mês a mês. `consolida_check_mr_bloco2.py` consolida vários meses. `le_mr.py` lê o Mercado Receita (também usado pelos checks do SAMP). `gera_receita_faturada_mr.py` monta o Mercado Receita no layout Receita Faturada e gera o check célula a célula: `python 3-Py/2-Checks/gera_receita_faturada_mr.py 2026-05` ou `--todos`. `gera_receita_faturada_bl1_vs_bl2.py` compara a tarifa homologada com a realizada e separa os descontos: `python 3-Py/2-Checks/gera_receita_faturada_bl1_vs_bl2.py --todos`. |
| [`3-Py/readme/tabelas_readme.ipynb`](3-Py/readme/tabelas_readme.ipynb) | Gera as tabelas da seção 5 deste documento a partir do pxq e dos checks. O período de cada tabela e a opção de rodar os checks antes ficam no Passo 0. O texto de leitura continua manual. |
| `tbl/` | Tabelas do processo: tarifas homologadas, calendário de leitura e o modelo do layout Receita Faturada. |
| `tbl/calendario/` | `calendario_leitura.csv`: uma linha por mês, grupo e lote, com a data de leitura, os dias de consumo e o início e fim do ciclo. `de_para_lote.csv`: o grupo e o nº do lote do calendário para cada lote do BW (campo `Conjunto de Contrato`), validado pela área do BW. Ambos em csv para facilitar a passagem para a TI. |
| `tbl/modelo/De-Para_Modelo_PyC.xlsx` | Modelo da área para o layout Receita Faturada: a aba "Receita Faturada" define as colunas e a grade de linhas; a aba "Instruções", o de-para dos itens tarifários para os grupos do layout. Sem ele a Receita Faturada não é gerada. Não tem dados da Enel e é versionado no Git. |
| `tbl/modelo/De-Para_Modelo_PyC_sugerido.xlsx` | O modelo da área com a linha GERAÇÃO / E (energia das geradoras) em FAT, REFAT e CNR, sugerida à área em 29/10/2026 (sessions, P27). É o modelo que o conversor usa; o arquivo da área fica intacto. Para voltar ao original, basta trocar `ARQ_MODELO` em `gera_receita_faturada.py`. Versionado no Git. |
| `src/` | Fontes externas: a extração de receita do BW, a planilha do SD, os PCATs (`src/pcats/`), o calendário de leitura (`calendario_leitura_<ano>.xlsx`) e o Mercado Receita do período tarifário. |

Os scripts localizam os arquivos a partir da pasta `SP/BW` (variável `PXQ_BASE`), e não da pasta onde
estão. Por isso os comandos acima rodam a partir de `SP/BW`.

**Rotina mensal:**

1. Deixar o csv do mês em `src/csv_original/` e rodar `python 3-Py/1-Entrega/prepara_base_bw.py`
   (converte em parquet, confere o total, compacta e apaga o csv).
2. Definir `REFERENCIA` no Passo 0 do notebook e executar.
3. Rodar `python 3-Py/1-Entrega/gera_receita_faturada.py <ref>`.
4. Quando chegar um calendário de leitura novo ou revisado: salvar em `src/calendario_leitura_<ano>.xlsx`
   e rodar `python 3-Py/1-Entrega/gera_calendario_leitura.py <ano>`.
5. Uma vez por ano, quando sai o Mercado Receita do reajuste, rodar os checks de `3-Py/2-Checks/`,
   inclusive o `gera_receita_faturada_mr.py`, depois do passo 3.
6. Para atualizar a seção 5 deste documento, rodar `3-Py/readme/tabelas_readme.ipynb` e copiar as
   tabelas de `2-Checks/tabelas_readme/tabelas_readme.md`.

### Versões

O desenvolvimento segue no branch `main`. As versões fechadas ficam marcadas por tag no Git:

| Tag | Versão |
|---|---|
| `vf-base1-antes-acrescimo-novas-colunas` | Versão final para a extração do BW em Excel (`Receita de Fornecimento Jan a Ago-26 com subclasse.xlsx`), de janeiro a agosto de 2026, antes da base em parquet com as colunas novas: data de leitura e de faturamento, indicador e referência original do refaturamento e identificação do TOI. |
| `VERSAO-FINAL` | Versão final com a base do BW em csv/parquet (lote, referência original e indicador de refaturamento), tipos de faturamento FAT/REFAT/CNR, Tarifa Social no bloco 1 e tarifa proporcional pelo ciclo de leitura do lote. Jan a set/2026. |
| `VF-com-PxQ-Regulacao` | Versão com a saída para a Regulação (`PxQ_regulacao_<ref>.xlsx`: entrega aberta em BE, CVA e FIN), as colunas `Energia com TUSD` e `Energia com TE`, a TE de suprimento da Ceris completada nas vigências 2024 e 2025 e a TUSD do Consumo Próprio separada da 2ª passada na documentação. Jan a set/2026. |
| `BW-VF-diferencas-solucionadas` | Versão com as diferenças de energia contra o informe de gestão resolvidas: energia do suprimento (Ceris) contada uma vez, energia compensada fora ponta (ZQC223/ZQC224) tarifável e linha GERAÇÃO / E no modelo sugerido do layout. TE de suprimento da Ceris completada nas vigências 2024 e 2025. Notebook de tabelas da seção 5. Jan a set/2026. |

Para consultar uma versão: no GitHub Desktop, aba **History**, o commit marcado com a tag; no
GitHub, **Tags**, onde o código daquela versão pode ser baixado.

## 3. Embasamento regulatório

Esta seção reúne as normas que determinam como cada parte da receita deve ser calculada. Cada
regra abaixo tem um passo correspondente no notebook (seção 4).

### 3.1 Estrutura da tarifa: TUSD, TE, funções de custo e camadas

**Norma:** PRORET Submódulo 7.1 (Procedimentos de Regulação Tarifária da ANEEL: estrutura
tarifária das distribuidoras), aplicado anualmente no PCAT de cada distribuidora.

A tarifa de fornecimento tem duas parcelas:

- **TUSD** (Tarifa de Uso do Sistema de Distribuição) remunera o uso da rede.
- **TE** (Tarifa de Energia) remunera a energia consumida.

Cada parcela se divide em funções de custo, que o projeto agrupa em famílias:

| Família | O que remunera |
|---|---|
| FIO A | Uso de ativos de terceiros: Rede Básica, conexão, transporte de Itaipu, uso de outras distribuidoras |
| FIO B | Uso dos ativos da própria distribuidora (Parcela B) |
| ENCARGOS | Encargos setoriais: CDE, PROINFA, P&D, TFSEE, ONS, ESS/ERR e outros |
| PERDAS | Perdas técnicas, não técnicas e da Rede Básica |
| ENERGIA | Compra de energia e Angra |
| RI | Receitas irrecuperáveis (zero em SP; mantida para outras distribuidoras) |
| OUTROS | Subsídios e benefícios embutidos na tarifa |

Cada item tem quatro camadas:

```
TA (tarifa de aplicação) = BE (base econômica) + CVA (Conta de Compensação de Variação de Valores de Itens da Parcela A) + FIN (componentes financeiros)
```

**Impacto no cálculo:** a quantidade faturada é multiplicada item a item e camada a camada. Como
todas as operações posteriores são multiplicações lineares, a identidade TA = BE + CVA + FIN se
mantém do início ao fim.

### 3.2 Cliente cativo e cliente livre

**Norma:** regras do Ambiente de Contratação Livre. O consumidor livre ou especial compra a
energia no mercado e se conecta em alta tensão.

**Impacto no cálculo:**

- O cliente cativo paga TUSD e TE à distribuidora.
- O cliente livre paga só a TUSD: a energia é comprada de outro fornecedor.
- Na base de SP, todos os clientes livres são do grupo A (alta tensão).

### 3.3 Desconto de fonte incentivada

**Normas:** REN ANEEL 1.031/2022 (que consolidou a REN 77/2004) e PRORET Submódulo 7.1, §38.

- A resolução garante desconto de 50% a 100% na TUSD ao consumidor que compra energia de fonte
  incentivada (PCH, solar, eólica, biomassa, cogeração qualificada).
- O desconto nunca incide sobre a TE.
- O PRORET determina que o percentual incide sobre a função de custo **TUSD Transporte**, que é
  formada exclusivamente por **Fio A + Fio B**. Encargos e perdas não recebem desconto.

**Impacto no cálculo:** para o cliente livre do grupo A, a diferença entre a conta pela tarifa
cheia e o faturado é absorvida apenas por Fio A e Fio B.

### 3.4 Autoprodutor por equiparação (APE)

**Normas:** Lei 11.488/2007, art. 26, que equipara o consumidor a autoprodutor e o isenta de
encargos setoriais como CDE e PROINFA sobre a energia autoproduzida. Na fatura, o benefício
aparece como "TAR.REDUZIDA RES.ANEEL 166/05".

**Impacto no cálculo:**

- O BW não identifica o APE. A isenção aparece em linhas próprias, sem quantidade: "TUSD -
  Desconto Consumo" e o estorno "Complemento Res. 166/05".
- O processo usa essas linhas para abater **somente os encargos** da TUSD de energia dos clientes
  livres do grupo A.

### 3.5 Benefícios sociais: Tarifa Social e Desconto Social

**Normas:**

- Lei 12.212/2010 (Tarifa Social de Energia Elétrica).
- Lei 15.235/2025 (Desconto Social).
- REN ANEEL 1.000/2021, art. 179, § 1º (Tarifa Social) e § 4º (Desconto Social), na redação dada pela
  REN ANEEL 1.147/2025.
- PRORET Submódulo 7.1, §18 e §23 (itens isentos).

| Benefício | Regra |
|---|---|
| Tarifa Social (Baixa Renda) | Até 80 kWh/mês, isenção total. Acima disso, tarifa B1 baixa renda. Isenção de CDE, PROINFA, CDE COVID e CDE GD. |
| Desconto Social | Até 120 kWh/mês, tarifa residencial sem CDE, CDE COVID e CDE GD. Acima disso, tarifa residencial comum. |

**Reembolso pela CDE (PRORET Submódulo 5.2, §81, §91, §129 e §130):**
- Só a Tarifa Social é reembolsada à distribuidora. O valor é a diferença entre a receita pela tarifa
  homologada B1 subclasse Baixa Renda e a receita com a redução, sem tributos e sem bandeiras.
- A régua é a tarifa da própria subclasse Baixa Renda, e não a residencial comum.
- A isenção de encargos (Baixa Renda e Desconto Social) não é reembolsada: o consumo isento sai do
  mercado de rateio da cota da CDE, e o custo vai para os demais consumidores pela tarifa.

**Impacto no cálculo:**

- Os itens isentos são zerados para essas subclasses.
- O BW não informa quanto do consumo de cada cliente cai em cada faixa. A participação de cada
  faixa é estimada pela própria receita faturada.
- No bloco 1, a Tarifa Social fica pela tarifa B1 subclasse Baixa Renda sem a redução; a faixa
  gratuita estimada é a primeira etapa do bloco 2 e é a estimativa do valor reembolsado (seção 5.4).
  O Desconto Social mantém as faixas no bloco 1.

### 3.6 Vigência tarifária e tarifa proporcional

**Norma:** REN ANEEL 1.000/2021, art. 173, § 2º. Quando a tarifa muda no decorrer do ciclo de
faturamento, a distribuidora aplica a tarifa proporcional:

```
TP = Σ (Ti × Pi) / D
```

em que `Ti` é a tarifa em vigor no período `i`, `Pi` o número de dias em que ela vigorou e `D` o
número de dias de fornecimento entre duas leituras consecutivas. Em SP, o reajuste anual ocorre em
4 de julho.

**Impacto no cálculo:** o peso de cada vigência vem do ciclo de leitura do lote. O BW traz o lote
do cliente (`Conjunto de Contrato`); o calendário de leitura dá a data de leitura e os dias de
consumo de cada lote (`tbl/calendario/`). Só os ciclos que contêm a data do reajuste misturam as duas
vigências; nos demais meses o cálculo não muda. Em 2026, isso afeta as leituras de julho e agosto: a
Média Tensão, lida no dia 1º, fatura julho inteiro pela tarifa antiga; a Baixa Tensão de julho tem
de 1 a 29 dias na tarifa antiga, conforme o lote. Refaturamento (REFAT) e faturamento manual (CNR)
ficam com a tarifa do mês: o REFAT exigiria o calendário e a tarifa do mês original, que pode ser de
anos atrás, para um efeito pequeno; o CNR não tem ciclo de leitura.

### 3.7 ERCAP: Encargo de Potência para Reserva de Capacidade

**Normas:** Decreto 10.707/2021, art. 8º, §3º, e REN ANEEL 1.103/2024.

**Impacto no cálculo:**

- O ERCAP não tem preço nem quantidade próprios. É o rateio, feito pela CCEE (Câmara de
  Comercialização de Energia Elétrica), do custo da reserva de capacidade, proporcional à demanda
  máxima horária de cada consumidor.
- No cliente cativo, o ERCAP já está embutido na tarifa homologada (item `TE_ESSERR`) e entra
  normalmente no P × Q.
- A cobrança direta aos clientes livres do grupo A fica **fora** do P × Q. Esse valor é obtido do
  Mercado Receita do Sparta.

### 3.8 O que não é tarifa

Tributos (ICMS, PIS/COFINS), bandeiras tarifárias, energia reativa, ultrapassagem de demanda e
substituição tributária não têm tarifa homologada por item. Esses lançamentos ficam fora do cálculo
P × Q. Tributos e bandeiras voltam no layout Receita Faturada, nas colunas próprias.

## 4. Implementação

O notebook percorre as etapas abaixo. O porquê de cada uma está no
[procedimento](1-Entrega/procedimentos/PROCEDIMENTO_PXQ.md), na seção de mesmo nome.

**Preparação da base**

- **Leitura do BW.** A extração de receita de fornecimento traz uma linha por categoria tarifária,
  código estatístico, operação, classe e mercado. A receita de cada linha é a referência que o
  resultado final precisa reproduzir.
- **Tipo de faturamento.** Cada lançamento é classificado como CNR (TOI, operação "Faturamento
  Manual de Energia"), REFAT (indicador de refaturamento X, estorno e conta nova) ou FAT (os demais).
  O tipo segue pelo cálculo e preenche a coluna `TipoFaturamento` do layout.
- **Separação do que é tarifa** (3.8). Só os 30 códigos estatísticos de consumo e demanda entram
  no cálculo. O restante é guardado linha a linha para a conversão final.
- **Chave tarifária.** Cada lançamento recebe a combinação de subgrupo, modalidade, classe,
  subclasse, detalhe, unidade consumidora (UC), posto e unidade de medida que localiza sua tarifa
  no PCAT.
- **Tarifas homologadas** (3.1). Todas as vigências e camadas ficam numa única tabela gerada a
  partir dos PCATs. Para as vigências anteriores à lei do Desconto Social, a tarifa desse benefício
  é derivada da residencial (3.5).
- **Quantidade e parcelas** (3.2). O SAP lança a quantidade uma única vez por par TUSD/TE. O
  cativo recebe TUSD e TE na linha que carrega a quantidade; o livre, só TUSD.

**Bloco 1: o valor pela tarifa cheia**

- **Valorização.** Quantidade × tarifa, nos 37 itens e nas quatro camadas.
- **Benefícios sociais** (3.5). Os itens isentos são zerados. A participação de cada faixa de
  consumo é estimada pela receita de cada grupo.
- **Tarifa proporcional** (3.6). Cada lançamento usa o ciclo de leitura do seu lote: o peso de cada
  vigência é a fração dos dias do ciclo que ela cobre. REFAT, CNR, o SD e os lotes fora do de-para
  ficam com a tarifa do mês.

**Bloco 2: aplicação dos descontos e ajuste de fechamento**

- **Tarifa Social** (3.5). A faixa gratuita até 80 kWh, estimada pela receita, sai do bloco 1 das
  linhas da Tarifa Social, item a item.

- **Regime APE** (3.4). A isenção registrada no BW abate os encargos da TUSD de energia dos livres
  do grupo A.
- **1ª passada: fonte incentivada** (3.3). A diferença restante dos livres do grupo A é absorvida
  por Fio A + Fio B, com desconto limitado a 100%.
- **2ª passada: ajuste de fechamento.** É o único ajuste do processo. O que ainda sobra é
  distribuído proporcionalmente entre todos os itens do grupo, até o total igualar a receita. É o indicador de qualidade do processo: quanto
  menor, mais a tarifa explica sozinha o faturamento.

**Geradoras, suprimentos e uso da distribuição (SD)**

- **Abertura do SD.** Esses clientes vêm de uma planilha de controle própria, com a tarifa
  efetivamente aplicada. A receita é aberta na proporção dos itens da tarifa homologada da mesma
  chave.

**Saída e conferência**

- **Conferências.** Ao final, o notebook informa se as conferências internas passaram. A
  planilha só deve ser usada se todas passarem: TA = BE + CVA + FIN; parte tarifável mais não
  tarifável igual ao total do BW; abatimento APE nunca maior que o desconto.
- **Layout Receita Faturada** (3.8). O conversor leva a base econômica por grupo tarifário, os
  totais de CVA e FIN, e os tributos e bandeiras aos campos do layout, nas grades FAT, REFAT e CNR.
  Uma aba de conciliação fecha o total com a receita do BW. Os tributos do TOI não são identificados
  no BW e ficam em FAT.

## 5. Resultados obtidos

Versão final do notebook, meses de janeiro a agosto de 2026.

### 5.1 Do valor pela tarifa ao valor faturado

O ponto de partida é a pergunta: quanto da receita faturada a tarifa homologada explica sozinha, e o
que explica o restante? A resposta é uma conta de chegada, que fecha exatamente todo mês:

```
(a) receita faturada = (b) bloco 1 + (d) Tarifa Social + (e) isenção APE + (f) fonte incentivada
                       + (g) 2ª passada + (h) resíduo
```

**O que é cada linha**

- **(a) Receita faturada (base tarifável).** Receita registrada no BW e no SD para os lançamentos
  que têm tarifa homologada (seção 3.8). É o valor que a abertura precisa reproduzir.
- **(b) Bloco 1.** Quantidade faturada × tarifa homologada cheia, com a isenção de encargos dos
  benefícios sociais, as faixas do Desconto Social e a Tarifa Social pela tarifa da subclasse Baixa
  Renda sem a redução (3.5), e a tarifa proporcional pelo ciclo de leitura do lote (3.6). É o que a
  tarifa explica sozinha.
- **(c) Diferença a explicar.** Quanto o bloco 1 fica acima da receita, (b) − (a). Existe porque
  a tabela de tarifas traz a tarifa cheia e parte dos clientes paga com desconto.
- **(d) Tarifa Social.** A faixa gratuita até 80 kWh, estimada pela receita (3.5).
- **(e) Isenção APE.** Redução dos encargos da TUSD de energia dos autoprodutores por equiparação,
  registrada no BW em linhas próprias (3.4).
- **(f) Fonte incentivada (1ª passada).** Desconto na TUSD dos clientes livres do grupo A que compram
  energia incentivada, aplicado só sobre Fio A e Fio B (3.3).
- **(g) 2ª passada, líquida.** O que sobra depois dos descontos regulatórios, distribuído por igual
  entre os itens de cada grupo para fechar na receita: arredondamentos, refaturamentos e acertos de
  faturamento. Não é desconto; é o único ajuste do processo e o principal indicador de qualidade do
  processo.
- **(h) Resíduo sem item tarifário.** Receita de grupos cujo bloco 1 é zero: sem item tarifário para
  escalar, ela fica fora da abertura e é informada à parte.
  - **(g1) dela, TUSD Consumo Próprio não faturada por decisão da Empresa.** A Empresa não fatura a
    TUSD do seu consumo próprio, que vira despesa da companhia; só a TE é faturada. O PxQ valoriza
    esse consumo pela TUSD + TE, porque ele é informado à ANEEL como faturamento, e a TUSD não
    faturada aparece aqui. Parte das instalações de consumo próprio paga TUSD, por isso o valor é
    menor que toda a TUSD do consumo próprio no bloco 1.
  - **(g2) dela, demais grupos.** (g) − (g1).
- **(i) 2ª passada, bruta.** Soma dos valores absolutos da 2ª passada de todos os grupos. Mostra se
  o resultado líquido pequeno esconde ajustes grandes de sinais opostos.

**Valores absolutos**

| Item | Unidade | jan | fev | mar | abr | mai | jun | jul | ago |
|---|---|---|---|---|---|---|---|---|---|
| (a) Receita faturada (base tarifável) | R$ milhões | 1.940,0 | 1.834,9 | 1.797,5 | 1.943,1 | 1.848,0 | 1.706,1 | 1.812,3 | 1.938,0 |
| (b) Bloco 1 (tarifa cheia) | R$ milhões | 2.051,4 | 1.950,9 | 1.910,8 | 2.053,1 | 1.958,6 | 1.815,6 | 1.932,0 | 2.074,7 |
| (c) Diferença a explicar, (b) − (a) | R$ milhões | 111,4 | 116,0 | 113,3 | 110,0 | 110,6 | 109,5 | 119,7 | 136,7 |
| (d) Tarifa Social (faixa gratuita) | R$ milhões | −38,8 | −37,1 | −36,0 | −36,1 | −35,7 | −34,8 | −40,5 | −45,5 |
| (e) Isenção APE | R$ milhões | −9,9 | −16,1 | −14,7 | −9,2 | −11,0 | −13,0 | −13,1 | −16,1 |
| (f) Fonte incentivada (1ª passada) | R$ milhões | −63,0 | −63,5 | −64,1 | −66,3 | −63,0 | −62,9 | −63,8 | −72,8 |
| (g) 2ª passada, líquida | R$ milhões | 0,3 | 0,8 | 1,4 | 1,5 | −0,9 | 1,2 | −2,3 | −2,3 |
| (g1) dela, TUSD Consumo Próprio não faturada por decisão da Empresa | R$ milhões | −1,6 | −1,4 | −0,7 | −0,9 | −2,6 | −0,7 | −0,8 | −3,3 |
| (g2) dela, demais grupos | R$ milhões | 1,9 | 2,2 | 2,2 | 2,3 | 1,7 | 1,9 | −1,5 | 1,0 |
| (h) Resíduo sem item tarifário | R$ milhões | 0,0 | 0,0 | 0,1 | 0,1 | 0,0 | 0,0 | 0,1 | 0,0 |
| (i) 2ª passada, bruta | R$ milhões | 3,8 | 8,0 | 3,4 | 4,1 | 4,7 | 3,1 | 5,8 | 7,9 |

**Em relação à receita faturada**

| Item | Unidade | jan | fev | mar | abr | mai | jun | jul | ago |
|---|---|---|---|---|---|---|---|---|---|
| (c) Diferença a explicar, (b) − (a) | % de (a) | 5,74 | 6,32 | 6,30 | 5,66 | 5,99 | 6,42 | 6,61 | 7,05 |
| (d) Tarifa Social (faixa gratuita) | % de (a) | −2,00 | −2,02 | −2,01 | −1,86 | −1,93 | −2,04 | −2,24 | −2,35 |
| (e) Isenção APE | % de (a) | −0,51 | −0,88 | −0,82 | −0,47 | −0,60 | −0,76 | −0,72 | −0,83 |
| (f) Fonte incentivada (1ª passada) | % de (a) | −3,25 | −3,46 | −3,56 | −3,41 | −3,41 | −3,69 | −3,52 | −3,76 |
| (g) 2ª passada, líquida | % de (a) | +0,01 | +0,04 | +0,08 | +0,08 | −0,05 | +0,07 | −0,13 | −0,12 |
| (g1) dela, TUSD Consumo Próprio não faturada | % de (a) | −0,08 | −0,08 | −0,04 | −0,04 | −0,14 | −0,04 | −0,04 | −0,17 |
| (g2) dela, demais grupos | % de (a) | +0,10 | +0,12 | +0,12 | +0,12 | +0,09 | +0,11 | −0,08 | +0,05 |
| (h) Resíduo sem item tarifário | % de (a) | 0,00 | 0,00 | 0,00 | 0,01 | 0,00 | 0,00 | 0,00 | 0,00 |
| (i) 2ª passada, bruta | % de (a) | 0,20 | 0,44 | 0,19 | 0,21 | 0,25 | 0,18 | 0,32 | 0,41 |

**Leitura**

- A tarifa cheia fica entre 5,7% e 7,1% acima do faturado. Os três descontos regulatórios
  explicam quase toda essa diferença: a fonte incentivada responde por cerca de 3,5 pontos
  percentuais, a Tarifa Social por cerca de 2 pontos e a isenção APE por 0,5 a 0,9 ponto.
- Nos meses sem reajuste, a 2ª passada líquida fica abaixo de 0,1% da receita e a bruta abaixo
  de 0,5%.
- Julho e agosto são os meses com tarifa proporcional, calculada pelo ciclo de leitura de cada lote
  (3.6). A 2ª passada bruta fica em 0,32% e 0,41%. Com o método anterior, que estimava o peso de
  cada vigência pela receita, ela era de 0,52% e 0,51%.
- A maior parte negativa da 2ª passada é a TUSD do consumo próprio, não faturada por decisão da
  Empresa (g1): de 0,7 a 3,3 milhões por mês. Os meses com mais consumo próprio (maio e agosto) são
  os de faturamento acumulado de várias referências. Sem ela, a 2ª passada dos demais grupos (g2) é
  positiva e estável, entre 0,09% e 0,12% da receita, exceto em julho, mês do reajuste.
- O notebook não separa (g1): a 2ª passada fecha o consumo próprio na receita do BW, distribuindo a
  TUSD não faturada entre todos os itens do grupo, TUSD e TE. No bloco 2, o consumo próprio fica na
  receita do BW, com a TE um pouco abaixo do que foi efetivamente cobrado. Por isso o PxQ fica igual
  ao informe de gestão, e não acima dele pela TUSD não faturada, como seria na visão informada à
  ANEEL.

### 5.2 Validações

| Fonte | Quando existe | Para que serve |
|---|---|---|
| Informe de gestão | Mensal | Mesma base do BW. Confere receita e quantidades por segmento. |
| Mercado Receita do Sparta (reajuste) | Uma vez por ano, após o reajuste | Mede o erro do processo e revela falhas graves. Retrata a forma como a ANEEL verifica a receita do período tarifário. Diferenças são esperadas, principalmente nos meses mais recentes. |
| SAMP (dados abertos da ANEEL) | Com meses de defasagem | Referência futura para conferência ao longo do ano. |
| Simulador Sparta | Mensal | Não é referência: é o que o P × Q substitui. |

O que foi conferido:

- **Base de partida.** A extração do BW reproduz exatamente o informe de gestão em receita,
  consumo e demanda, nos segmentos cativo AT, cativo BT e livre AT.
- **Conferências internas.** Passam em todos os meses. O layout Receita Faturada concilia com a
  receita do BW com diferença zero.
- **Desconto de fonte incentivada.** Em março, a 1ª passada ficou a 1,1% do subsídio de fonte
  incentivada calculado pelo Mercado Receita. Isso valida a regra do Fio A + Fio B (3.3).
- **Tarifa Social.** A faixa isenta estimada pela receita ficou em 40,9% da energia em março,
  contra 39,1% no Mercado Receita.

### 5.3 Comparação com o Mercado Receita do reajuste, por família e camada

A tabela compara o bloco 2 do P × Q com o Mercado Receita, de janeiro a maio de 2026, o período em
que o MR do reajuste é a referência mais segura. A base econômica está aberta por família; CVA e
financeiros estão totalizados.

A comparação cobre só a receita tarifável, aberta por item tarifário. Ficam fora tributos (PIS/COFINS
e ICMS), bandeiras tarifárias, ultrapassagem de demanda e energia reativa (UD/ER) e os demais
lançamentos sem tarifa homologada (seção 3.8), que o Mercado Receita não traz.

O subsídio que o Mercado Receita calcula à parte é distribuído pelos itens com o mesmo mecanismo
do próprio MR (abatimento de quantidade). Assim, os dois lados carregam os descontos dentro dos
itens. A cobrança direta de ERCAP aos livres fica fora dos dois lados (3.7).

**O que é cada linha**

- **BE · família.** Parcela da base econômica de cada família de itens (seção 3.1): Fio A, Fio B,
  encargos, perdas e energia. É a parte da tarifa que remunera os custos do ano.
- **CVA (total).** Soma, em todos os itens, da camada que compensa a diferença entre os custos
  previstos e os realizados da Parcela A no ciclo anterior.
- **FIN (total).** Soma dos componentes financeiros da tarifa. Em SP essa camada é negativa (devolução
  ao consumidor); por isso o percentual é calculado sobre o valor absoluto do MR, e um número positivo
  significa que o P × Q devolve menos que o MR.
- **Total.** Soma das linhas acima: a receita tarifável inteira.

Diferença = P × Q − Mercado Receita. Valor negativo indica P × Q abaixo do MR. Nas colunas mensais e no
Acumulado, cada percentual é calculado sobre o valor do MR da própria linha (por isso uma família
menor, como o Fio A, tem percentual maior para a mesma diferença em R$). A última coluna divide a
diferença acumulada de cada linha pela receita total do MR no período (R$ 9.395,9 milhões), o que
mostra o peso de cada família no desvio total.

| Diferença P × Q − MR | Unidade | jan | fev | mar | abr | mai | Acumulado | Acumulado, % da receita do MR |
|---|---|---|---|---|---|---|---|---|
| BE · FIO A | % do MR | −2,38 | −0,72 | −0,75 | −0,95 | −0,75 | −1,13 | −0,10 |
| BE · FIO B | % do MR | −0,74 | −0,70 | −0,69 | −0,93 | −0,67 | −0,75 | −0,21 |
| BE · ENCARGOS | % do MR | +0,62 | −1,28 | −0,76 | +0,49 | +0,39 | −0,09 | −0,03 |
| BE · PERDAS | % do MR | −0,02 | −0,05 | −0,03 | −0,21 | −0,03 | −0,07 | 0,00 |
| BE · ENERGIA | % do MR | −0,13 | −0,20 | −0,17 | −0,41 | −0,19 | −0,22 | −0,07 |
| BE · RI e OUTROS | % do MR | — | — | — | — | — | — | — |
| CVA (total) | % do MR | +2,56 | −2,22 | −1,02 | +2,18 | +1,76 | +0,66 | +0,01 |
| FIN (total) | % do MR | +1,98 | +1,88 | +1,60 | +1,97 | +1,63 | +1,82 | +0,06 |
| **Total** | % do MR | **−0,21** | **−0,66** | **−0,50** | **−0,26** | **−0,14** | **−0,35** | **−0,35** |

Em valores absolutos:

| Diferença P × Q − MR | Unidade | jan | fev | mar | abr | mai | Acumulado |
|---|---|---|---|---|---|---|---|
| BE · FIO A | R$ milhões | −4,3 | −1,2 | −1,3 | −1,7 | −1,3 | −9,8 |
| BE · FIO B | R$ milhões | −4,0 | −3,6 | −3,4 | −5,0 | −3,5 | −19,5 |
| BE · ENCARGOS | R$ milhões | 3,5 | −7,0 | −4,0 | 2,8 | 2,1 | −2,6 |
| BE · PERDAS | R$ milhões | 0,0 | −0,1 | 0,0 | −0,3 | 0,0 | −0,4 |
| BE · ENERGIA | R$ milhões | −0,8 | −1,1 | −0,9 | −2,4 | −1,0 | −6,2 |
| BE · RI e OUTROS | R$ milhões | 0,0 | 0,0 | 0,0 | 0,0 | 0,0 | 0,0 |
| CVA (total) | R$ milhões | 0,4 | −0,4 | −0,2 | 0,4 | 0,3 | 0,5 |
| FIN (total) | R$ milhões | 1,2 | 1,1 | 0,9 | 1,2 | 0,9 | 5,2 |
| **Total** | R$ milhões | **−4,1** | **−12,2** | **−9,0** | **−5,0** | **−2,5** | **−32,7** |
| Referência: receita do Mercado Receita | R$ milhões | 1.944,0 | 1.847,1 | 1.806,4 | 1.948,0 | 1.850,5 | 9.395,9 |

Leitura:

- **Total.** O P × Q fica a menos de 0,7% do Mercado Receita em todos os meses, e a −0,35% no
  acumulado.
- **CVA e FIN.** Os percentuais são maiores porque as camadas são pequenas. A maior diferença
  mensal é de R$ 1,2 milhão, menos de 0,1% da receita do mês.
- **Encargos.** Oscilam em torno de zero (acumulado −0,09%). A isenção APE registrada num mês
  inclui créditos de meses anteriores, e o BW não permite separá-los.
- **Fio A de janeiro.** É o único ponto fora do padrão: a demanda dos livres A2 no BW fica abaixo
  da do MR nesse mês. A investigação está registrada em [`sessions.md`](sessions.md).
- **RI e OUTROS.** São zero nos dois lados em SP.

Detalhe completo em `2-Checks/consolidado_check_mr_camadas_2026-01_a_2026-05.xlsx` e, por opção
de compra e subgrupo, em `2-Checks/consolidado_check_mr_bloco2_2026-01_a_2026-05.xlsx`.

### 5.4 Tarifa homologada × tarifa realizada: onde estão os descontos

O script `3-Py/2-Checks/gera_receita_faturada_bl1_vs_bl2.py` monta o layout Receita Faturada pela
tarifa homologada e pelo bloco 2 (tarifa realizada, a entrega) e separa a diferença em Tarifa Social –
DMR, fonte incentivada e 2ª passada, item a item. A tarifa homologada é o bloco 1 sem as duas parcelas
que não são reembolsadas pela CDE, que ficam só como memória:

- **Isenção de encargos (APE):** com base no PRORET 5.2 (§81 e §91), o consumo isento de encargos sai do
  mercado de rateio da cota da CDE e não é reembolsado. Esses encargos simplesmente não são cobrados
  desse cliente, e o custo fica com os demais consumidores pelo rateio da cota.
- **SCEE - Baixa Renda:** é a TE do SCEE, homologada no reajuste, aplicada sobre a energia compensada.

A DMR faturada do mês vem de `src/dmr_tarifa_social.csv` (informe de gestão). Regras e limites no
[procedimento](1-Entrega/procedimentos/PROCEDIMENTO_PXQ.md), seção "Check tarifa homologada × tarifa
realizada".

Diferença tarifa homologada − tarifa realizada (R$ milhões; positivo = desconto):

| | jan | fev | mar | abr | mai | jun | jul | ago |
|---|---|---|---|---|---|---|---|---|
| Tarifa Social – DMR | 37,3 | 34,5 | 33,3 | 32,7 | 32,9 | 31,9 | 37,5 | 42,7 |
| Fonte incentivada | 63,0 | 63,5 | 64,1 | 66,3 | 63,0 | 62,9 | 63,8 | 72,8 |
| 2ª passada | −0,3 | −0,8 | −1,4 | −1,5 | 0,9 | −1,2 | 2,3 | 2,3 |
| dela, TUSD Consumo Próprio não faturada | 1,6 | 1,4 | 0,7 | 0,9 | 2,6 | 0,7 | 0,8 | 3,3 |
| dela, demais grupos | −1,9 | −2,2 | −2,2 | −2,3 | −1,7 | −1,9 | 1,5 | −1,0 |
| **Total** | **100,0** | **97,3** | **96,0** | **97,5** | **96,9** | **93,7** | **103,6** | **117,8** |
| Retiradas da tarifa homologada (memória): |  |  |  |  |  |  |  |  |
| SCEE - Baixa Renda | 1,5 | 2,6 | 2,7 | 3,4 | 2,7 | 2,8 | 3,1 | 2,8 |
| Isenção de encargos (APE) | 9,9 | 16,1 | 14,7 | 9,2 | 11,0 | 13,0 | 13,1 | 16,1 |
| **Bloco 1 − tarifa realizada** | **111,4** | **116,0** | **113,3** | **110,1** | **110,6** | **109,5** | **119,8** | **136,7** |

A linha "Bloco 1 − tarifa realizada" é a linha (c) da seção 5.1, a menos do resíduo; Tarifa Social –
DMR mais SCEE - Baixa Renda, isenção de encargos (APE), fonte incentivada e 2ª passada são as linhas (d),
(e), (f) e (g) com o sinal invertido.

Leitura:

- **Fonte incentivada.** Cai só nos itens de rede: em maio, Parcela B (Fio B) R$ 37,9 milhões, Rede
  Básica R$ 24,2 milhões, Conexão T R$ 1,1 milhão e CUSD R$ 0,2 milhão.
- **Tarifa Social.** A linha DMR é o subsídio homologado: R$ 282,9 milhões de janeiro a agosto.
- **Isenção de encargos (APE), fora da análise.** Cai só nos encargos: em maio, CDE USO R$ 8,8 milhões e
  PROINFA R$ 1,6 milhão.
- **SCEE - Baixa Renda, fora da análise.** R$ 21,7 milhões no período, estável em cerca de R$ 2,7 milhões
  por mês. Nos clientes Baixa Renda com geração distribuída, a energia compensada é faturada pela TE do
  SCEE (R$ 37,19 contra R$ 316,78/MWh; REN 1.000/2021, art. 655-I), e o benefício da Tarifa Social
  também incide sobre ela (art. 655-G, §8º). O BW não identifica essa energia, que é valorizada pela TE
  comum; a diferença para a DMR está só na tarifa.
- **Também fora do check.** Desconto Social, isenção de encargos da Baixa Renda e as tarifas próprias das
  subclasses ficam no bloco 1 e não aparecem na diferença: não são reembolsados pela CDE.
- **Consumo Próprio.** A 2ª passada do Consumo Próprio cativo é de R$ 2,6 milhões em maio: a tarifa
  cheia dá R$ 5,2 milhões e o faturado é R$ 2,6 milhões. A diferença é a TUSD, que a Empresa não
  fatura no seu consumo próprio; só a TE é faturada (seção 5.1, linha g1).

## 6. Pendências

As pendências abertas do projeto estão na ata, [`sessions.md`](sessions.md), seção 1. Esta seção
explica a que afeta diretamente a entrega.

### 6.1 SCEE GD1 e GD2 no layout Receita Faturada

O layout Receita Faturada separa o cativo em três opções de compra: CATIVO, SCEE – GD1 e
SCEE – GD2. O P × Q não faz essa separação. Todo o movimento dos clientes de geração distribuída
(GD) fica em CATIVO, e as linhas SCEE ficam vazias.

**Por que o BW não permite separar**

- A extração não tem campo que identifique a instalação no SCEE (Sistema de Compensação de
  Energia Elétrica) nem a modalidade, GD I ou GD II. Os clientes GD aparecem como "Cliente Cativo",
  junto com os demais.
- Os códigos de energia compensada (`ZQC223`, `ZQC224`, `ZQC234` e `ZQC235`) não resolvem:
  - cobrem quase só o grupo A na modalidade Verde e a Tarifa Branca. No grupo B convencional, onde
    está a maior parte da GD, a compensação já vem abatida do kWh faturado, sem linha própria;
  - registram o crédito da compensação entre postos (valores negativos), e não o consumo faturado
    do cliente GD;
  - não distinguem GD I de GD II.

  Em junho/2026 esses códigos somam −R$ 0,5 milhão. Pelo Mercado Receita, só a classe Comercial
  teria cerca de R$ 18 milhões em GD1 e GD2 no mesmo mês.

**Alternativas avaliadas e descartadas**

As duas usam o Mercado Receita (MR) do Sparta, que separa GD I, GD II e GD II – SCEE no campo
`TipoMercado`.

| Alternativa | Como funciona | Por que foi descartada |
|---|---|---|
| Proporção do MR | A linha CATIVO / classe / energia do layout é dividida pela participação de GD1 e GD2 no MWh do cativo no MR, no mesmo mês e classe | Atribui ao GD a tarifa média do cativo. O GD II – SCEE tem R$/MWh cerca de 40% menor e não tem ICMS nem PIS/COFINS no MR |
| Valores do MR com abatimento | As linhas SCEE recebem os valores do MR, coluna a coluna, e a linha CATIVO fica com o BW menos o MR | O MR é simulado (tarifas arredondadas, sem tarifa proporcional, meses aparentemente repetidos) e não separa energia de demanda nos componentes. Todo o erro do MR iria para o saldo do CATIVO |

As duas dão resultados bem diferentes entre si. Em junho/2026, na classe Comercial, a proporção
levaria R$ 9,7 milhões para GD1 e R$ 8,5 milhões para GD2; os valores do MR, R$ 9,1 milhões e
R$ 6,4 milhões.

**Por que a aproximação não foi feita**

- **A entrega é receita faturada.** Todo valor do layout vem de um lançamento do BW. Qualquer das
  duas alternativas colocaria uma estimativa nas linhas SCEE e alteraria a linha CATIVO, que hoje
  é dado real.
- **O erro seria grande e não mensurável.** No GD2 Comercial de junho, as duas alternativas
  divergem em um terço (R$ 8,5 contra R$ 6,4 milhões), e não há referência faturada para saber qual está mais perto da
  realidade.
- **O total não ganha nada.** O total do cativo de cada classe é o mesmo com ou sem a separação. A
  aproximação só redistribuiria valores dentro dele, com um critério sem base no faturamento.
- **Manter em CATIVO é o tratamento conservador.** O valor está correto no agregado, e a ausência
  das linhas SCEE fica explícita, em vez de ser preenchida com um número que parece faturado e não
  é.

**O que resolve**

Um campo na extração do BW que identifique a instalação no SCEE e a modalidade (GD I ou GD II),
ou o encontro de contas da SCEE separado do faturamento bruto. Com ele, cada lançamento do BW é
classificado na origem, e o conversor passa a gravar as linhas SCEE com o valor faturado. O pedido
está registrado em [`sessions.md`](sessions.md), item I4.
