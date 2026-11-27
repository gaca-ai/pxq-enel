# Procedimento RNF (Enel CE e Enel RJ): especificação técnica

Especificação dos arquivos de RNF do sistema comercial e das regras para abri-los por item tarifário. As
decisões estão em [`sessions.md`](../../sessions.md); a visão geral, no [README](../../README.md). Quando
uma regra é igual à de SP, a referência é o [procedimento de SP](../../../../SP/RNF/1-Entrega/procedimentos/PROCEDIMENTO_RNF.md).

Valores monetários em milhões de reais (MR$), energia em MWh e demanda em MW.

## 1. Arquivos do sistema

| Item | Especificação |
|---|---|
| Arquivos | Um por nível de tensão: AT e BT, com exatamente o mesmo layout. O BT tem cerca de 1 GB por mês: leitura em blocos e conversão para parquet |
| Nível de detalhe | Uma linha por cliente (sem agregação) |
| Mercado | Somente cativos; livres em arquivo separado |
| Formato | Texto UTF-8, separador `;`, decimal `.`, primeira linha com o nome dos campos |
| Nome (exemplo) | `CLIENTES_ALTA_AG_28092026_082440.txt`: AT, gerado em 28/09/2026 às 08:24:40 |
| Unidades | Energia em kWh; preço em R$/kWh; valores em R$ |

### 1.1 Campos

| Grupo | Campo | Conteúdo |
|---|---|---|
| Identificação | `NUMERO CLIENTE`, `CONTA CONTRATO` | Chave do cliente |
| | `CLIENTE NOVO` | SIM ou NÃO |
| | `CLASSE` | Código numérico (10, 20, 30...). Não corresponde à classe de consumo e não é usado |
| | `SUBCLASSE` | Subclasse (ex.: COMOS, PPMUN, COMBOPT); dela sai a classe (seção 4) |
| | `CODIGO DO MUNICIPIO` | Código do município |
| | `CODIGO TARIFA` | Subgrupo e modalidade: A2_AZ, A3_AZ, A3A_VD, A4_AZ, A4_VD, B1_RESID, B2_RUR, B3_OUTROS |
| Ciclo | `LOTE ANTERIOR`, `LEITURA ANTERIOR`, `ULTIMO LOTE`, `ULTIMA LEITURA` | Lote e data (AAAAMMDD) das duas últimas leituras |
| | `DIAS FATURADOS` | Dias do último ciclo faturado; separado em dias dentro do mês de referência e do mês anterior |
| | `DIAS ESTIMADOS` | Dias entre a última leitura e o fim do mês de referência |
| Energia | `CONSUMO LIDO` | Consumo do último ciclo (kWh) |
| | `COMPRAS PARA O PERIODO FATURADO`, `COMPRAS PARA O PERIODO ESTIMADO` | Energia comprada pela distribuidora em cada período |
| | `FATOR` | Ajuste pela compra de energia (seção 2) |
| | `RNF SEM FATOR`, `RNF COM FATOR` | Energia não faturada (kWh) |
| Valor da energia | `PREÇO PRO RATA TARIFA` | Zerado no exemplo |
| | `PREÇO MÉDIO` | R$/kWh, com PIS/COFINS, sem ICMS e sem bandeira |
| | `VALOR MONETÁRIO` | Valor da energia não faturada (R$) |
| Bandeira | `BANDEIRA TARIFÁRIA SEM TRIBUTO`, `BANDEIRA TARIFÁRIA COM TRIBUTO` | Adicional de bandeira (R$) |
| Demanda | `DEMANDA FATURADA` | Valor da demanda do último ciclo (R$; não vem em kW) |
| | `RENDA NÃO FATURADA DA DEMANDA FATURADA SEM ENCARGOS` | Demanda não faturada (R$) |
| | `PIS/COFINS REFERENTE A RENDA NÃO FATURADA DA DEMANDA FATURADA` | PIS/COFINS sobre a demanda não faturada (R$) |
| Versão CONSFAT | Campos com sufixo `(CONSFAT)` | Mesmo cálculo sobre `CONSUMO FATURADO` em vez de `CONSUMO LIDO` |

## 2. Regras de cálculo do arquivo

Verificadas linha a linha no exemplo.

| Campo | Regra |
|---|---|
| `RNF SEM FATOR` | `CONSUMO LIDO` × `DIAS ESTIMADOS` / `DIAS FATURADOS` |
| `FATOR` | (compras do período estimado / dias estimados) / (compras do período faturado / dias faturados) |
| `RNF COM FATOR` | `RNF SEM FATOR` × `FATOR` |
| `VALOR MONETÁRIO` | `RNF COM FATOR` × `PREÇO MÉDIO` |
| Demanda não faturada | `DEMANDA FATURADA` × `DIAS ESTIMADOS` / `DIAS FATURADOS` |
| PIS/COFINS da demanda | 4,29% da demanda não faturada |

PIS/COFINS: os valores de energia e demanda contêm PIS/COFINS e não contêm ICMS, como em SP. O valor líquido
é o bruto / (1 + alíquota). A bandeira vem em campo próprio, com e sem tributo.

A versão CONSFAT difere da versão principal em 8 MWh no exemplo; o processo usa a versão principal (a
confirmar).

## 3. Totais do exemplo (AT, referência set/2026)

5.323 clientes; últimas leituras em 01/09 na maioria (29 dias estimados).

| Tarifa | Energia (MWh) | Valor da energia (MR$) |
|---|---:|---:|
| A4 Verde | 36.922 | 27,06 |
| A4 Azul | 4.320 | 2,48 |
| A2, A3 e A3a | 165 | 0,09 |
| B (optantes) | 7.874 | 8,70 |
| Total | 49.281 | 38,32 |

| Outros componentes | MR$ |
|---|---:|
| Demanda não faturada (sem encargos) | 13,39 |
| PIS/COFINS da demanda | 0,57 |
| Bandeira (com tributo) | 1,01 |

## 4. Abertura por item tarifário

| Etapa | Regra |
|---|---|
| Chave tarifária | `CODIGO TARIFA` → subgrupo e modalidade (a tarifa segue sempre o código, mesmo quando difere da classe); `SUBCLASSE` → classe pelo de-para `docs/CL_SUB_TAR.xlsx` (seção 4.1) |
| PIS/COFINS | Retirado antes do bloco 2 |
| Bloco 1 | Energia × (TE + TUSD) homologadas; mesma quantidade nas duas componentes |
| Posto horário | Energia do grupo A separada entre ponta e fora ponta pela composição do faturado do mês (mesmo subgrupo e modalidade) |
| Demanda | Valor aberto por item tarifário pela composição do faturado do mês |
| TE e TUSD | Valor dividido pela composição do faturado do mês |
| Bloco 2 | Diferença entre o valor líquido do arquivo e o bloco 1, aberta como em SP |
| Bandeira | Fora da abertura, em coluna própria do layout |
| Saídas | Iguais às de SP: relatório detalhado e `receita_nao_faturada` com provisão, estorno, saldo e saldo YTD |

### 4.1 De-para de subclasse para classe

Arquivo `docs/CL_SUB_TAR.xlsx`, aba `classe_subclasse_tarifa`: colunas `Subclasse` e `Classe` (45 subclasses).
As colunas à esquerda (`CLASSE`, `SUBCLASSE`, `CODIGO TARIFA`, `Contagem`, `Tensão`) listam as combinações
encontradas nos arquivos AT e BT, com a quantidade de clientes, e servem só de referência. Uma subclasse
nova, fora do de-para, interrompe o processo.

| Classe | Subclasses |
|---|---|
| Residencial | `REPLN`, `RESBOPT` |
| Residencial Baixa Renda | `REBRBPC`, `REBRIND`, `REBRQUI`, `REBXDS`, `REBXR` |
| Comercial | `COMBOPT`, `COMCON`, `COMCT`, `COMER`, `COMFIL`, `COMOS`, `COMREL`, `COMROD`, `COMSEM`, `COMTP` |
| Industrial | `INDBOPT`, `INDUST` |
| Rural | `RUAGU`, `RUAQC`, `RUAQCOPT`, `RUCES`, `RURAG`, `RURBOPT`, `RURCR`, `RUREA`, `RURIR`, `RURRR`, `RURSP` |
| Poder Público | `PPEBOPT`, `PPEST`, `PPFBOPT`, `PPFED`, `PPMBOPT`, `PPMUN` |
| Iluminação Pública | `ILPUBL` |
| Serviço Público | `SPAES`, `SPAESBOT`, `SPERVBOT`, `SPTER` |
| Consumo Próprio | `CPCPE`, `CPCPO`, `CPCPSE`, `CPCPU` |

Subclasses terminadas em `BOPT`/`BOT` são optantes pela tarifa B (clientes do grupo A faturados pela
tarifa do grupo B). Combinações de classe e código de tarifa incoerentes (ex.: subclasse comercial com
`B1_RESID`; `REPLN` com `A4_VD`) também são tratadas como optante B.

### 4.2 Residencial Baixa Renda

Regras de cada benefício em `docs/tarifas_baixa_renda_aneel.xlsx`.

| Subclasse | Benefício | Tratamento |
|---|---|---|
| `REBXR` | Baixa Renda geral (Tarifa Social) | Faixa de consumo estimada pelo valor, como em SP |
| `REBRBPC` | BPC | Idem |
| `REBRIND` | Indígena | Idem |
| `REBRQUI` | Quilombola | Idem |
| `REBXDS` | Desconto Social (DSEE) | Chave própria, como em SP |

## 5. Pontos em aberto

Ver [`sessions.md`](../../sessions.md), parte 1.
