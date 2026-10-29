# Procedimento RNF: como a renda não faturada vira abertura por item tarifário

Manual gerencial do notebook `Abertura_RNF.ipynb`. Cada seção corresponde a um passo do notebook e
explica o que ele faz e por que faz. As regras tarifárias comuns ao Faturamento (estrutura da tarifa,
cativo e livre, fonte incentivada, APE, benefícios sociais) estão no
[README do BW](../../../BW/README.md), seção 3, e não são repetidas aqui.

Convenção de unidades: valores monetários em milhões de reais (MR$), energia em MWh e demanda em MW.
A exceção é o consolidado `receita_nao_faturada`, que segue o layout da área (R$ mil).

---

## A ideia em uma página

A RNF (Renda Não Faturada) é a receita da energia já entregue e ainda não faturada no fechamento do
mês. Ela é provisionada no mês e estornada no mês seguinte, quando a conta é emitida. O efeito na
receita do mês é a **provisão do mês menos o estorno da provisão do mês anterior**. No acumulado do
ano, os estornos intermediários se anulam e o efeito é a provisão do mês menos a provisão de
dezembro do ano anterior.

O SAP calcula a RNF do sistema na transação ZUBIL_146, por lote, classe/subclasse e categoria
tarifária: quantidade não faturada e valor. A contabilidade completa com ajustes calculados fora do
SAP (TOI, faturas pendentes, art. 323, iluminação pública, ajuste de carga, ajuste livre). O
resultado é um valor total, sem composição tarifária.

O processo decompõe esse valor nos 37 itens da tarifa homologada, nas sete famílias e nas quatro
camadas (TA = BE + CVA + FIN), com a mesma régua do Faturamento:

| Bloco | O que é | Para que serve |
|---|---|---|
| **Bloco 1, calculado** | quantidade não faturada × tarifa homologada cheia | Mostra o que a tarifa explica sozinha |
| **Bloco 2, entrega** | o bloco 1 levado ao valor registrado, linha a linha | É a RNF aberta por item |

O ZUBIL não informa tudo o que o valor contém. Ele traz a energia sem posto horário e, no livre do
grupo A, só parte da demanda; o valor, porém, inclui a energia de ponta, a demanda e os descontos de
fonte incentivada e APE. O bloco 2 devolve cada diferença aos itens a que ela pertence:

| Diferença | Onde vai |
|---|---|
| Desconto sobre a demanda do livre | Fio A e Fio B (aplicação do desconto) |
| Ponta e demanda não informadas (grupo A) | composição do faturado do mesmo mês, nos postos que o bloco 1 não valoriza |
| Demais linhas | rateio proporcional (2ª passada) |

---

## Passo 1 — Parâmetros

Todo mês muda só `REFERENCIA` (`aaaa-mm`). Os arquivos são localizados pelo padrão do nome:

| Arquivo | Pasta | Para que serve |
|---|---|---|
| `RNF CCS <mmaaaa> Fator K Nova Metodologia.xlsb` | `src/` | Fonte do ZUBIL_146 (só a aba `Relatório SAP_ZUBIL_146` é lida) |
| `BASE_ESTIMADO_RNF_POR_CLUSTER_<aaaamm>.xlsx` | `src/` | Ajustes manuais e alíquotas de PIS e COFINS |
| `SP_Consolidacao Margem_<aaaa>.<mm>.xlsx` | `src/` | Valor de comparação (aba `Base RNF`) |
| `tarifa_long.csv` | `../BW/tbl/tarifas/` | Tabela de tarifas homologadas, única para Faturamento e RNF |
| `pxq_<aaaa-mm>.xlsx` | `../BW/1-Entrega/pxq/` | P × Q do Faturamento do mesmo mês (composição do faturado, Passo 9) |

`POSTO_DEMANDA_AZUL` define o posto em que a demanda AW do Azul é valorizada (fora ponta).

**Ordem de execução.** O P × Q do Faturamento do mês precisa estar pronto antes da RNF. Os meses da
RNF rodam em ordem cronológica, porque cada mês lê a provisão do mês anterior para montar o estorno.

## Passo 2 — Ler o ZUBIL_146

O notebook lê a aba `Relatório SAP_ZUBIL_146` e confere o cabeçalho de cada coluna usada. Se o
layout mudar, o processo para. Também confere que o exercício do ZUBIL é o mês de referência.

Colunas usadas (cabeçalho na linha 3):

| Coluna | Rótulo | Uso |
|---|---|---|
| C, J, K | Classe 1, Classe/SubClasse, Categoria Tarifa | Classe, subclasse, subgrupo, modalidade e mercado |
| H | Lote | Lote de faturamento |
| I | Classe (ELGA/ELGB) | Separa as linhas de detalhe dos subtotais |
| S | Valor Ativo Não faturado TE | Valor do cativo |
| U | Valor Ativo Não fat. TUSD TE | Valor do cativo; no livre, valor da demanda |
| W | Valor Ativo Não faturado TUSD | Valor do livre |
| AA | Valor Reativo Não Faturado | Reativo (não tarifável) |
| AI | Consumo Ativo Não Faturado TE | Energia do cativo |
| AM | Consumo Ativo Não Faturado TUSD | Energia do livre |
| AW | Demanda Ativa Não Faturado | Demanda (cativo e livre) |
| BG | Valor Adic. Bandeiras Não Faturado | Bandeira (não tarifável) |

O script `3-Py/2-Checks/extrai_zubil_colunas.py` grava essas colunas num Excel à parte
(`2-Checks/zubil/`), para conferência.

## Passo 3 — Linhas do SAP

As linhas do ZUBIL são somadas por classe, subclasse e categoria tarifária. As categorias e as classes
do SAP são traduzidas por tabelas explícitas; uma categoria nova interrompe o processo, para ser
mapeada.

- **Mercado:** dado pela categoria (`LIVRE` no nome é livre).
- **Cativo:** energia da coluna AI; demanda da AW no grupo A; valor = S + U + W.
- **Livre:** cada linha gera duas.

  | Linha | Quantidade | Valor |
  |---|---|---|
  | Energia (`SAP ZUBIL_146`) | coluna AM | coluna W |
  | Demanda (`SAP ZUBIL_146 (demanda livre)`) | coluna AW | colunas S + U |

  **Por quê.** O livre paga só TUSD. Nas categorias livres, o SAP lança a receita de demanda na
  coluna U ("TUSD TE"), sem energia correspondente, e a de energia na coluna W. A demanda AW aparece
  exatamente nessas linhas. Na modalidade Verde, o valor da coluna U é 50% de AW × tarifa de demanda:
  é a demanda com o desconto de fonte incentivada.

- **Benefício social:** subclasses de Tarifa Social e de Desconto Social são identificadas (Passo 8).
- **Não tarifáveis:** reativo (AA) e bandeira (BG) ficam em colunas à parte.

## Passo 4 — Ajustes manuais

Os ajustes vêm da base no formato novo. É usado tudo o que não é o cluster `Sistema`, que é a própria
RNF do SAP e já foi lida do ZUBIL.

| Cluster | Origem | Tratamento |
|---|---|---|
| `CNR` | TOI | linha a linha, com os atributos da base |
| `PENDENTES` | Faturas pendentes | linha a linha |
| `323` | Art. 323 | linha a linha |
| `RNF IP` | Iluminação pública | linha a linha (B4a) |
| `Ajuste Carga` | Ajuste de carga | sem atributos: distribuído pelas linhas cativas do SAP |
| `Ajuste Livre` | Ajuste livre | sem atributos: distribuído pelas linhas livres do SAP |
| `bandeira Tarifária` | Bandeira dos ajustes | só bandeira: não tarifável |

Na distribuição, a energia segue a participação de cada linha do SAP na energia, e o valor, a
participação no valor. A modalidade `GERAÇÃO` não tem tarifa de energia e é valorizada como Verde. A
classe `RESIDENCIAL BAIXA RENDA` recebe o tratamento da Tarifa Social: a base valoriza essa classe
pela mesma tarifa média da Tarifa Social do ZUBIL.

## Passo 5 — PIS/COFINS

Os valores do ZUBIL contêm PIS/COFINS; a tarifa homologada não tem tributos. O alvo do bloco 2 é o
valor líquido, pela mesma convenção da base de ajustes:

`valor líquido = valor com tributos ÷ (1 + alíquota de PIS + alíquota de COFINS)`

As alíquotas vêm da aba `PIS_COFINS` da base de ajustes. A base de ajustes já traz o valor líquido e
o tributo de cada linha. O tributo vai para os não tarifáveis e volta na conciliação.

## Passo 6 — Chave tarifária e posto horário

| Origem | Chave |
|---|---|
| Grupo A e B3 | classe e subclasse `Não se aplica` |
| B1 | classe `Residencial`; subclasse `Residencial`, ou a do benefício social |
| B2 / B4 | `Rural` / `Iluminação pública – B4a` |
| Detalhe e acessante | `Não se aplica` (APE, SCEE e fonte incentivada não são identificados) |

**Posto horário.** O ZUBIL não separa ponta e fora ponta. No bloco 1, toda a energia do grupo A e da
Branca vai para a tarifa fora ponta; a demanda AW do Azul, para a demanda fora ponta. O que a ponta
acrescenta ao valor é tratado no bloco 2 (Passo 9).

## Passo 7 — Tarifa da vigência

Vale a vigência em vigor no último dia do mês. No mês de reajuste (4 de julho em SP), a tarifa
proporcional ainda não está implementada: o processo avisa e usa a vigência nova.

## Passo 8 — Bloco 1

Quantidade × tarifa homologada, item a item e camada a camada, com energia (E) e demanda (D)
separadas desde aqui, porque o layout da área exige a unidade.

- Cativo paga TUSD e TE na energia; livre, só TUSD. A demanda é só TUSD.
- **Benefícios sociais.** A faixa de consumo não é informada. Cada linha é valorizada nas duas faixas,
  e a participação da faixa 02 é a que reproduz o valor registrado: `s2 = (R − A) / (B − A)`,
  limitada a [0, 1]. Tarifa Social: faixa 01 gratuita, faixa 02 pela tarifa `Baixa Renda`. Desconto
  Social: chaves próprias das duas faixas.

## Passo 9 — Bloco 2

A diferença entre o valor líquido e o bloco 1 é devolvida aos itens a que pertence:

| Linha | Regime | Regra |
|---|---|---|
| Demanda do livre | fio | Bloco 1 = AW × tarifa cheia de demanda. A diferença é a **aplicação do desconto** de fonte incentivada e APE, levada a Fio A e Fio B na composição da tarifa de demanda. |
| Energia do grupo A (livre e cativo) | faturado | A diferença vai para a composição do faturado do mesmo mês (bloco 2 do P × Q do BW), mesmo mercado, subgrupo e modalidade, só nos postos que o bloco 1 não valoriza. |
| Demais | proporcional | Rateio por todos os itens, na proporção do bloco 1 (2ª passada). |

Postos que entram na composição do faturado:

| Mercado | O bloco 1 valoriza | A diferença segue o faturado de |
|---|---|---|
| Livre | energia fora ponta | energia de ponta + demanda (ponta e fora ponta) |
| Cativo | energia fora ponta + demanda AW fora ponta | energia de ponta + demanda de ponta (Azul) |

**Por quê.** A RNF do grupo A é praticamente um mês de faturamento: energia e R$/MWh ficam próximos do
faturado do mesmo mês. O que o bloco 1 não captura é, no Verde, a energia de ponta (a TUSD de ponta é
muito mais cara) e, no Azul, a demanda que a coluna AW não traz. O faturado do mês tem exatamente
essa composição, já com os descontos aplicados pela receita. Sem faturado no grupo, vale o regime fio
no livre e o rateio proporcional no cativo.

O P × Q do BW não traz CVA e FIN por item: em cada linha do faturado, a parte TA − BE de cada item é
separada entre CVA e FIN na proporção dos totais da linha.

Linha com bloco 1 zero e valor registrado fica como resíduo, informado à parte.

## Passo 10 — Conferências

1. Bloco 2 + resíduo = valor líquido, linha a linha.
2. TA = BE + CVA + FIN no bloco 2.
3. Famílias = componentes.
4. ZUBIL × cluster `Sistema` da base de ajustes: energia, valor com tributos e valor líquido por
   mercado devem ser iguais (a base de ajustes também traz o `Sistema`, calculado do ZUBIL).
5. Medida, não aprovação: valor líquido ÷ bloco 1 por subgrupo e modalidade no rateio proporcional.

A planilha só deve ser usada se as conferências 1 a 4 passarem.

## Passo 11 — Conciliação com a `Base RNF`

O total aberto mais PIS/COFINS, reativo e bandeira é comparado com a aba `Base RNF` da consolidação
da margem, bloco a bloco:

| Bloco da `Base RNF` | Composição na abertura (com tributos) |
|---|---|
| Consumo Cativo | colunas S e U do ZUBIL (todas as categorias) + reativo + iluminação pública |
| TOI | cluster `CNR` |
| Pendentes | cluster `PENDENTES` |
| Artigo 113 | cluster `323` |
| Carga | cluster `Ajuste Carga` |
| Bandeira | bandeira do ZUBIL e dos ajustes |
| Total Livres | coluna W do ZUBIL + cluster `Ajuste Livre` |

A `Base RNF` classifica o valor da demanda do livre (colunas S e U) como cativo; a abertura o trata
como livre. Por isso a comparação usa o agrupamento da `Base RNF`.

## Passo 12 — Base no layout Receita Faturada

A provisão é convertida no layout **Receita Faturada** da área, com o mesmo modelo e o mesmo de-para
de itens para grupos do Faturamento (`../BW/tbl/modelo/De-Para_Modelo_PyC.xlsx`, conversor
`../BW/3-Py/1-Entrega/gera_receita_faturada.py`).

| Campo | Regra |
|---|---|
| `TipoFaturamento` | `RNF` (SAP); `RNF_TOI`, `RNF_PENDENTES`, `RNF_ART323`, `RNF_CARGA`, `RNF_IP`, `RNF_AJUSTE_LIVRE`; `RNF_BANDEIRA` (bandeira dos ajustes) |
| `OpcaoCompra` | CATIVO ou LIVRE |
| `Classe` | classe do layout; Tarifa Social e Desconto Social em `Residencial Baixa Renda` |
| `Unidade` | E (MWh) ou D (MW) |
| Parcela B a TFSEE | bloco 2 na camada BE, pelo de-para do modelo |
| Base Financeira / Base CVA | totais FIN e CVA do bloco 2 |
| Bandeira | bandeira do ZUBIL e dos ajustes |
| UD/ER | reativo |
| PIS/COFINS | tributo de cada linha, repartido entre E e D pela receita de cada unidade |
| ICMS | sem valor |

## Passo 13 — Relatório detalhado

`1-Entrega/rnf/rnf_<aaaa-mm>.xlsx`:

| Aba | Conteúdo |
|---|---|
| `Provisao`, `Estorno`, `Saldo`, `Saldo_YTD` | efeito na receita por origem, mercado, classe, subgrupo e modalidade, com os 37 itens e as sete famílias (TA) e os não tarifáveis em linhas próprias |
| `Componentes` | bloco 1 e bloco 2 por linha, 37 itens, tarifa de aplicação |
| `Componentes_BE` | bloco 2 na base econômica |
| `Familias` | bloco 2 nas sete famílias e quatro camadas |
| `Resumo` | MWh, valor e tarifa média por origem, mercado e classe |
| `Nao_tarifaveis` | PIS/COFINS, reativo e bandeira por mercado e classe |
| `Conciliacao`, `Conciliacao_Base_RNF` | fechamento com a `Base RNF` |
| `Layout_base` | provisão no layout Receita Faturada, em R$; fonte do estorno do mês seguinte |

O estorno é a provisão do mês anterior, lida da saída já gravada daquele mês, com sinal trocado: usa
a tarifa e a composição daquele mês, e não uma reabertura com dados atuais. O `Saldo_YTD` só é
gravado quando existe a saída de dezembro do ano anterior.

## Passo 14 — Consolidado `receita_nao_faturada`

`1-Entrega/receita_nao_faturada/receita_nao_faturada_<aaaa-mm>.xlsx`, a partir do modelo da área, no
mesmo layout e unidade da `receita_faturada` (R$ mil; quantidade em MWh ou MW):

| Aba | Conteúdo |
|---|---|
| `Provisao` | provisão do mês |
| `Estorno` | estorno da provisão do mês anterior (`Layout_base` do relatório detalhado daquele mês, com sinal trocado) |
| `Saldo` | provisão + estorno |
| `Saldo_YTD` | provisão − provisão de dezembro do ano anterior (quando houver) |
| `Conferencia` | total gravado + fora do layout = total da RNF |

Sem a saída do mês anterior, as abas `Estorno` e `Saldo` trazem um aviso em vez de dados.

---

## Rotina mensal

1. Rodar o P × Q do Faturamento do mês (`SP/BW`).
2. Colocar em `src/` o relatório da RNF do mês (fonte do ZUBIL), a base de ajustes manuais e a
   consolidação da margem.
3. Ajustar `REFERENCIA` no Passo 1 e executar o notebook.
4. Verificar a mensagem "TODAS AS CONFERENCIAS PASSARAM" e a diferença com a `Base RNF`.
