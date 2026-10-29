# Procedimento P × Q: como o faturamento vira abertura por item tarifário

Manual gerencial do notebook `Abertura_PxQ.ipynb`. Cada seção corresponde a um passo do notebook e
explica o que ele faz e por que faz. As normas citadas estão resumidas na seção 3 (Embasamento
regulatório) do [README](../../README.md).

---

## A ideia em uma página

O SAP registra quanto foi faturado, mas não por quê. Uma linha de receita de R$ 1 milhão de um
cliente cativo residencial não diz quanto daquele valor é rede, encargo, perda ou energia.

A tarifa homologada pela ANEEL diz exatamente isso. Ela decompõe o preço de cada combinação de
subgrupo, modalidade, classe e posto em 37 itens, cada um com quatro camadas: base econômica, CVA,
financeiros e o total, a tarifa de aplicação. Se eu souber a quantidade faturada e a tarifa
daquela combinação, sei a composição.

O processo tem duas etapas, que geram dois blocos de valores lado a lado na saída:

| Bloco | O que é | Para que serve |
|---|---|---|
| **Bloco 1, calculado** | quantidade × tarifa homologada cheia | É a régua: mostra o que a tarifa explica sozinha |
| **Bloco 2, com descontos** | o bloco 1 após a aplicação dos descontos regulatórios e do ajuste de fechamento (2ª passada) | É o produto final: a receita aberta por item |

O bloco 1 nunca fecha com a receita, porque a tabela de tarifas traz a tarifa cheia e muitos
clientes pagam com desconto. A essência do método está em **como** o bloco 2 trata essa
diferença. Cada desconto é devolvido aos itens sobre os quais a regulação o faz incidir. Ele não é
rateado por igual sobre a tarifa inteira. É isso que mantém, por exemplo, os encargos intactos
quando o desconto é sobre a rede.

O notebook tem duas partes:

- **Parte I.** O fornecimento faturado pelo SAP, que vem pelo BW.
- **Parte II.** Geradoras, suprimentos e uso da distribuição (SD), que são faturados numa planilha
  de controle à parte.

---

# Parte I — Fornecimento (BW)

## Passo 0 — Parâmetros

Tudo o que muda de um mês para outro está nesta célula. O restante do notebook não é editado.

| Parâmetro | O que define | Quando mexer |
|---|---|---|
| `REFERENCIA` | Mês de referência (`aaaa-mm`); o arquivo do BW do mês (`ARQ_BW`) sai dele | Todo mês |
| `ARQ_SD` | Planilha do SD em `src/` (`SD_ENEL SP_completo.xlsx`) | Quando chegar um arquivo novo |
| `APLICAR_2A_PASSADA` | Ligado: entrega padrão, fecha na receita. Desligado: o bloco 2 para antes do fechamento e as diferenças saem linha a linha, em arquivo com sufixo `_sem_2a_passada` | Para investigar diferenças |
| `REGIME_APE` | Liga o tratamento da isenção do autoprodutor por equiparação (Passo 9) | Mantido ligado |
| `ARQ_CALENDARIO`, `ARQ_DE_PARA_LOTE` | Calendário de leitura e de-para dos lotes, em `tbl/calendario/` (Passo 4) | Não mudam; as tabelas é que são atualizadas |

## Passo 1 — Ler a extração do BW

**Preparação (antes do notebook).** A extração mensal chega em csv (mais de 200 MB) e é deixada em
`src/csv_original/`. O script `3-Py/1-Entrega/prepara_base_bw.py`:

1. lê o csv (separador `;`, UTF-8, 4 linhas de cabeçalho, números no formato brasileiro);
2. grava o parquet limpo em `src/parquet/MM-AAAA_ENEL SP_RECEITA.parquet`: nomes de coluna únicos,
   números e datas convertidos, só a coluna "Resultado" de cada medida, sem a linha de total. Quando
   o BW grava a unidade junto do número ("6.481.782 KW"), ela é retirada; quando imprime "*" no
   Resultado (unidades misturadas), o valor é a soma das colunas kW, kW.h e "Não atribuído". As
   colunas são localizadas pelo nome no cabeçalho, nunca pela posição: uma coluna nova ou retirada
   pelo BW (como a data de faturamento, retirada em 2026) não quebra a leitura; se faltar uma coluna
   obrigatória, o script para com erro;
3. confere as cinco medidas contra a linha "Result.global" do csv: montante ao centavo e
   quantidades com tolerância de 1 parte em 10 mil (o BW arredonda as quantidades linha a linha);
4. só se a conferência passar, compacta o csv em `src/csv_original/` (mesmo nome, `.zip`) e o apaga.

Os meses que já têm zip e parquet são pulados; `--forcar` refaz todos a partir dos zips.

**Leitura no notebook.** O parquet traz, além das colunas da extração antiga, o lote de leitura
(Conjunto de Contrato), a referência original e o indicador de refaturamento.
Cada lançamento recebe um **tipo de faturamento**, que preenche a coluna `TipoFaturamento` do layout
Receita Faturada:

| Tipo | Regra |
|---|---|
| `CNR` | Operação principal "Faturamento Manual de Energia": o TOI (Termo de Ocorrência e Inspeção). Vale mesmo quando o lançamento também tem o indicador de refaturamento. |
| `REFAT` | Indicador de refaturamento = X. O estorno da conta original e a conta nova vêm em pares, ambos marcados; os dois entram, e o REFAT fica líquido. |
| `FAT` | Os demais lançamentos. |

Lançamentos sem o X com mês de referência antigo são faturamento normal (FAT). Os tributos vêm do
BW numa operação própria ("R/Não atribuído"), sem identificação de TOI: os tributos do TOI ficam em
FAT; os do refaturamento têm o X e ficam em REFAT.

Lote e referência original são somados. Do lote fica, em cada linha, a quantidade ponderada pelo
peso de cada vigência no ciclo de leitura (Passo 4). Cada linha da base no notebook é uma
combinação de:

- categoria tarifária (que contém o subgrupo e a modalidade);
- código estatístico (`ZQC` para consumo, `ZQD` para demanda);
- operação e suboperação;
- classe e subclasse;
- mercado (cativo, livre ou híbrido);
- tipo de faturamento.

A coluna de montante é a **receita registrada**. É ela que o resultado final precisa reproduzir. Os
totais de consumo (TUSD + TE, em kWh) e de demanda (TUSD + faturada, em kW) são calculados na leitura,
como as antigas colunas de controle da extração em Excel.

Essa base é a mesma do informe de gestão.

## Passo 2 — Separar o que é fornecimento

Só entram no cálculo os lançamentos que têm tarifa homologada por item: **30 códigos
estatísticos** de consumo e demanda.

- **Consumo do cativo:** 16 códigos, inclusive a energia compensada do posto convencional.

  | Código | Descrição no BW |
  |---|---|
  | `ZQC145` / `ZQC146` | Consumo Ativo em Ponta (kWh), TE / TUSD |
  | `ZQC147` / `ZQC148` | Consumo Ativo Fora Ponta Capacitivo, TE / TUSD |
  | `ZQC153` / `ZQC154` | Consumo (Convencional Ativo-kWh), TE / TUSD |
  | `ZQC163` / `ZQC164` | Desconto Baixa Renda Quilombola e Indígena, TE / TUSD |
  | `ZQC165` / `ZQC166` | Consumo Semáforo - Faturamento, TE / TUSD |
  | `ZQC167` / `ZQC168` | Consumo IP - Faturamento, TE / TUSD |
  | `ZQC218` / `ZQC219` | Consumo Intermediário, TE / TUSD |
  | `ZQC225` / `ZQC226` | Energia Compensada (Convencional), TE / TUSD |

- **Consumo do livre:** 2 códigos.

  | Código | Descrição no BW |
  |---|---|
  | `ZQC042` | TUSD Energia Fora Ponta |
  | `ZQC043` | TUSD Energia Ponta |

- **Demanda:** 12 códigos.

  | Código | Descrição no BW |
  |---|---|
  | `ZQD001` | Demanda em Ponta |
  | `ZQD003` | Demanda Faturada Convencional (kW) |
  | `ZQD004` | Demanda Fora de Ponta (kW) |
  | `ZQD017` | TUSD Demanda Fora Ponta |
  | `ZQD018` | TUSD Demanda Ponta |
  | `ZQD036` | DEMANDA/TUSD |
  | `ZQD058` | TUSD Dif. Demanda Reg. Contrat. Fora Ponta |
  | `ZQD059` | TUSD Dif. Demanda Reg. Contrat. Ponta |
  | `ZQD060` | Dif. Demanda Reg. Contrat. Fora Ponta |
  | `ZQD061` | Dif. Demanda Reg. Contrat. (Convencional - kW) |
  | `ZQD062` | Dif. Demanda Reg. Contrat. Ponta (kW) |
  | `ZQD063` | TUSD Dif. Demanda Reg. Contrat. |

  Os códigos de diferença de demanda contratada (`ZQD058` a `ZQD063`) cobram, pela tarifa de demanda,
  a parcela medida acima do contrato. São tarifa e entram no cálculo. A ultrapassagem, cobrada sobre
  os mesmos kW, é multa e fica fora (tabela abaixo).

O restante tem receita, mas não tem tarifa por item. Fica marcado como **não tarifável** e não
some: é gravado linha a linha numa aba própria, para que a soma da saída continue igual ao total do
BW e para que o conversor possa aproveitar o que interessa ao layout final.

| Lançamento | Por que fica fora | Destino |
|---|---|---|
| ICMS, PIS/COFINS (operação principal `R/#`) | Tributo, não tarifa | Colunas de tributos da Receita Faturada |
| Bandeiras (códigos `ZQDFAM`, `ZQDPAM`, `ZQDFVM`, `ZQDPVM`, `ZQODAM`, `ZQODVM`, `ZQDIAM`, `ZQDIVM`) | Adicional definido mês a mês, fora da tarifa homologada | Coluna Bandeira da Receita Faturada |
| Encargo da conta COVID (`ZQC237`) | Lançado sem quantidade | Grupo CDE Contas da Receita Faturada |
| Reativos; ultrapassagem de demanda: cativo (`ZQD021`, `ZQD022`, `ZQD023`) e livre (`ZQD019`, `ZQD020`, `ZQD037`) | Multa sobre a demanda acima da contratada, sem item na tabela homologada | Coluna UD/ER da Receita Faturada |
| Substituição tributária dos livres | Débito e crédito se anulam no mês | Coluna Outros da Receita Faturada (efeito zero) |
| "TUSD - Desconto Consumo" (`ZQC036`, `ZQC181`, `ZQC182`) e o estorno "Complemento Res. 166/05" (`ZQC059`, `ZQC183`, `ZQC184`) | Isenção do APE, sem quantidade | Usados no Passo 9 |
| Acertos homologados pela CCEE, energia compensada fora ponta e intermediário, desconto comercial | Sem tarifa por item | Coluna Outros da Receita Faturada |

Os tributos e as bandeiras são identificados pela regra da área de faturamento: operação principal
`R/#` para os tributos e os oito códigos acima para as bandeiras. Dentro dos tributos, o campo de
montante separa ICMS, PIS e COFINS por parcela (TE ou TUSD).

## Passo 3 — Montar a chave tarifária

A tarifa não é um número por cliente: é um número por combinação de características do
fornecimento. A chave junta oito atributos, sem separador:

```
SUBGRUPO + MODALIDADE + CLASSE + SUBCLASSE + DETALHE + UC + POSTO + UNIDADE
```

Onde não há valor, o texto é literalmente `Não se aplica`. Cada atributo é deduzido do que o BW
traz:

| Atributo | Como é deduzido |
|---|---|
| Subgrupo | Primeiro termo da categoria tarifária. B4a e B4b viram B4. |
| Modalidade | Azul ou Verde pela categoria. Branca quando a categoria termina em `_TB` e o código é de posto horário. Convencional no restante. |
| Classe | Decorre do subgrupo: B1 residencial, B2 rural, B4 iluminação pública. |
| Subclasse | Só em B1 e B4. Separa Baixa Renda, Desconto Social e residencial comum, e iluminação B4a ou B4b. |
| Detalhe e UC | Sempre `Não se aplica` no fornecimento: só existem nas geradoras. |
| Posto | Prefixo E (consumo) ou D (demanda), mais o posto lido do texto do lançamento: ponta, fora ponta, intermediário ou nenhum. |
| Unidade | MWh para consumo, kW para demanda. |

O faturamento por posto tarifário é comum tanto para AT quanto para BT (Tarifa Branca). Os códigos
são os mesmos para os dois grupos de clientes. Por isso eles definem o posto em qualquer grupo, e só
indicam a modalidade Branca junto com a categoria `_TB`.

No mesmo passo, a subclasse bruta do BW é classificada quanto aos **benefícios sociais**: Baixa
Renda (inclusive BPC, Indígena e Quilombola) ou Desconto Social. Essa marca é usada no Passo 7.

## Passo 4 — As tarifas homologadas

As tarifas vêm de uma tabela única, `tbl/tarifas/tarifa_long.csv`. Ela traz uma linha por vigência,
camada e chave, com as datas de início e fim de cada vigência.

- **Como a tabela é gerada.** O script `3-Py/1-Entrega/extrai_pcat.py` lê os PCATs publicados pela ANEEL.
  As colunas são localizadas pelo nome, nunca pela posição, então um item novo publicado num ano
  futuro entra sem alteração de código.
- **Camadas.** A camada financeira é calculada como TA − BE − CVA. O script confere que a identidade
  fecha em todas as vigências.
- **Incluir uma vigência nova.** Uma vez por ano, quando a ANEEL publica o PCAT do reajuste:
  1. Salvar o PCAT em `src/pcats/`.
  2. Em `3-Py/1-Entrega/extrai_pcat.py`, acrescentar uma linha à lista `VIGENCIAS`, com o ano, o arquivo e
     as datas de início e fim. Exemplo para o reajuste de 2027:

     ```python
     {"ano": 2027, "arquivo": BASE / "src" / "pcats" / "390 PCAT Enel SP 2027 <versão>.xlsx",
      "inicio": dt.date(2027, 7, 4), "fim": dt.date(2028, 7, 3)},
     ```

  3. Rodar `python 3-Py/1-Entrega/extrai_pcat.py`. O script regrava `tbl/tarifas/tarifa_long.csv` com todas
     as vigências e confere TA = BE + CVA + FIN em cada uma.

  Os notebooks do P × Q e da RNF não mudam: ambos leem a mesma tabela e escolhem a vigência pelas
  datas.

**Tarifa Social e Desconto Social são benefícios diferentes** (README 3.5), cada um com chaves
próprias na tabela.

- Os PCATs anteriores à lei do Desconto Social não trazem a tarifa desse benefício, mas o BW já o
  fatura.
- Para essas vigências, o script deriva a tarifa a partir da residencial, retirando os itens
  isentos (CDE, CDE COVID e CDE GD). As linhas derivadas ficam marcadas na coluna `ORIGEM`.
- É a mesma regra aplicada pelo Mercado Receita.

**TE do suprimento a distribuidoras** (chaves A4 Distribuição: Ceris, CPFL Piratininga e Neoenergia
Elektro). Até a vigência de 2025, o PCAT publica essas chaves com a TE zerada nas abas "TA - ...". A TE
do suprimento está só nas abas de TE, na linha "Energia suprimento permissionária" (Ceris) ou
"concessionária" (as demais), `<500GWh`, com as camadas separadas: `TE BE` (BE), `TE CVA` (CVA) e
`TE BF` (financeiros). O script completa a TE dessas chaves com essas linhas, item a item, e marca
`ORIGEM` = `TE_SUPRIMENTO`. As colunas são localizadas pelo rótulo do item, e o mapeamento foi
conferido contra uma chave comum (A4 Azul ponta), com diferença zero nas três camadas. Resultado: 267,15
R$/MWh em 2024 e 278,27 em 2025, a mesma tarifa que a planilha do SD cobra da Ceris. A partir de 2026 o
PCAT já traz a TE na própria chave.

O PCAT escreve o posto por extenso e o BW usa o código abreviado do SAP. Para casar os dois, a
mesma regra do Passo 3 é aplicada sobre a tabela de tarifas.

**Qual vigência vale em cada linha.** A tarifa muda no reajuste anual (4 de julho em SP). Pela regra
da ANEEL, quem tem o ciclo de leitura cruzando essa data paga parte pela tarifa antiga e parte pela
nova, na proporção dos dias de cada vigência (REN 1.000/2021, art. 173, § 2º; README 3.6). O peso
de cada vigência vem do ciclo de leitura do lote do cliente, e a tarifa da linha é a média
ponderada. Como a ponderação é linear, TA = BE + CVA + FIN continua exata.

**O calendário de leitura.**

- **Fonte.** O calendário de leitura da área (`ESCALA AT MT BTr e BT`), salvo em
  `src/calendario_leitura_<ano>.xlsx`, uma aba por mês (`Jan` a `Dez`).
- **Tabela.** O script `3-Py/1-Entrega/gera_calendario_leitura.py <ano>` grava
  `tbl/calendario/calendario_leitura.csv`, uma linha por mês, grupo e lote: data de leitura, dias de
  consumo e o ciclo (`INICIO_CICLO` = leitura − dias + 1; `FIM_CICLO` = leitura).
  - Carrega só as abas presentes no arquivo e substitui apenas esses meses. O calendário pode chegar
    incompleto e ser revisado ao longo do ano; basta rodar de novo.
  - Para com erro se uma aba presente estiver sem algum lote. Avisa quando um ciclo não começa no
    dia seguinte à leitura anterior.
- **Grupos do calendário.**

  | Grupo | Lotes | Data de leitura |
  |---|---|---|
  | `MT` (Média Tensão) | 1 a 17 | Coluna Data Leitura; todos os lotes têm a mesma data |
  | `BT` (Baixa Tensão) | 1 a 20 | Coluna Data Leitura do lote |
  | `TELE` (Telemedidos) | 40 e 42 | A "Leitura Real" do cabeçalho da seção: último dia do mês anterior. A coluna Data Leitura é a data de processamento |

  O Cash Power não é carregado: não tem faturamento no BW.
- **Ligação com o BW.** A tabela `tbl/calendario/de_para_lote.csv`, validada pela área do BW, dá o
  grupo e o nº do lote do calendário para cada lote do BW (campo `Conjunto de Contrato`).
  - O nome do lote é comparado em minúsculas: grafias que diferem só em maiúsculas são o mesmo lote.
  - Regra geral: o número no nome indica o lote, no nível de tensão do sufixo (`_bt`, `_mt`).
  - Sem número de lote (`cativos_at`, telemedição): vale a data única do grupo.
  - `MES_CIVIL` (`ip_avencado`, `temporário aes eletropaulo`): faturados no último dia do mês; o
    ciclo é o mês civil da referência.
  - `SEM_FATURAMENTO` (`cashpower`): sem tarifa proporcional.
  - Lote do BW que não está na tabela: fica com a tarifa do mês, sem proporcionalidade, e é listado
    na aba `Lotes_sem_calendario`. A tabela é mantida à mão, com uma linha por lote novo.
  - Tabela com lote repetido, ou grupo sem nº de lote com mais de um ciclo no mês: o notebook para
    com erro.
- **Peso de cada vigência.** Dias do ciclo em cada vigência ÷ dias do ciclo.
  - Só os ciclos que contêm o início de uma vigência misturam as duas; nos demais o peso é 100% de
    uma vigência e o cálculo é o de sempre.
  - O calendário só é lido quando uma vigência começa dentro do ciclo mais longo possível de uma
    leitura do mês (62 dias antes do início do mês). Nesses meses, se o calendário do mês não
    estiver na tabela, o notebook para com erro; nos demais, o calendário não é necessário.
- **Tarifa do mês.** Dias do mês em cada vigência ÷ dias do mês. Vale para:
  - refaturamento e faturamento manual (REFAT e CNR). O REFAT exigiria o calendário e a tarifa do
    mês original, que pode ser de anos atrás, para um efeito pequeno; o CNR não tem ciclo de leitura;
  - lotes sem faturamento (`cashpower`) e lotes fora do de-para;
  - o SD (Parte II), que não tem lote.
- **Da linha do BW à linha agregada.** A base é somada sem o lote (Passo 1). Para não perder o ciclo,
  a soma guarda, para cada vigência, a quantidade multiplicada pelo peso dela. O peso da linha
  agregada é essa soma dividida pela quantidade total: a média dos pesos ponderada pela quantidade,
  exata porque o valor é linear no peso. Linhas sem quantidade, ou com quantidades de sinais opostos
  que levem a média para fora de 0% a 100%, ficam com a tarifa do mês.
- **Onde o peso é usado.** No bloco 1 (Passo 7), na tarifa da faixa 01 do Desconto Social (Passo 8)
  e na diferença de tarifa livre − APE (Passo 9, peso médio do grupo). Se a chave não existe numa das
  vigências, o peso é redistribuído entre as vigências em que ela existe.

## Passo 5 — A quantidade

O consumo é convertido de kWh para MWh, porque a tarifa de energia é publicada em R$/MWh. A demanda
fica em kW.

Um comportamento do SAP muda o resultado: **a quantidade é lançada uma vez só**.

- **Cliente cativo.** O volume aparece apenas na linha lançada na parcela TE, e a receita é dividida
  entre essa linha e a da TUSD. A linha da TUSD vem com quantidade zero. Por isso o cálculo é sempre
  feito pelo par TUSD/TE, nunca pela linha isolada.
- **Cliente livre.** O espelho é exato: toda a quantidade, de consumo e de demanda, está do lado
  TUSD, e as colunas de TE são zero. Não é uma falha de extração: o cliente livre compra a energia
  de outro fornecedor e a distribuidora não fatura TE (README 3.2).

## Passo 6 — Quais parcelas incidem

A tabela de tarifas traz, para cada chave, os 37 itens: os da TUSD (rede) e os da TE (energia). Este
passo define quais deles são multiplicados pela quantidade de cada linha. Sem ele, toda quantidade
receberia a tarifa inteira, e o cliente livre, que não compra energia da distribuidora, receberia TE.

- **Consumo do cativo:** TUSD e TE.
- **Demanda:** só TUSD. Não existe TE sobre demanda.
- **Livre e híbrido:** só TUSD (README 3.2).

O SAP não repete a quantidade nas duas linhas (TUSD e TE) do mesmo consumo, para que a soma da coluna
dê o total de energia (Passo 5). Por isso a tarifa inteira do cativo, TUSD + TE, é aplicada à linha
que traz a quantidade.

Na saída, a coluna `Consumo TUSD (kWh)` reproduz o BW e fica zerada no cativo, embora a TUSD seja
cobrada. Para mostrar a energia sobre a qual cada parcela é devida, as abas de itens trazem no fim as
colunas `Energia com TUSD (kWh)` e `Energia com TE (kWh)`, calculadas por esta mesma regra: no cativo,
a energia entra nas duas; no livre e no híbrido, só em TUSD. No SD, cada linha traz a própria
quantidade por parcela. As colunas de consumo originais não mudam, porque o RNF, o SAMP e os checks
as leem como vêm do BW.

## Passo 7 — Multiplicar: o bloco 1

Para cada linha, cada um dos 37 itens e cada camada:

```
valor do item = quantidade × tarifa do item     (zero se a parcela do item não incide na linha)
```

Aqui entra a primeira regra de benefício social (README 3.5): nas subclasses Baixa Renda e Desconto
Social, os **itens isentos de encargos** recebem tarifa zero.

| Benefício | Itens isentos |
|---|---|
| Baixa Renda | Quota CDE, PROINFA e CDE COVID (TUSD); CDE GD (TE) |
| Desconto Social | Quota CDE e CDE COVID (TUSD); CDE GD (TE) |

Não é um desconto percentual: é a não incidência desses itens, como prevê o PRORET.

O valor calculado é a soma dos 37 itens. A tarifa total publicada vem arredondada em duas casas, por
isso pode diferir alguns milhares de reais no mês de quem calcule pela tarifa total.

## Passo 8 — Agrupar e estimar as faixas dos benefícios sociais

**Agrupamento.** A saída traz uma linha por **grupo**, não por linha do BW. O grupo junta os
lançamentos que descrevem o mesmo fornecimento: classe principal, categoria tarifária, operação,
mercado, subclasse, chave e tipo de faturamento. Como o tipo faz parte do grupo, FAT, REFAT e CNR
passam pelo bloco 2 separados, e cada um fecha na própria receita. Os códigos do par TUSD/TE aparecem concatenados na mesma célula, e a
coluna `Linhas de origem` guarda as linhas da extração que formam cada grupo.

O agrupamento é necessário por causa do par TUSD/TE do cativo (Passos 5 e 6). Separadas, nenhuma
das duas linhas pode ser comparada com a tarifa: a linha da TE traz a quantidade e recebe o bloco 1
inteiro (TUSD + TE), mas só a receita da TE; a linha da TUSD traz a receita da TUSD, mas quantidade
e bloco 1 zero. Juntas, a receita do grupo (TUSD + TE) e o bloco 1 do grupo falam da mesma coisa e
podem ser comparados.

**Faixas dos benefícios sociais** (README 3.5). Os dois benefícios dependem do consumo de cada
cliente: até 80 kWh para a Tarifa Social e até 120 kWh para o Desconto Social. O BW não traz o
consumo por cliente nem a faixa. A participação de cada faixa é estimada pela receita do grupo:

- `A` = valor do grupo se toda a energia estivesse na faixa 1;
- `B` = valor do grupo se toda a energia estivesse na faixa 2;
- participação da faixa 2 = `(receita − A) / (B − A)`, limitada entre 0% e 100%;
- o bloco 1 do grupo passa a ser `A + participação × (B − A)`, item a item e camada a camada.

O valor de `A` em cada benefício:

| Benefício | Faixa 1 (A) | Faixa 2 (B) |
|---|---|---|
| Tarifa Social (Baixa Renda) | zero, porque a faixa é isenta | tarifa baixa renda com a isenção de encargos |
| Desconto Social | tarifa do Desconto Social | tarifa residencial comum |

É o mesmo tratamento que o Mercado Receita dá à Tarifa Social: energia isenta abatida, não tarifa
reduzida. A participação estimada aparece na coluna `% energia na faixa 01`.

**Tarifa Social no bloco 1: a tarifa homologada sem a redução** (README 3.5). Nas linhas da Tarifa
Social (Baixa Renda, BPC, Indígena, Quilombola), o bloco 1 fica com o valor `B` sobre toda a energia:
a tarifa B1 subclasse Baixa Renda homologada, sem a redução de 100% até 80 kWh. A faixa gratuita
estimada (`B` − valor estimado) vira a primeira etapa do bloco 2 (Passo 9). O Desconto Social mantém
as faixas no bloco 1.

- A régua é a da ANEEL para o reembolso: a tarifa homologada da subclasse Baixa Renda, sem e com o
  benefício, sem tributos e sem bandeiras (PRORET 5.2, §129 e §130).
- O Desconto Social e a isenção de encargos (CDE, PROINFA, CDE COVID, CDE GD) ficam de fora: não são
  reembolsados. O consumo isento é retirado do mercado de rateio da cota da CDE, e o custo vai para os
  demais consumidores pela tarifa (PRORET 5.2, §81 e §91).
- Como a faixa é estimada pela receita, o valor da etapa Tarifa Social também é uma estimativa.
- O bloco 2 da Tarifa Social não muda com isso: na Tarifa Social a faixa gratuita vale zero, então o
  valor estimado é uma fração do valor `B`, com os mesmos itens na mesma proporção.

## Passo 9 — O bloco 2: aplicação dos descontos e ajuste de fechamento

A diferença entre o bloco 1 e a receita é tratada em quatro etapas, nesta ordem. As três primeiras
aplicam os descontos previstos na regulação, cada um só sobre os itens em que a regulação o faz
incidir. A quarta, a 2ª passada, é o único ajuste do processo.

**0. Tarifa Social** (README 3.5). Linhas da Tarifa Social. A faixa gratuita até 80 kWh estimada no
Passo 8 sai do bloco 1, item a item e camada a camada (coluna `Ajuste Tarifa Social (R$)`).

**1. Regime APE** (README 3.4). Linhas do grupo A no mercado livre.

- A isenção de encargos do autoprodutor por equiparação aparece no BW em linhas sem quantidade:
  "TUSD - Desconto Consumo" (`ZQC036`, `ZQC181`, `ZQC182`) e o estorno "Complemento Res. 166/05"
  (`ZQC059`, `ZQC183`, `ZQC184`).
- O valor líquido do mês é distribuído entre os grupos de energia TUSD do mesmo subgrupo,
  modalidade e posto (e, havendo, da mesma subclasse), na proporção da capacidade de isenção de cada
  grupo: quantidade × (tarifa do livre − tarifa do APE). O tipo de faturamento não entra nessa
  escolha: as linhas APE com X vão para grupos FAT e REFAT pela capacidade de cada um. Por isso o
  REFAT do layout difere do REFAT da base pelo APE marcado com X.
- O abatimento atinge só os itens da TUSD em que as duas tarifas diferem. Na prática, CDE e PROINFA.
- Quando o valor de um mês supera a capacidade de um grupo (crédito de meses anteriores), o
  excedente vai para os demais grupos do mesmo subgrupo e modalidade.
- Uma conferência garante que o abatimento nunca supera o desconto registrado.

**2. 1ª passada: desconto de fonte incentivada** (README 3.3). Clientes livres do grupo A.

- O que ainda separa o bloco 1 da receita é absorvido **somente por Fio A e Fio B**.
- O desconto é limitado entre 0% e 100%, porque acima disso deixaria de ser desconto.
- Encargos e perdas ficam intactos. Esse é o ponto central do método: um rateio uniforme cortaria
  também a CDE e o PROINFA, que não têm desconto.

**3. 2ª passada: ajuste de fechamento.** Todos os grupos.

- O que sobra é distribuído proporcionalmente entre todos os itens do grupo, para cima ou para
  baixo, até o total igualar a receita. A sobra vem de arredondamentos, refaturamentos, acertos
  pontuais e do limite de 100% da 1ª passada.
- No cativo, que não tem descontos a aplicar nas etapas anteriores, só a 2ª passada atua. Ela é pequena
  porque os benefícios sociais já foram tratados no bloco 1.
- A 2ª passada **não é um desconto**, e por isso o tamanho dela é o indicador de qualidade do
  processo. A saída mostra, por grupo, quanto veio da 1ª passada e quanto da 2ª.
- Com `APLICAR_2A_PASSADA` desligado, a sobra não é distribuída e sai linha a linha do BW nas abas
  `Diferencas_linha` e `Diferencas_grupo`, para investigação.

**Grupos sem cálculo.** Um grupo com receita e bloco 1 igual a zero não tem item para escalar. A
receita dele fica fora do bloco 2 e é informada como resíduo sem item tarifário, nunca escondida.

Como todas as operações são multiplicações aplicadas igualmente às quatro camadas, TA = BE + CVA +
FIN continua exata no bloco 2.

## Passo 10 — Conferências

Antes de usar a saída, todas as conferências precisam passar. O notebook informa o resultado no fim
da célula.

| Conferência | O que garante |
|---|---|
| Bloco 2 = receita, grupo a grupo (com a 2ª passada ligada) | A abertura fecha no faturado |
| Resíduo sem item tarifário | Mostra quanto ficou fora do bloco 2 |
| Abatimento APE ≤ desconto registrado | A isenção não cria desconto inexistente |
| BE + CVA + FIN = TA | A identidade tarifária se mantém no bloco 2 |
| Famílias = soma dos itens | Nenhum item ficou sem família |
| Tarifável + não tarifável = total do BW | Nada se perdeu na separação |

---

# Parte II — Geradoras, suprimentos e uso da distribuição (SD)

Esses clientes não passam pelo BW. São faturados numa planilha de controle própria, com uma aba por
cliente. A diferença em relação ao fornecimento é favorável: **a planilha informa a tarifa
efetivamente aplicada**, então não há vigência a estimar. O que falta é apenas abrir essa tarifa
nos 37 itens.

## Passo 11 — Classificação dos clientes

A classificação de cada cliente fica no próprio notebook, porque a aba da planilha que a traz chega
oculta e sem garantia de que venha sempre. Cliente novo que não esteja na lista é reportado, nunca
adivinhado.

| Tipo | Clientes |
|---|---|
| Geração | PCH Rasgão, EE Pedreira, EE Traição (EMAE), Petrobras, PCH Pirapora, UTE São João, Tecipar |
| Suprimento | Ceris |
| Uso da distribuição | Neoenergia Elektro, CPFL Piratininga |

## Passo 12 — Ler a planilha do SD

Cada aba tem um desenho próprio: colunas em posições diferentes, chave tarifária numa linha acima do
cabeçalho, tarifa impressa em umas e implícita em outras. O layout de cada aba fica declarado num
único lugar do código.

**Mês de referência.** É o mês de **emissão da fatura**, não o de consumo. Esses clientes faturam
no fim do mês, e a fatura sai no mês seguinte.

- Quando a aba não traz a data de emissão, usa-se o mês de consumo mais um.
- A Ceris é a exceção: a aba dela já traz o mês de competência, que é usado como está.

**PIS, COFINS e ultrapassagem ficam fora.** As abas trazem o encargo e os tributos separados, e a
receita usada é a líquida de tributos. A ultrapassagem não tem item na tarifa. As três rubricas são
lidas e reportadas, para que a conferência contra a aba consolidada as desconte de forma explícita.

## Passo 13 — Abrir a receita do SD

- **Clientes sem aba própria.** A receita dos geradores que não têm aba própria vem da aba
  consolidada.
- **Tarifa implícita.** Quando a aba não imprime a tarifa, ela é obtida pela receita dividida pela
  quantidade.
- **Abertura.** A receita é aberta na proporção de cada item dentro da tarifa homologada da mesma
  chave (Passo 4). O total fica exato, cada item vai para o seu lugar e o SD sai aberto nas quatro
  camadas.
- **Ceris.** O lado TE é aberto pela TE de suprimento da tabela homologada (Passo 4), nas quatro
  camadas. A aba `SD tarifas` só é usada se a tabela tiver o lado da linha zerado, o que não ocorre
  mais em 2024 a 2026.

## Passo 14 — Juntar BW e SD e conferir

O SD não passa pelo bloco 2: a tarifa aplicada já está na planilha, então o calculado é a própria
receita. As linhas entram com regime `direto` e diferença zero, identificadas na coluna `Source`.
As conferências do SD comparam a receita aberta com a aba consolidada da planilha, cliente a
cliente e no mesmo mês de referência.

## Passo 15 — Gerar a planilha

A saída `1-Entrega/pxq/pxq_<ref>.xlsx` traz:

| Aba | Conteúdo |
|---|---|
| `Leia-me` | O que contém cada aba do arquivo e quais scripts a usam |
| `Componentes` | Os 37 itens na tarifa de aplicação, blocos 1 e 2, mais os totais de BE, CVA e FIN do bloco 2. Colunas de ajuste: Tarifa Social, regime APE, 1ª e 2ª passadas. No fim, `Energia com TUSD (kWh)` e `Energia com TE (kWh)` (Passo 6) |
| `Componentes_BE` | Os mesmos itens na base econômica (é o que alimenta o layout final). No fim, os totais por camada dos dois blocos: `[2] Total TA/BE/CVA/FIN` e `[1] Total TA/BE/CVA/FIN` (estes usados pelo check tarifa homologada × tarifa realizada) |
| `Familias` | As sete famílias, nas quatro camadas, blocos 1 e 2 |
| `Nao_tarifaveis_linha` | Cada lançamento não tarifável do BW, linha a linha, com classe, mercado e tipo de faturamento |
| `Lotes_sem_calendario` | Lotes do BW com faturamento normal (FAT) fora do de-para do calendário de leitura: ficam com a tarifa do mês (Passo 4) |
| `Diferencas_linha`, `Diferencas_grupo` | Só com `APLICAR_2A_PASSADA` desligado |

Todas as abas trazem a coluna `Tipo de faturamento` (FAT, REFAT ou CNR; o SD é sempre FAT). Nas abas
de itens e famílias, os lançamentos fora da conta aparecem no fim, agrupados também por tipo de
faturamento e marcados em `Tarifável?`:

- **NÃO:** sem tarifa homologada.
- **SEM TARIFA:** lançamento de fornecimento cuja chave não existe na tabela. É o que merece
  investigação: ou o cadastro está errado, ou falta uma linha na tabela de tarifas.

---

## Passo 16 — Saída para a Regulação

O notebook grava também `1-Entrega/pxq_regulacao/PxQ_regulacao_<ref>.xlsx`, para a área de
Regulação, que precisa da abertura nas camadas da tarifa e não só na tarifa de aplicação.

- **Linhas:** as mesmas da aba `Componentes` (BW tarifável, não tarifáveis e SD), numa aba só.
- **Colunas, nesta ordem:** identificação do BW; quantidades (`Energia com TUSD (kWh)`, `Energia com
  TE (kWh)` e `Demanda (kW)`); receita; a entrega (bloco 2) item a item
  na base econômica (`BE <item>`), na CVA (`CVA <item>`) e nos financeiros (`FIN <item>`).
- **Fora:** bloco 1, colunas de ajuste e de regime, valor calculado e diferença, e as colunas de
  consumo como vêm do BW (`Consumo TUSD`, `Consumo TE`) e os totais, que confundiam a leitura: no
  cativo o BW lança a energia só na TE. Sem painéis
  congelados (como no pxq).
- **Conferências:** a receita é a mesma da aba `Componentes`, e BE + CVA + FIN reproduz a TA do bloco 2
  em cada linha. Uma linha aberta só na TA, sem camadas, fica com os itens zerados e o notebook avisa
  o valor; com a TE de suprimento na tabela (Passo 4), não há nenhuma em 2026.

---

# Depois do notebook

## Conversão para o layout Receita Faturada

O script `3-Py/1-Entrega/gera_receita_faturada.py` lê o pxq do mês e grava
`1-Entrega/receita_faturada/receita_faturada_<ref>.xlsx`, a partir do próprio modelo da área
(`tbl/modelo/De-Para_Modelo_PyC.xlsx`). Valores em R$ mil; só as linhas com valor. O layout é montado
pelo bloco 2; a função de conversão aceita o bloco como parâmetro, para que o check tarifa homologada ×
tarifa realizada use exatamente a mesma regra com o bloco 1.

- **Grupos tarifários (colunas I a X).** Base econômica do bloco 2, item a item, agrupada pelo
  de-para da aba `Instruções` do modelo. Se o modelo mudar o agrupamento, o conversor acompanha sem
  alteração de código.
- **Base Financeira e Base CVA.** Totais de FIN e de CVA do bloco 2.
- **Consumo.** MWh nas linhas de energia e MW nas de demanda.
- **Opção de compra.** Cativo; livre (o híbrido entra como livre); geração, suprimento e uso da
  distribuição para o SD. SCEE – GD1 e SCEE – GD2 não são preenchidos: o BW não identifica a instalação no SCEE, e o
  movimento desses clientes fica em cativo (README 6.1).
- **Tipo de faturamento.** O modelo tem uma grade de 77 linhas para cada tipo (FAT, REFAT, CNR),
  gravadas nessa ordem, pela coluna `Tipo de faturamento` do pxq (Passo 1). O SD é sempre FAT.
- **Classe.** A classe do layout correspondente à do BW. Residencial Baixa Renda reúne Baixa Renda,
  BPC, Indígena, Quilombola e Desconto Social.
  - Lançamentos com classe principal "Não atribuído" são classificados pela categoria tarifária.
    Categoria residencial vai para Residencial, ou para Residencial Baixa Renda quando todas as
    subclasses listadas no lançamento são de baixa renda ou desconto social.
- **Bandeira, PIS/COFINS e ICMS.** Saem da aba `Nao_tarifaveis_linha`, com a classe de cada
  lançamento.
  - A parcela TE vai para a linha de energia.
  - A parcela TUSD e os tributos sem parcela, inclusive subvenções, são repartidos entre demanda e
    energia na proporção da receita do bloco 2 da mesma opção, classe e tipo de faturamento. Se o
    tipo não tem bloco 2 nessa opção e classe, vale a proporção dos três tipos juntos.
- **Conta COVID.** Entra em CDE Contas.
- **UD/ER e Outros.** Duas colunas que não existem no modelo da área, inseridas antes de PIS/COFINS:
  - **UD/ER:** ultrapassagem de demanda e energia reativa (ultrapassagem, UFER, UFDR, demanda e
    energia reativa excedente);
  - **Outros:** os demais lançamentos sem tarifa (acertos da CCEE, energia compensada, atualização
    de montante, ERCAP, desconto ex-funcionário, encargo de conexão, substituição tributária) e o
    resíduo sem item tarifário do bloco 2. Na substituição tributária, débito e crédito se anulam
    dentro de cada classe;
  - a unidade segue o próprio lançamento: demanda, potência, kW, UFDR e conexão vão para a linha D;
    o restante, para a linha E.
- **Geração na linha de energia.** O modelo só tem a linha D para geração. O valor da energia vai
  para Outros da linha GERAÇÃO / D, e o MWh, que não pode ser somado ao MW, aparece em nota na aba
  `Conferencia`.
- **Aba `Lotes sem calendário`.** Cópia da aba `Lotes_sem_calendario` do pxq: lotes do BW com
  faturamento normal fora do de-para do calendário de leitura, que ficaram com a tarifa do mês
  (Passo 4).
- **Aba `Conferencia`.** Concilia: receita total do pxq = valor gravado no layout + o que ficou fora.
  Com as regras acima nada fica fora, e a diferença tem de ser zero. A aba traz também o total
  gravado por tipo de faturamento. O total gravado é a receita
  bruta do BW mais o SD, a mesma base da aba 3.2 do informe de gestão. A diferença para o informe
  vem só do SD (PIS/COFINS das geradoras e receita que falte na planilha do SD).

## Conferência anual contra o Mercado Receita

Uma vez por ano, quando sai o Mercado Receita do reajuste (README 5.2), os scripts de
`3-Py/2-Checks/` comparam o bloco 2 com ele. Para que a comparação seja justa, três regras:

- **Agregação do MR.** O MR é mais detalhado que o BW. Os tipos de mercado, as faixas dos benefícios
  e o detalhe (fonte incentivada, APE, SCEE) são somados, e "livre APE" entra em livre.
- **Subsídios do MR.** O MR guarda o subsídio fora dos itens e o calcula abatendo quantidade. O
  subsídio é devolvido aos itens pelo mesmo abatimento, linha a linha. Assim os dois lados carregam
  os descontos dentro dos itens.
- **Quantidade do cativo.** O BW lança a energia do cativo só no lado TE; para comparar, a energia
  TUSD do cativo é igual à TE.

A cobrança direta de ERCAP aos livres fica fora dos dois lados.

## Check do layout Receita Faturada contra o Mercado Receita

O script `3-Py/2-Checks/gera_receita_faturada_mr.py` monta o Mercado Receita (MR) no mesmo layout da
Receita Faturada e compara as duas versões célula a célula. Roda depois do conversor, porque lê o
`receita_faturada_<ref>.xlsx` já gerado.

```
python 3-Py/2-Checks/gera_receita_faturada_mr.py 2026-05            # um mês
python 3-Py/2-Checks/gera_receita_faturada_mr.py 2026-01 2026-08    # intervalo
python 3-Py/2-Checks/gera_receita_faturada_mr.py --todos            # todos os receita_faturada
python 3-Py/2-Checks/gera_receita_faturada_mr.py --todos --classes  # abertura por classe
```

Por padrão o check é aberto só por opção de compra × unidade (D/E), com as classes somadas na linha
"Todas as classes". A abertura por classe fica disponível pelo parâmetro `ABRIR_CLASSES` no topo do
script (padrão `False`) ou pela opção `--classes`.

Saída: `2-Checks/receita_faturada_MR/receita_faturada_MR_<ref>.xlsx`. Quando o script roda para
mais de um mês, grava também `consolidado_receita_faturada_MR_<ini>_a_<fim>.xlsx`: uma linha por mês
(soma de todas as linhas do layout), com as mesmas colunas, nas abas `PxQ`, `MR` e `check`
(fórmulas). A coluna "Obs. MR" marca o mês em que o MR é igual ao do mês anterior, ou seja,
projetado e sem valor de conferência.

| Aba | Conteúdo |
|---|---|
| `PxQ` | O receita_faturada do mês, por opção de compra × unidade (ou, com `--classes`, na grade completa do modelo) |
| `MR` | O Mercado Receita do mês, no mesmo layout e nas mesmas linhas |
| `check` | Fórmulas Excel PxQ − MR, célula a célula, com a linha de totais |
| `Conciliacao` | O que ficou fora do MR (ERC), o subsídio distribuído nos itens e as linhas repartidas entre demanda e energia |

**Premissas**

*Bases*

1. O lado PxQ é o `receita_faturada_<ref>.xlsx` já gerado, sem recálculo.
2. O lado MR é o `src/SP_mercado_receita.xlsx`, filtrado pelo `AnoMes` do mês.
3. O script aceita todos os meses do MR. Meses projetados no MR (junho/2026 repete maio) geram um
   check sem valor de conferência; a aba `Conciliacao` avisa.
4. Valores em R$ mil. Consumo em MWh na linha E e em MW na linha D (`D` do MR ÷ 1000). No cativo a
   energia é a `TE_E`; nas demais opções, a `TUSD_E`.

*Grade e classificação*

5. As abas `PxQ` e `MR` têm as mesmas linhas, inclusive as zeradas: por padrão, todas as
   combinações de opção de compra × unidade do modelo, com as classes somadas; com `--classes`, a
   grade completa do modelo (opção × classe × unidade). Por isso o check é uma subtração da mesma
   célula nas duas abas.
   - O padrão soma as classes porque o cadastro de classe do cliente livre é diferente no MR e no
     BW (sessions, P21).
6. Opção de compra: LIVRE APE entra como LIVRE e DISTRIBUIÇÃO como USO D. Os clientes GD (Sistema de
   Compensação) ficam em CATIVO, como no PxQ (README 6.1). O ERC fica fora, porque não passa pelo BW,
   e o valor aparece na aba `Conciliacao`.
7. Classe (só com `--classes`): a do próprio MR. A Residencial Baixa Renda é identificada pela subclasse (baixa renda,
   Tarifa Social ou Desconto Social). Geração, suprimento e uso da distribuição usam a própria opção
   como classe, como no PxQ.

*Componentes e camadas*

8. Os componentes do MR vão para as colunas do modelo pelo de-para da aba `Instruções`, com o
   vocabulário do MR: `DISTRIBUICAO` → Parcela B; `EnergiaRevenda` e `ANGRA` → Compra de Energia;
   perdas técnicas, não técnicas e de rede básica → Perdas; `TUST_RB`, `TUST_FR`, `TUSDG_T`,
   `TUSDG_ONS`, `TUST_ITAIPU` e `TUST_CI` → Rede Básica; `Itaipu` → Transporte Itaipu; `CUSD` e
   `Conexao_D` → CUSD; `Conexao_T` → Conexão T; `CDE` → CDE USO; CDE Covid (TUSD e TE) e `CDE GD` →
   CDE CONTAS; `Cde_Elet` → CDE ELET; `PROINFA`, `ESSERR`, `ONS`, `PeD`, `TFSEE` e `RI` nas colunas de
   mesmo nome.
9. `CCC`, `CFURH`, `LIMINAR_1` e `OUTROS` não têm coluna no modelo e vão para Outros.
10. As colunas `CVA_*` vão para Base CVA e as `FinCT_*` para Base Financeira. Os totais do MR
    (`CVA_Totaltusd`, `CVA_Totalte`, `FinCT_Geral`) ficam fora, para não contar em dobro.

*Subsídio*

11. O subsídio fica dentro dos itens, pelo abatimento de quantidade do próprio MR: `Daj/D` na
    demanda, `TUSD_Eaj/TUSD_E` na energia TUSD e `TE_Eaj/TE_E` na TE. É o mesmo tratamento do PxQ,
    que carrega os descontos dentro dos itens.
12. O que o abatimento não explica (a diferença de tarifa da Baixa Renda, principalmente) é
    distribuído proporcionalmente nos itens da própria linha, que passa a somar o `FinTotAj` do MR
    (receita depois dos subsídios). O valor distribuído aparece na aba `Conciliacao`.

*Demanda e energia*

13. Componentes TE vão para a linha E.
14. Componentes TUSD vão para a linha D quando a linha do MR só tem demanda, e para a linha E quando
    só tem energia.
15. Nas linhas com demanda e energia juntas (cativo Azul do grupo A, cerca de 1% do valor), cada
    componente BE é repartido pela tarifa homologada BE do próprio componente: demanda × tarifa em
    kW contra energia × tarifa em MWh, na vigência que contém o dia 15 do mês. CVA e FIN são
    repartidos pela tarifa BE Total TUSD.
16. O P&D do MR é uma coluna única para TUSD e TE. Na repartição das linhas mistas, a parte de
    energia considera as duas tarifas.

*Tributos e itens que o MR não tem*

17. `icms` vai para ICMS; `cofins` + `pispasep` vão para PIS/COFINS. A repartição entre D e E segue a
    proporção do valor da opção × classe, como no PxQ.
18. Bandeira e UD/ER ficam zerados no MR, porque ele não tem esses valores. Na comparação do total,
    essas colunas (e Outros) aparecem inteiras como diferença.
19. A geração na linha E, que não existe no modelo, vai para Outros da linha GERAÇÃO / D, como no
    PxQ. O MWh não é gravado.

*Lado PxQ (herdado do receita_faturada)*

20. As linhas FAT, REFAT e CNR do layout do PxQ são somadas por opção × classe × unidade; o MR não
    distingue tipo de faturamento. O híbrido entra como LIVRE.
21. A classe principal "Não atribuído" é classificada pela categoria tarifária.
22. UD/ER e Outros ficam antes de PIS/COFINS; a receita do pxq (BW + SD) está toda gravada no
    layout.

## Check tarifa homologada × tarifa realizada (descontos tarifários)

O script `3-Py/2-Checks/gera_receita_faturada_bl1_vs_bl2.py` monta o layout Receita Faturada duas
vezes, pelo bloco 1 (tarifa homologada sem os descontos; Passo 8) e pelo bloco 2 (tarifa realizada, a
entrega), e separa a diferença por origem. O objetivo é identificar os itens tarifários que recebem
descontos, inclusive o benefício da Tarifa Social reembolsado pela CDE.

```
python 3-Py/2-Checks/gera_receita_faturada_bl1_vs_bl2.py 2026-05            # um mês
python 3-Py/2-Checks/gera_receita_faturada_bl1_vs_bl2.py 2026-01 2026-08    # intervalo
python 3-Py/2-Checks/gera_receita_faturada_bl1_vs_bl2.py --todos            # todos os pxq
```

Saída: `2-Checks/receita_faturada_Bl1_vs_Bl2/receita_faturada_Bl1_vs_Bl2_<ref>.xlsx` e, com mais de
um mês, o consolidado com uma linha por mês. Os três tipos de faturamento são somados em cada linha
de opção × classe × unidade.

| Aba | Conteúdo |
|---|---|
| `tarifa_homologada` | Layout pelo bloco 1, com a observação de que a Tarifa Social está valorizada pela tarifa B1 subclasse Baixa Renda sem a redução |
| `tarifa_realizada` | Layout pelo bloco 2 (igual à entrega) |
| `diferenca` | Fórmulas `tarifa_homologada − tarifa_realizada`, célula a célula; positivo = desconto |
| `ta_somada` | Tarifa de aplicação somada (colunas I a Z), por opção × classe × unidade: tarifa homologada, tarifa realizada, diferença e a parte de cada origem |
| `desconto_tarifa_social_dmr` | Tarifa Social – DMR: a DMR faturada do mês, aberta pelos itens na proporção da etapa Tarifa Social |
| `scee_baixa_renda` | SCEE - Baixa Renda(\*): estimado − DMR; não é subsídio homologado (nota de rodapé na aba) |
| `desconto_ape` | Parte da diferença que vem da isenção de encargos do APE |
| `desconto_fonte_incentivada` | Parte que vem do desconto de fonte incentivada (1ª passada) |
| `ajuste_2a_passada` | Ajuste de fechamento; não é desconto |
| `conciliacao` | Tarifa Social + APE + fonte incentivada + 2ª passada = diferença, e conferência com os ajustes que o notebook grava em `Componentes` |

**Regras**

1. O layout de cada bloco é montado pela mesma função da entrega (`gera_receita_faturada.py`), com a
   mesma grade, de-para, classes e unidades. As linhas são todas as do modelo (opção × classe ×
   unidade), inclusive as zeradas.
2. Bandeira, UD/ER, Outros, tributos e consumo não dependem do bloco: são iguais nos dois lados e a
   diferença nessas colunas é zero.
3. A separação por origem segue a ordem em que o notebook aplica as etapas (Passos 8 e 9), linha a
   linha do pxq:
   - **Tarifa Social:** a etapa do bloco 2 (coluna `Ajuste Tarifa Social (R$)`), o valor da faixa
     gratuita até 80 kWh, em duas partes:
     - **Tarifa Social – DMR:** a DMR (Diferença Mensal de Receita) faturada do mês, lida de
       `src/dmr_tarifa_social.csv` (informe de gestão, aba 1.2_Clientes Baixa Renda, coluna
       "Subsídio Faturado – DMR"). É aberta pelos itens e camadas na proporção do estimado;
     - **SCEE - Baixa Renda(\*):** estimado − DMR; não é subsídio homologado com a DMR. Nos clientes
       Baixa Renda com geração distribuída, a energia compensada é faturada pela TE do SCEE, sem o
       componente de energia (REN 1.000/2021, art. 655-I, §1º, II, "b"), e o benefício da Tarifa
       Social também incide sobre ela (art. 655-G, §8º). O BW não identifica essa energia; valorizada
       pela TE comum, ela infla o estimado. A explicação vai em nota de rodapé nas abas
       `scee_baixa_renda`, `ta_somada` e `conciliacao` e no consolidado;
     - mês sem DMR no arquivo: a Tarifa Social fica toda na linha DMR, com o valor estimado, e o
       script avisa.
   - **2ª passada:** fator único sobre todos os itens da linha, f = bloco 2 ÷ (bloco 2 − ajuste da
     2ª passada). O valor antes da 2ª passada é o bloco 2 ÷ f, item a item e camada a camada.
   - **Fonte incentivada:** diferença entre o bloco 1 e o valor antes da 2ª passada nos itens de
     Fio A e Fio B, os únicos em que a 1ª passada atua.
   - **APE:** a mesma diferença nos demais itens (a isenção do APE atua nos encargos).
4. Base CVA e Base Financeira: a parte da 2ª passada é exata. Entre APE e fonte incentivada, a soma
   CVA + FIN de cada origem é exata; a divisão entre CVA e FIN segue a proporção da própria linha.
   Só é aproximada nas linhas que têm as duas etapas.
5. Geradoras, suprimentos e uso da distribuição (SD) não passam pelo bloco 2 e não têm origem
   atribuída.
6. O total de cada origem fecha com os ajustes gravados pelo notebook (aba `conciliacao`).

**Limites do check**

- O Desconto Social, a isenção de encargos da Baixa Renda e as tarifas próprias das subclasses
  (rural, irrigante) continuam dentro do bloco 1 e não aparecem na diferença. Os dois primeiros não
  são reembolsados pela CDE (Passo 8).
- A Tarifa Social depende da faixa de consumo estimada pela receita (Passo 8). A parte DMR é o valor
  oficial; SCEE - Baixa Renda(\*) é a diferença atribuída à energia compensada dos clientes Baixa
  Renda com geração distribuída, que o BW não identifica. A DMR precisa ser incluída todo mês em `src/dmr_tarifa_social.csv`.
- Nem todo desconto da diferença é bancado pela CDE (Conta de Desenvolvimento Energético). A Tarifa
  Social e o desconto de fonte incentivada são; a isenção do APE e o ajuste da 2ª passada precisam
  ser lidos à parte.

---

## Vocabulário

| Termo | Significado |
|---|---|
| TUSD | Tarifa de Uso do Sistema de Distribuição: o uso da rede |
| TE | Tarifa de Energia: a energia em si |
| TA | Tarifa de Aplicação: o total homologado, igual a BE + CVA + FIN |
| BE | Base Econômica: a tarifa sem CVA e sem financeiros |
| CVA | Conta de Compensação de Variação de Valores de Itens da Parcela A |
| FIN | Componentes financeiros da tarifa |
| Fio A | Uso de ativos de terceiros (transmissão, conexão, outras distribuidoras) |
| Fio B | Uso da rede da própria distribuidora |
| PCAT | Planilha de cálculo tarifário publicada pela ANEEL em cada reajuste |
| APE | Autoprodutor por equiparação |
| SCEE | Sistema de Compensação de Energia Elétrica (geração distribuída) |
| SD | Geradoras, suprimentos e uso da distribuição |
| BW | SAP Business Warehouse, origem da extração de receita |
| Mercado Receita (MR) | Base do simulador Sparta usada no processo de reajuste |
| Posto tarifário | Faixa horária: ponta, fora ponta, intermediário |
| Cliente cativo | Compra energia e rede da distribuidora: paga TUSD e TE |
| Cliente livre | Compra energia no mercado livre: paga só TUSD |
| Lote de leitura | Grupo de clientes lidos na mesma data; no BW, campo `Conjunto de Contrato` |
| Ciclo de leitura | Período de consumo entre duas leituras consecutivas do lote |
| Telemedidos | Clientes com leitura remota; no calendário, lotes 40 e 42 |
