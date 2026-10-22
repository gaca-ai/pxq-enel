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
| **Bloco 2, ajustado** | o bloco 1 corrigido até fechar na receita faturada | É o produto final: a receita aberta por item |

O bloco 1 nunca fecha com a receita, porque a tabela de tarifas traz a tarifa cheia e muitos
clientes pagam com desconto. A essência do método está em **como** o bloco 2 corrige essa
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
| `ABA_BW`, `REFERENCIA` | Aba da extração e mês de referência (`aaaa-mm`) | Todo mês |
| `ARQ_BW`, `ARQ_SD` | Arquivos de origem em `src/` | Quando chegar um arquivo novo |
| `APLICAR_2A_PASSADA` | Ligado: entrega padrão, fecha na receita. Desligado: o bloco 2 para antes do fechamento e as diferenças saem linha a linha, em arquivo com sufixo `_sem_2a_passada` | Para investigar diferenças |
| `REGIME_APE` | Liga o tratamento da isenção do autoprodutor por equiparação (Passo 9) | Mantido ligado |
| `CALIBRAR_TARIFA_PROPORCIONAL`, `N_MESES_CALIBRACAO` | Estimativa do peso de cada vigência nos meses de reajuste, e quantos meses (1 a 12) formam a régua | Mantidos; o padrão é 3 meses |
| `DIAS_POR_SUBGRUPO` | Peso manual de cada vigência por subgrupo; prevalece sobre a estimativa | Só se o faturamento informar o calendário de leitura |

## Passo 1 — Ler a extração do BW

A extração "Receita de Fornecimento com subclasse" traz uma aba por mês, com cabeçalho na linha 40.
Cada linha é uma combinação de:

- categoria tarifária (que contém o subgrupo e a modalidade);
- código estatístico (`ZQC` para consumo, `ZQD` para demanda);
- operação e suboperação;
- classe e subclasse;
- mercado (cativo, livre ou híbrido).

A coluna de montante é a **receita registrada**. É ela que o resultado final precisa reproduzir.

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
| Reativos; ultrapassagem de demanda: cativo (`ZQD021`, `ZQD022`, `ZQD023`) e livre (`ZQD019`, `ZQD020`, `ZQD037`) | Multa sobre a demanda acima da contratada, sem item na tabela homologada | Fora do layout |
| Substituição tributária dos livres | Débito e crédito se anulam no mês | Fora do layout |
| "TUSD - Desconto Consumo" e "Complemento Res. 166/05" | Isenção do APE, sem quantidade | Usados no Passo 9 |
| Acertos homologados pela CCEE, energia compensada fora ponta e intermediário, desconto comercial | Sem tarifa por item | Fora do layout |

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

- **Como a tabela é gerada.** O script `2-Apoio/extrai_pcat.py` lê os PCATs publicados pela ANEEL.
  As colunas são localizadas pelo nome, nunca pela posição, então um item novo publicado num ano
  futuro entra sem ajuste de código.
- **Camadas.** A camada financeira é calculada como TA − BE − CVA. O script confere que a identidade
  fecha em todas as vigências.
- **Incluir uma vigência nova.** Uma vez por ano, quando a ANEEL publica o PCAT do reajuste:
  1. Salvar o PCAT em `src/pcats/`.
  2. Em `2-Apoio/extrai_pcat.py`, acrescentar uma linha à lista `VIGENCIAS`, com o ano, o arquivo e
     as datas de início e fim. Exemplo para o reajuste de 2027:

     ```python
     {"ano": 2027, "arquivo": BASE / "src" / "pcats" / "390 PCAT Enel SP 2027 <versão>.xlsx",
      "inicio": dt.date(2027, 7, 4), "fim": dt.date(2028, 7, 3)},
     ```

  3. Rodar `python 2-Apoio/extrai_pcat.py`. O script regrava `tbl/tarifas/tarifa_long.csv` com todas
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

O PCAT escreve o posto por extenso e o BW usa o código abreviado do SAP. Para casar os dois, a
mesma regra do Passo 3 é aplicada sobre a tabela de tarifas.

**Qual vigência vale no mês.** Cada vigência recebe um peso igual à fração de dias do mês que ela
cobre, e a tarifa do mês é a média ponderada. Como a ponderação é linear, TA = BE + CVA + FIN
continua exata. Nos meses de reajuste esse peso de calendário é apenas o ponto de partida; o
Passo 7b o substitui por uma estimativa.

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

## Passo 7b — Tarifa proporcional nos meses de reajuste

**O problema.** A tarifa muda no reajuste anual (4 de julho em SP). Pela regra da ANEEL, quem tem o
ciclo de leitura cruzando essa data paga parte pela tarifa antiga e parte pela nova, na proporção
dos dias de cada vigência (REN 1.000/2021, art. 173, § 2º; README 3.6).

- O BW não traz o calendário de leitura de cada cliente.
- O calendário do mês diz que, em julho, 90% dos dias já estão na tarifa nova.
- Mas a alta tensão é lida no início do mês: a conta de julho é, na prática, consumo de junho.
- Usar o calendário atribuiria tarifa nova a consumo faturado pela antiga. A diferença iria para o
  fechamento e distorceria a abertura.

**A ideia.** Usar a própria receita como termômetro.

- **A régua.** Nos meses sem reajuste, a receita real de cada subgrupo difere da conta tarifa ×
  quantidade por um fator estável. Esse fator é a régua do subgrupo.
- **A pergunta.** No mês do reajuste: qual mistura de tarifa antiga e nova faz a conta bater com a
  receita, descontada a régua?
- **A resposta** é o peso da tarifa nova naquele subgrupo.

**Como funciona.**

1. **Histórico.** Todo mês processado grava `tbl/calibracao/calibracao_<ref>.csv`, com a receita e
   o valor calculado por subgrupo e classe. Esses arquivos formam o histórico.
2. **Régua.** No mês do reajuste e no seguinte, a régua é a média dos últimos `N_MESES_CALIBRACAO`
   meses normais.
3. **Base usada.** A estimativa usa só clientes cativos sem benefício social. Neles a receita é
   tarifa × quantidade, sem desconto de rede nem isenção.
4. **Cálculo do peso.** Com `A` = valor do subgrupo 100% na tarifa antiga e `B` = 100% na nova, o
   peso da nova é `(receita / régua − A) / (B − A)`, limitado entre 0% e 100%.
5. **Aplicação.** O peso vale para todas as linhas do subgrupo (cativo, livre, benefício social e
   SD), e a tarifa e o bloco 1 são refeitos.
6. **Salvaguardas.** O subgrupo fica com o peso do calendário quando:
   - não há histórico;
   - o reajuste do subgrupo é menor que 2%, porque a conta divide por `B − A` e fica instável;
   - a receita cativa do subgrupo é menor que R$ 1 milhão no mês. Em subgrupos pequenos, `B − A`
     vale poucos milhares de reais, e um refaturamento ou acerto pontual do mesmo tamanho muda o
     peso estimado em dezenas de pontos percentuais. Como esses subgrupos pesam pouco no total, o
     peso do calendário é suficiente.

   Um peso informado em `DIAS_POR_SUBGRUPO` prevalece sobre o estimado.

**Como conferir.** A estimativa é feita por subgrupo, então a conferência é feita por classe, que é
uma visão independente dela. O resultado fica nas abas `Calibracao_subgrupo` e `Calibracao_classe`
da saída.

## Passo 8 — Agrupar e estimar as faixas dos benefícios sociais

**Agrupamento.** A saída traz uma linha por **grupo**, não por linha do BW. O grupo junta os
lançamentos que descrevem o mesmo fornecimento: classe principal, categoria tarifária, operação,
mercado, subclasse e chave. Os códigos do par TUSD/TE aparecem concatenados na mesma célula, e a
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

## Passo 9 — O ajuste: o bloco 2

A diferença entre o bloco 1 e a receita é corrigida em três etapas, nesta ordem. Cada etapa trata
uma causa e atua só sobre os itens que essa causa afeta.

**1. Regime APE** (README 3.4). Linhas do grupo A no mercado livre.

- A isenção de encargos do autoprodutor por equiparação aparece no BW em linhas sem quantidade:
  "TUSD - Desconto Consumo" e o estorno "Complemento Res. 166/05".
- O valor líquido do mês é distribuído entre os grupos de energia TUSD do mesmo subgrupo,
  modalidade e posto, na proporção da capacidade de isenção de cada grupo: quantidade × (tarifa do
  livre − tarifa do APE).
- O abatimento atinge só os itens da TUSD em que as duas tarifas diferem. Na prática, CDE e PROINFA.
- Quando o valor de um mês supera a capacidade de um grupo (crédito de meses anteriores), o
  excedente vai para os demais grupos do mesmo subgrupo e modalidade.
- Uma conferência garante que o abatimento nunca supera o desconto registrado.

**2. 1ª passada: desconto de fonte incentivada** (README 3.3). Clientes livres do grupo A.

- O que ainda separa o bloco 1 da receita é absorvido **somente por Fio A e Fio B**.
- O desconto é limitado entre 0% e 100%, porque acima disso deixaria de ser desconto.
- Encargos e perdas ficam intactos. Esse é o ponto central do método: um rateio uniforme cortaria
  também a CDE e o PROINFA, que não têm desconto.

**3. 2ª passada: fechamento.** Todos os grupos.

- O que sobra é distribuído proporcionalmente entre todos os itens do grupo, para cima ou para
  baixo, até o total igualar a receita. A sobra vem de arredondamentos, refaturamentos, acertos
  pontuais e do limite de 100% da 1ª passada.
- No cativo, que não passa pelas etapas anteriores, a 2ª passada é o único ajuste. Ela é pequena
  porque os benefícios sociais já foram tratados no bloco 1.
- A 2ª passada **não é um desconto**, e por isso o tamanho dela é o indicador de qualidade do
  processo. A saída mostra, por grupo, quanto veio da 1ª passada e quanto da 2ª.
- Com `APLICAR_2A_PASSADA` desligado, a sobra não é distribuída e sai linha a linha do BW nas abas
  `Diferencas_linha` e `Diferencas_grupo`, para investigação.

**Grupos sem cálculo.** Um grupo com receita e bloco 1 igual a zero não tem item para escalar. A
receita dele fica fora do bloco 2 e é informada como resíduo não ajustável, nunca escondida.

Como todas as operações são multiplicações aplicadas igualmente às quatro camadas, TA = BE + CVA +
FIN continua exata depois do ajuste.

## Passo 10 — Conferências

Antes de usar a saída, todas as conferências precisam passar. O notebook informa o resultado no fim
da célula.

| Conferência | O que garante |
|---|---|
| Bloco 2 = receita, grupo a grupo (com a 2ª passada ligada) | A abertura fecha no faturado |
| Resíduo não ajustável | Mostra quanto ficou fora do bloco 2 |
| Abatimento APE ≤ desconto registrado | A isenção não cria desconto inexistente |
| BE + CVA + FIN = TA | A identidade tarifária sobreviveu aos ajustes |
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
- **Ceris.** Na tabela homologada, a chave da Ceris tem o lado TE zerado. Para esse lado, a
  abertura usa a aba `SD tarifas` da planilha, que só tem a tarifa de aplicação, sem as camadas BE,
  CVA e FIN.

## Passo 14 — Juntar BW e SD e conferir

O SD não precisa de ajuste: a tarifa aplicada já está na planilha, então o calculado é a própria
receita. As linhas entram com regime `direto` e diferença zero, identificadas na coluna `Source`.
As conferências do SD comparam a receita aberta com a aba consolidada da planilha, cliente a
cliente e no mesmo mês de referência.

## Passo 15 — Gerar a planilha

A saída `1-Entrega/pxq/pxq_<ref>.xlsx` traz:

| Aba | Conteúdo |
|---|---|
| `Componentes` | Os 37 itens na tarifa de aplicação, blocos 1 e 2, mais os totais de BE, CVA e FIN do bloco 2 |
| `Componentes_BE` | Os mesmos itens na base econômica (é o que alimenta o layout final) |
| `Familias` | As sete famílias, nas quatro camadas, blocos 1 e 2 |
| `Nao_tarifaveis_linha` | Cada lançamento não tarifável do BW, linha a linha, com classe e mercado |
| `Calibracao_subgrupo`, `Calibracao_classe` | Só nos meses de reajuste: o peso estimado de cada vigência e a conferência por classe (Passo 7b) |
| `Diferencas_linha`, `Diferencas_grupo` | Só com `APLICAR_2A_PASSADA` desligado |

Nas abas de itens e famílias, os lançamentos fora da conta aparecem no fim, marcados em
`Tarifável?`:

- **NÃO:** sem tarifa homologada.
- **SEM TARIFA:** lançamento de fornecimento cuja chave não existe na tabela. É o que merece
  investigação: ou o cadastro está errado, ou falta uma linha na tabela de tarifas.

---

# Depois do notebook

## Conversão para o layout Receita Faturada

O script `2-Apoio/gera_receita_faturada.py` lê o pxq do mês e grava
`1-Entrega/receita_faturada/receita_faturada_<ref>.xlsx`, a partir do próprio modelo da área
(`aux_files/De-Para_Modelo_PyC.xlsx`). Valores em R$ mil; só as linhas com valor.

- **Grupos tarifários (colunas I a X).** Base econômica do bloco 2, item a item, agrupada pelo
  de-para da aba `Instruções` do modelo. Se o modelo mudar o agrupamento, o conversor acompanha sem
  alteração de código.
- **Base Financeira e Base CVA.** Totais de FIN e de CVA do bloco 2.
- **Consumo.** MWh nas linhas de energia e MW nas de demanda.
- **Opção de compra.** Cativo; livre (o híbrido entra como livre); geração, suprimento e uso da
  distribuição para o SD. Todo o movimento é `FAT`, porque a extração não distingue refaturamento.
- **Classe.** A classe do layout correspondente à do BW. Residencial Baixa Renda reúne Baixa Renda,
  BPC, Indígena, Quilombola e Desconto Social.
- **Bandeira, PIS/COFINS e ICMS.** Saem da aba `Nao_tarifaveis_linha`, com a classe de cada
  lançamento.
  - A parcela TE vai para a linha de energia.
  - A parcela TUSD e os tributos sem parcela, inclusive subvenções, são repartidos entre demanda e
    energia na proporção da receita do bloco 2 da mesma opção e classe.
- **Conta COVID.** Entra em CDE Contas.
- **Aba `Conferencia`.** Concilia: receita total do pxq = valor gravado no layout + o que ficou fora
  (substituição tributária, reativos, ultrapassagem, energia compensada, acertos da CCEE, resíduo
  não ajustável). A diferença tem de ser zero.

## Conferência anual contra o Mercado Receita

Uma vez por ano, quando sai o Mercado Receita do reajuste (README 5.2), os scripts de
`3-Checks/py/` comparam o bloco 2 com ele. Para que a comparação seja justa, três regras:

- **Agregação do MR.** O MR é mais detalhado que o BW. Os tipos de mercado, as faixas dos benefícios
  e o detalhe (fonte incentivada, APE, SCEE) são somados, e "livre APE" entra em livre.
- **Subsídios do MR.** O MR guarda o subsídio fora dos itens e o calcula abatendo quantidade. O
  subsídio é devolvido aos itens pelo mesmo abatimento, linha a linha. Assim os dois lados carregam
  os descontos dentro dos itens.
- **Quantidade do cativo.** O BW lança a energia do cativo só no lado TE; para comparar, a energia
  TUSD do cativo é igual à TE.

A cobrança direta de ERCAP aos livres fica fora dos dois lados.

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
