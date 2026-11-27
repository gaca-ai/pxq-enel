# Procedimento P × Q (Enel CE e Enel RJ): como o faturamento vira abertura por item tarifário

Manual gerencial do notebook `1-Entrega/Abertura_PxQ.ipynb`. Cada seção corresponde a um passo e explica o
que ele faz e por que faz. A ideia geral (bloco 1 = quantidade × tarifa homologada; bloco 2 = descontos
regulatórios + 2ª passada, fechando com a receita) é a de SP ("A ideia em uma página") e não é repetida.
Diferenças entre as duas distribuidoras, quando houver, vêm marcadas com `[CE]` ou `[RJ]`.

Valores monetários em milhões de reais (MR$), energia em MWh e demanda em MW.

---

## Fontes

| Base | Conteúdo | Parte do notebook |
|---|---|---|
| SAMP LV | Clientes livres, geradoras e outras distribuidoras: um registro por documento emitido, com a conta contrato | Parte I |
| SAMP de documentos (`zisu_t_samp_m_dt`) | Clientes cativos: um registro por documento emitido | Parte II |
| SD | `[RJ]` suprimento à Elektro; `[CE]` não há | A definir |

As receitas das bases SAMP vêm sem tributos; ICMS, PIS e COFINS vêm em colunas próprias.

---

# Parte I — Clientes livres e geradoras (base SAMP LV)

## Passo 0 — Parâmetros

| Parâmetro | O que define |
|---|---|
| `DISTRIBUIDORA`, `REFERENCIA` | Distribuidora (RJ ou CE) e mês contábil; os únicos que mudam todo mês |
| `APLICAR_2A_PASSADA` | Como em SP: `True` fecha o bloco 2 na receita |
| `TOL_DESCONTO_PCT` | Conferência do desconto (Passo 7b): diferença mínima, em proporção do desconto calculado (0,5%) |
| `TOL_DESCONTO_MIN_AT_BRL`, `TOL_DESCONTO_MIN_BT_BRL` | Diferença mínima em R$ por documento, no grupo A (R$ 10) e na baixa tensão (Parte II) |

Os caminhos seguem o nome padrão de cada arquivo: base `src/LV/<MM>.<AAAA>_SAMP_<UF> LV.XLSX`, Informe de
Gestão em `2-Checks/Informe_Gestao`, tarifas em `common/tbl/tarifas/tarifa_long.csv`.

## Passo 1 — Ler a base e conferir o recebimento

- Só entram as linhas com número de documento (`Nº doc.`); linhas de totalização ou de conferência
  escritas na planilha ficam de fora.
- Cada linha guarda o número da linha no Excel, para rastrear qualquer valor até a origem.
- O tipo de faturamento vem do `Código Arquivo`: 05 = faturamento (FAT), 06 = refaturamento (REFAT).
  Código fora desse de-para interrompe o processamento.

**Conferência.** A energia faturada (ponta + fora ponta) é comparada com a linha "Consumo Faturado
(Livres)" do Informe de Gestão do mês (aba `2.1_Consumo Faturado`, coluna Total). A diferença é a energia
faturada fora da base SAMP LV (`[RJ]` suprimento à Elektro).

## Passo 2 — Separar o que é fornecimento (tarifável × não tarifável)

O documento traz as grandezas em colunas. Ele é desdobrado em uma linha por grandeza com quantidade ou
receita, o equivalente aos lançamentos do BW em SP:

| Quantidade | Receita | Grandeza | Posto |
|---|---|---|---|
| `CONS_PONTA_TUSD` | `RECEITA_PONTA` | energia (kWh) | ponta |
| `CONS_FPONTA_TUSD` | `RECEITA_FPONTA` | energia (kWh) | fora ponta |
| `CONS_INTERMED_TUSD`, `CONS_RESERVADO_TUSD`, `CONS_NSA_TUSD` | idem | energia (kWh) | intermediário, reservado, sem posto |
| `DEM_PONTA`, `DEM_FPONTA`, `DEM_NSA` | `RECEITA_DEM_*` | demanda (kW) | ponta, fora ponta, única |

Têm tarifa por item só a energia e a demanda. Ficam como **não tarifáveis**, linha a linha:

- ultrapassagem de demanda (`RECEITA_ULT_DEM_*`): penalidade, não item tarifário;
- bandeira, ICMS, PIS e COFINS;
- **Outros** = `TOTAL_FAT` − todas as colunas acima, documento a documento.

A soma das linhas reproduz o `TOTAL_FAT` da base, centavo a centavo.

## Passo 3 — Montar a chave tarifária

A chave é a de SP: `SUBGRUPO + MODALIDADE + CLASSE + SUBCLASSE + DETALHE + ACESSANTE + POSTO + UNIDADE`, sem
separador.

| Atributo | Regra |
|---|---|
| Subgrupo | Parte da `Ctg.tar.` antes do `_` (`A3A` → A3a) |
| Modalidade | Sufixo da `Ctg.tar.`: `LVAZ` → Azul, `LVVD` → Verde, `GER` → Geração |
| Classe, subclasse, detalhe | Não se aplica (o grupo A não diferencia) |
| Acessante | Não se aplica; geradoras A2 e distribuidoras recebem o acessante no Passo 4b |
| Posto | `E` (energia) ou `D` (demanda) + `P`, `FP`, `INT` ou vazio; geradoras têm demanda única (`D`) |
| Unidade | MWh (energia) ou kW (demanda) |

**Subclasse `REVENDA` = uso da distribuição.** São outras distribuidoras conectadas à rede. A categoria é
`A*_LVAZ`, mas a receita é a tarifa "Distribuição" do PCAT. Essas linhas recebem a modalidade
Distribuição, a energia sem posto (o PCAT publica uma tarifa única de energia para a Distribuição) e a
classe "Uso da Distribuição" (classe 09 do Informe de Gestão).

**Classe principal.** Vem da subclasse (`ZnTemp.`) pelo de-para `RNF/docs/CL_SUB_TAR.xlsx`, o mesmo da RNF de
CE e RJ. Subclasse sem de-para fica como "(sem de-para)" e é informada.

**Mercado (`Flag Cli.Liv./ Cat.`).** `LVAZ` e `LVVD` = Cliente Livre (detalhado no Passo 4c); `GER` =
GERAÇÃO; `REVENDA` = DISTRIBUIÇÃO.

## Passo 4 — As tarifas homologadas

Fonte: `common/tbl/tarifas/tarifa_long.csv`, gerada por `common/3-Py/extrai_pcat.py` (uma linha por
distribuidora, vigência, camada e chave). O posto do PCAT, escrito por extenso, é convertido para o código da
chave (E/D + P, FP, INT).

**Tarifa do mês.** A base não traz data de leitura nem mês de referência de cada documento. Todas as linhas
usam a tarifa do mês contábil: as vigências que tocam o mês são ponderadas pelos dias do mês em cada uma.
Como a ponderação é linear, TA = BE + CVA + FIN continua exata.

### A tabela de tarifas comum

O script `common/3-Py/extrai_pcat.py` é a versão comum do script de SP, com a mesma leitura (colunas pelo
nome, FIN = TA − BE − CVA, Desconto Social derivado da Residencial quando o PCAT não o publica). O que mudou:

- Os arquivos e as vigências são configurados por distribuidora, no próprio script.
- Cada vigência tem uma identificação própria (`VIGENCIA`). Isso permite duas vigências no mesmo ano, como no
  CE em 2026 (`2026_reajuste` e `2026_recalculo`).
- A saída ganhou as colunas `DISTRIBUIDORA` e `VIGENCIA`.
- TE de suprimento a outras distribuidoras (chaves A4 Distribuição com TE zero): a linha de suprimento das
  abas de TE é escolhida por distribuidora e acessante. `[RJ]` permissionárias (CERCI, Ceral Araruama,
  Ceres) pela linha de permissionária, faixa <500GWh; concessionárias (Energisa Minas Rio, Neoenergia
  Elektro) pela linha de concessionária, faixa "Não se aplica", a única publicada no PCAT 2026. `[CE]` o
  PCAT não tem linha de suprimento: o CE só presta uso da rede à Energisa PB e à Equatorial PI, e a TE dessas
  chaves fica zero.
- O item `TE_ANGRA`, novo em 2026, também é lido nas abas de TE de suprimento.

| Distribuidora | `VIGENCIA` | Início | Fim |
|---|---|---|---|
| RJ | 2025 | 15/03/2025 | 14/03/2026 |
| RJ | 2026 | 15/03/2026 | 14/03/2027 |
| CE | 2025 | 22/04/2025 | 21/04/2026 |
| CE | 2026_reajuste | 22/04/2026 | 25/08/2026 |
| CE | 2026_recalculo | 26/08/2026 | 21/04/2027 |

Conferências do script: TA = BE + CVA + FIN em todas as vigências; vigências contínuas, sem buraco nem
sobreposição. Rodado para SP, o script reproduz célula a célula a tabela de SP
(`SP/BW/tbl/tarifas/tarifa_long.csv`). O PCAT 2024 de RJ e CE não foi recebido: não há tarifa para
refaturamento com referência anterior a 15/03/2025 (RJ) e 22/04/2025 (CE).

**Desconto Social nas vigências de 2025.** Os PCATs de 2025 não publicam o Desconto Social, faturado desde
jan/2026. A tarifa é derivada da Residencial sem os itens isentos (CDE, CDE Covid, CDE GD), como em SP. A
regra reproduz a tarifa de aplicação do Mercado Receita do reajuste de 2026 (RJ, jan/2026: faixa 01
810,44 e faixa 02 925,39 R$/MWh; CE, jan e fev/2026: 645,10 e 710,07 R$/MWh). Nas vigências de 2026 vale a
tarifa publicada no PCAT.

## Passo 4b — Acessantes com tarifa própria

Algumas tarifas do PCAT são publicadas por acessante: a TUSDg de cada usina do A2 e a tarifa "Distribuição"
de cada distribuidora conectada. A base não traz o nome do acessante. A tarifa aplicada (receita ÷
quantidade) é comparada com a TUSD de cada acessante do mesmo subgrupo, posto e unidade, cheia ou com 50%
de desconto, e a linha recebe o acessante cuja tarifa coincide.

- Quando dois acessantes têm a mesma tarifa (caso comum na Distribuição), o valor da abertura é o mesmo com
  qualquer um deles; a linha recebe o primeiro em ordem alfabética.
- Linha sem quantidade recebe o acessante já identificado na mesma conta contrato.
- Linha sem coincidência fica sem tarifa.

É uma inferência: toda linha que passa por aqui vai para a lista de exceções (Passo 7b).

## Passo 4c — Classificação dos livres: APE e fonte incentivada

A base não marca o APE (autoprodutor por equiparação) nem o cliente de fonte incentivada. O resultado do
faturamento mostra os dois, conta contrato a conta contrato:

| Classificação | Regra | Por que é seguro |
|---|---|---|
| APE | A tarifa aplicada na energia fora ponta da conta coincide com a tarifa APE da chave (e difere da do livre) | A tarifa APE é a TUSD sem os encargos, um valor que não coincide com nenhum outro do PCAT; na energia fora ponta não há fio, então o desconto de fonte incentivada não a altera |
| Fonte incentivada | A tarifa de demanda aplicada na conta fica abaixo da TUSD cheia da chave (desconto > 2%) | O desconto incide no fio, e a demanda é toda fio; o percentual sai direto (50%, 100%) |

- A conta é avaliada com valores absolutos: num refaturamento, o estorno e a conta nova se anulariam na
  soma com sinal.
- A classificação de fonte incentivada é comparada com a presença de `SUBSCARFI` na conta; divergência é
  informada.
- Desconto fora de 0%, 50% ou 100% vai para a lista de exceções.

A classificação vai para a coluna `Flag Cli.Liv./ Cat.` (`Cliente Livre`, `Cliente Livre APE`, `Cliente
Livre Fonte Incentivada` ou `Cliente Livre APE Fonte Incentivada`) e para a aba `APE_Fonte_Incentivada` da
saída (Nº doc., Cta.contrato, classificação e percentual de desconto). A valorização continua pela chave da
base: as contas APE seguem na lista de exceções até o faturamento confirmar.

## Passo 5 — A quantidade

Energia de kWh para MWh (a tarifa de energia é publicada em R$/MWh); demanda em kW (a tarifa de demanda é
publicada em R$/kW). No livre, a quantidade e a receita vêm na mesma linha.

## Passo 6 — Quais parcelas incidem

| Mercado | Grandeza | TUSD | TE |
|---|---|:-:|:-:|
| Livre (todas as classificações) | energia e demanda | sim | não |
| Geração | demanda (TUSDg) | sim | não |
| Distribuição | energia e demanda | sim | não |

Famílias dos itens: as de SP (Fio A, Fio B, Encargos, Perdas, Energia, RI, Outros).

## Passo 7 — Multiplicar: o bloco 1

Para cada linha, cada um dos 37 itens e cada uma das 4 camadas: quantidade × tarifa do item, zero se a
parcela do item não incide. Linha cuja chave não existe na tabela de tarifas não é valorizada: vai para o
fim da saída como `SEM TARIFA`.

## Passo 7b — Associação direta da tarifa e lista de exceções

Uma linha tem **associação direta** quando a chave sai só da base (categoria tarifária + coluna) e a tarifa
aplicada é compatível com a tarifa dessa chave: receita ÷ quantidade entre a TUSD cheia e a TUSD com 100% de
desconto no fio. Toda linha fora disso vai para a lista de exceções do mês,
`2-Checks/lista_excecoes/lista_excecoes_<UF>_<ref>.xlsx`:

| Motivo | O que aconteceu | O processo segue com |
|---|---|---|
| `CHAVE SEM TARIFA` | A chave montada não existe no PCAT | A linha fora da conta (`SEM TARIFA`) |
| `RECLASSIFICADA COMO DISTRIBUIÇÃO` | Subclasse `REVENDA` lida como uso da distribuição | Inferência |
| `ACESSANTE INFERIDO PELA TARIFA` | Geradora A2 ou distribuidora: acessante escolhido pela tarifa aplicada | Inferência |
| `TARIFA APLICADA FORA DA FAIXA DA CHAVE` | A receita não é explicada pela tarifa da chave (ex.: APE) | A chave da base; a diferença vai para a 2ª passada |
| `RECEITA SEM QUANTIDADE` | Não há como conferir a tarifa | Resíduo não ajustável |
| `DESCONTO FORA DE 0%, 50% OU 100%` | Desconto de fonte incentivada atípico na demanda (Passo 4c) | A chave da base |
| `DESCONTO DIFERENTE DO SUBSCARFI` | O desconto pela tarifa difere do informado na base, num documento da conta | A chave da base |

**Conferência do desconto, documento a documento.** Para os livres (exceto APE): desconto pela tarifa =
quantidade × TUSD cheia da chave − receita, somado no documento, contra o `SUBSCARFI` do documento. O
documento é listado quando a diferença, em valor absoluto, supera os dois limites: `TOL_DESCONTO_PCT` do
desconto calculado e `TOL_DESCONTO_MIN_AT_BRL`. Documento sem desconto pela tarifa e com desconto informado
entra se superar o valor mínimo.

Abas do arquivo:

| Aba | Conteúdo |
|---|---|
| `Por_conta` | Uma linha por conta contrato: valores usados pelo processo (`Ctg.tar.`, `ZnTemp.`, `Detalhe`, `Acessante`), cada um seguido de uma coluna com o sufixo `correto` (fundo cinza claro) para o faturamento preencher; motivo, pista de chave compatível, documentos, receita e observação do faturamento. Coluna "correto" em branco = o valor usado está certo |
| `Detalhe_linhas` | Linha a linha: grandeza, quantidade, receita, tarifa aplicada, chave usada e faixa aceita |
| `Divergencia_desconto` | Contas com divergência de desconto: todos os documentos da conta (para ler o refaturamento com o estorno e a conta nova), com tipo de fatura, valor total, desconto pela tarifa, desconto informado e diferença. As três primeiras linhas descrevem a tolerância usada |

A coluna `Chave compatível (pista)` sugere outra chave do PCAT do mesmo subgrupo e posto cuja faixa explica
a tarifa aplicada (por exemplo, a tarifa APE). É uma pista para o faturamento: o processo não troca a chave
sozinho.

## Passo 8 — Agrupar

Uma linha da saída por grupo, com as chaves de grupo de SP: classe principal + categoria tarifária + operação
principal + mercado + classe/subclasse + chave + tipo de faturamento. Correspondência com a base:

| Coluna da saída | Origem na base SAMP LV |
|---|---|
| `Operação principal` | `Código Arquivo` |
| `Classe/Sub-classe` | `ZnTemp.` |
| `Grp.estatísts.quant.`, `Grp.estatísts.mont.` | Nomes das colunas de quantidade e de receita |
| `Qtd. Cta.contrato` | Contas contrato do grupo, contando todas as ocorrências (a soma bate com as listas da base) |
| `Flag Cli.Liv./ Cat.` | Mercado e classificação do Passo 4c |

Não há faixas de benefício social nos livres; as colunas da Tarifa Social ficam zeradas.

## Passo 9 — O bloco 2

| Etapa | Regra |
|---|---|
| APE | Sem linhas próprias na base: a etapa fica zerada (colunas mantidas pelo layout) |
| 1ª passada — desconto de fonte incentivada | Livres do grupo A e geradoras (regime `fio`): a diferença entre o bloco 1 e a receita é aplicada só no Fio A e no Fio B, com desconto limitado a [0%, 100%] |
| 2ª passada | O que sobra é distribuído proporcionalmente entre todos os itens do grupo, até o total igualar a receita. Distribuidoras (regime `uniforme`) só têm esta etapa |

O tamanho da 2ª passada é o indicador de qualidade do processo. A 1ª passada dos livres de fonte
incentivada é comparada, por categoria tarifária, com o `SUBSCARFI`.

## Passo 10 — Conferências

| Conferência | O que garante |
|---|---|
| Bloco 2 = receita, grupo a grupo | A abertura fecha no faturado |
| TA = BE + CVA + FIN no bloco 2 | As camadas continuam coerentes depois da 2ª passada |
| Tarifável + não tarifável = `TOTAL_FAT` da base | Nenhum real se perdeu no desdobramento |

Grupos com receita e bloco 1 zerado (receita sem quantidade) não têm item para receber a receita: ficam como
resíduo não ajustável, informado.

---

# Parte II — Clientes cativos (base SAMP de documentos)

A construir. Base: `zisu_t_samp_m_dt`, um registro por documento, com energia e receita por posto, demanda e
ultrapassagem, SCEE, bandeira e tributos. O subgrupo SAMP identifica as faixas da Tarifa Social (67: até
80 kWh; 68: acima de 80 kWh), classificando o documento inteiro.

---

## Passo 15 — Gerar a planilha

`1-Entrega/pxq/pxq_<UF>_<ref>.xlsx`, com as abas e colunas de SP:

| Aba | Conteúdo |
|---|---|
| `Leia-me` | O que cada aba contém e onde é usada |
| `Componentes` | Os 37 itens na TA, blocos 1 e 2, totais do bloco 2 por camada |
| `Familias` | As sete famílias, nas 4 camadas, blocos 1 e 2 |
| `Componentes_BE` | Os itens na base econômica, com os totais por camada dos dois blocos |
| `Nao_tarifaveis_linha` | Cada linha não tarifável da base |
| `Lotes_sem_calendario` | Mantida pelo layout de SP; vazia enquanto a base não trouxer o lote |
| `APE_Fonte_Incentivada` | Documentos dos livres classificados como APE e/ou fonte incentivada (Passo 4c) |

Diferenças de conteúdo em relação a SP: `Source` = SAMP (em SP: BW ou SD); `Qtd. Cta.contrato` no lugar de
`Linhas de origem`; `Flag Cli.Liv./ Cat.` com a classificação do Passo 4c, além de GERAÇÃO e DISTRIBUIÇÃO.

## Passo 16 — Saída para a Regulação

`1-Entrega/pxq_regulacao/PxQ_regulacao_<UF>_<ref>.xlsx`, uma aba, com as linhas e colunas de SP: identificação,
quantidades (energia com TUSD, energia com TE, demanda), receita e a entrega (bloco 2) item a item nas camadas
BE, CVA e FIN. A coluna `Covid Migrantes` é mantida pelo layout.

## Vocabulário

O de SP vale aqui. Termos de CE e RJ:

| Termo | Significado |
|---|---|
| SAMP LV | Extração do sistema comercial com os clientes livres, geradoras e distribuidoras, um registro por documento |
| `Ctg.tar.` | Categoria tarifária do documento (ex.: `A4_LVVD` = A4 livre Verde; `A3_GER` = geradora A3) |
| `ZnTemp.` | Subclasse do cliente no sistema comercial |
| `SUBSCARFI` | Desconto de fonte incentivada concedido no documento (carga) |
| `REVENDA` | Subclasse das outras distribuidoras conectadas à rede (uso da distribuição) |
| Associação direta | Tarifa obtida só pelos campos da base e compatível com a receita faturada |
| Lista de exceções | Arquivo mensal com as contas sem associação direta ou com divergência de desconto |
