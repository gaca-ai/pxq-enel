# Procedimento P × Q (Enel CE e Enel RJ): como o faturamento vira abertura por item tarifário

> TEMPLATE criado em 26/11/2026 a partir do procedimento de SP (`../../../docs/referencias/`). Cada passo
> traz, em "Em SP", o que foi feito lá, e em "CE e RJ", o que se espera mudar com a base do sistema
> comercial (especificação v1.0). Tudo que está em "CE e RJ" é hipótese até a base chegar. Ao preencher,
> substituir os dois blocos pela regra final e remover esta nota.

Manual gerencial do processo de CE e RJ. Cada seção corresponde a um passo e explica o que ele faz e por
que faz. A ideia geral (bloco 1 = quantidade × tarifa homologada; bloco 2 = descontos regulatórios + 2ª
passada, fechando com a receita) é a de SP ("A ideia em uma página") e não é repetida.

---

## Passo 0 — Parâmetros

- Em SP: `REFERENCIA`, `APLICAR_2A_PASSADA`, `PROPRIO_FATURA_TUSD` e caminhos.
- CE e RJ: mais `DISTRIBUIDORA` (CE ou RJ). `[A PREENCHER]`

## Passo 1 — Ler a base e conferir o recebimento

- Em SP: csv do BW convertido em parquet (`prepara_base_bw.py`), conferido contra o próprio csv.
- CE e RJ: quatro arquivos por distribuidora e mês (BT FAT, BT REFAT, AT FAT, AT REFAT), conferidos contra
  o arquivo de controle (linhas, receita bruta, valor sem tributos, PIS/COFINS, ICMS, energia e demanda)
  antes de qualquer processamento. `[A PREENCHER]`

## Passo 2 — Separar o que é fornecimento (tarifável × não tarifável)

- Em SP: códigos do BW classificados em tarifáveis e não tarifáveis (bandeira, tributos, UD/ER, Outros,
  Covid Migrantes).
- CE e RJ: a separação já vem nos campos (`VALOR_TUSD*`, `VALOR_TE*` tarifáveis; `BANDEIRA_RS`, `UD_ER_RS`,
  `ERCAP_RS`, `OUTROS_RS`, `PIS_COFINS_RS`, `ICMS_RS` não tarifáveis). CIP, serviços, juros e multa e ICMS-ST
  ficam fora da base (especificação 2.5). `[A PREENCHER]`

## Passo 3 — Montar a chave tarifária

- Em SP: SUBGRUPO + MODALIDADE + CLASSE + SUBCLASSE + DETALHE + ACESSANTE + posto + UNIDADE.
- CE e RJ: `CODIGO_TARIFA` (tarifa efetivamente aplicada, pode diferir do subgrupo de cadastro) ligado à
  chave da tabela de tarifas pela tabela de códigos tarifários (Layout 5). `[A PREENCHER]`

## Passo 4 — As tarifas homologadas

- Em SP: `extrai_pcat.py` lê os PCATs e grava `tarifa_long.csv` (todas as vigências e camadas; Desconto
  Social derivado quando o PCAT não traz).
- CE e RJ: mesmo script, adaptado para os PCATs de CE e RJ. Tabela comum ao FAT e à RNF de CE e RJ
  `[A DEFINIR: local, ex. CE_RJ/tbl/tarifas]`. `[CE]` dois PCATs em 2026 (reajuste e recálculo). PCAT 2024
  não recebido (I1). `[A PREENCHER]`

## Passo 5 — A quantidade

- Em SP: energia (MWh) e demanda (kW) por lançamento; energia de TUSD e TE.
- CE e RJ: BT com `ENERGIA_TUSD_KWH` e `ENERGIA_TE_KWH`; AT por posto (P, FP) e demanda faturada sem
  ultrapassagem; energia já líquida da compensação do SCEE e com o custo de disponibilidade do BT
  (especificação 2.3). `[A PREENCHER]`

## Passo 6 — Quais parcelas incidem

- Em SP: TUSD e TE por mercado (cativo: as duas; livre: só TUSD), por posto e unidade.
- CE e RJ: `[A PREENCHER]`

## Passo 7 — Multiplicar: o bloco 1

- Em SP: quantidade × tarifa homologada, com tarifa proporcional pelo ciclo de leitura do lote; REFAT e
  CNR pela tarifa do mês.
- CE e RJ: BT pelo `LOTE_LEITURA` e calendário; AT pelas datas de leitura da fatura. `MES_REFERENCIA` define
  a tarifa (especificação 2.2). `[A PREENCHER]`

## Passo 8 — Agrupar e faixas dos benefícios sociais

- Em SP: faixas da Tarifa Social e do Desconto Social estimadas pela receita.
- CE e RJ: `FAIXA_CONSUMO` informada no BT (Tarifa Social F01 até 80 kWh; Desconto Social F01 até 120 kWh).
  `[A PREENCHER]`

## Passo 9 — O bloco 2: descontos e ajuste de fechamento

- Em SP: etapas Tarifa Social, APE, 1ª passada (fonte incentivada, regime fio), 2ª passada (fechamento);
  regimes uniforme, fio e proprio (`PROPRIO_FATURA_TUSD`).
- CE e RJ: o AT traz `DESC_FONTE_RS` por UC (percentual dedutível) e `MERCADO` = APE; o BT traz
  `DESC_SOCIAL_RS`. Os descontos podem deixar de ser estimados. `[A PREENCHER]`

## Passo 10 — Conferências

- Em SP: bloco 2 = receita por grupo; 1ª + 2ª passada = diferença total; total contra a receita do BW.
- CE e RJ: mais o arquivo de controle e a soma de conferência de cada linha (especificação 2.4).
  `[A PREENCHER]`

## Passo 11 — Classificação dos clientes

- Em SP: cativo, livre, SD; classes do layout.
- CE e RJ: `MERCADO` (CAT, LIV, HIB, APE, GER, SUP, USD) e classe/subclasse do sistema. Ver o de-para de
  subclasses já usado na RNF de CE e RJ (`../../RNF/docs/CL_SUB_TAR.xlsx`). `[A PREENCHER]`

## Passos 12 a 14 — Geradoras, suprimento e uso da distribuição

- Em SP: planilha do SD à parte, aberta e juntada ao BW.
- CE e RJ: no arquivo AT (`MERCADO` = GER, SUP, USD; `AGENTE`). Provavelmente sem parte separada.
  `[A PREENCHER]`

## Passo 15 — Gerar a planilha

- Em SP: `pxq_<ref>.xlsx` (Componentes, Componentes_BE, Familias, Nao_tarifaveis_linha, Leia-me).
- CE e RJ: `[A PREENCHER]` (mesmo formato, para reaproveitar os scripts de SP).

## Passo 16 — Saída para a Regulação

- Em SP: `PxQ_regulacao_<ref>.xlsx` (BE, CVA, FIN item a item).
- CE e RJ: `[A PREENCHER]`

---

## Conversão para o layout Receita Faturada

- Em SP: `gera_receita_faturada.py` com o modelo da área (`De-Para_Modelo_PyC_sugerido.xlsx`).
- CE e RJ: `[A PREENCHER]`

## Relatórios e checks derivados

Em SP: consolidado anual por família, receitas consolidadas (com subsídio e RNF), base longa em csv,
relatório de energia para BIP/FCT, check contra o Mercado Receita, check tarifa homologada × tarifa
realizada (Bl1 × Bl2), check de subsídios contra o faturamento. Quais serão replicados em CE e RJ:
`[A PREENCHER]`

## Vocabulário

O de SP vale aqui. Termos novos de CE e RJ: `[A PREENCHER]`
