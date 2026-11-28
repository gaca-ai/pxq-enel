# P × Q — Abertura do faturamento de energia por item tarifário (Enel CE e Enel RJ)

Ponto de entrada do P × Q do faturamento de Ceará (Enel CE) e Rio de Janeiro (Enel RJ). As duas
distribuidoras usam o mesmo sistema comercial e recebem o mesmo tratamento; diferenças entre elas, quando
houver, são indicadas no texto com `[CE]` ou `[RJ]`. O detalhamento de cada transformação está no
[procedimento](1-Entrega/procedimentos/PROCEDIMENTO_PXQ.md). Milestones, pendências e decisões ficam em
[`sessions.md`](sessions.md).

O contexto, o objetivo e o embasamento regulatório são os de SP (README de SP, seções 0 a 3, cópia em
`../docs/referencias/`) e não são repetidos aqui. Esta documentação registra só o que muda.

Valores monetários em milhões de reais (MR$), energia em MWh e demanda em MW.

---

## 0. O que muda em relação a SP

| Tema | SP | CE e RJ |
|---|---|---|
| Fontes do faturamento | Extração do SAP BW (csv por lançamento) + planilha do SD | Três bases do sistema comercial: **livres e geradoras** (SAMP LV, um registro por documento emitido); **cativos** (SAMP de documentos, tabela `zisu_t_samp_m_dt`, Parte II do notebook); **SD** (`[RJ]` um cliente, a Elektro; `[CE]` não há) |
| Granularidade | Lançamento do BW (código ZQC/ZQD), agregado | Documento (fatura), com a conta contrato; cada documento é desdobrado em uma linha por grandeza (energia por posto, demanda por posto) |
| Chave tarifária | Montada a partir de subgrupo, modalidade, classe, subclasse, detalhe e posto do BW | Montada a partir da categoria tarifária do documento (`Ctg.tar.`, ex.: `A4_LVVD`) e da coluna de cada grandeza (posto e unidade) |
| Acessantes com tarifa própria | Na planilha do SD | Na base de livres: geradoras A2 (TUSDg por usina) e outras distribuidoras (subclasse `REVENDA`, tarifa "Distribuição"); o acessante é identificado pela tarifa aplicada |
| Fonte incentivada | Desconto estimado (1ª passada) | Desconto aplicado na 1ª passada e conferido, documento a documento, contra o desconto informado na base (`SUBSCARFI`); percentual de cada cliente inferido pela tarifa de demanda |
| APE | Linhas de isenção do BW (regime APE) | Sem marcação na base; identificado pela tarifa aplicada na energia fora ponta, igual à tarifa APE do PCAT |
| Tarifa proporcional | Calendário de leitura por lote | Base sem data de leitura nem mês de referência: tarifa do mês contábil (vigências ponderadas pelos dias do mês) |
| Vigências tarifárias | Reajuste anual (julho), uma vigência por ano | `[RJ]` reajuste em 15/03. `[CE]` reajuste em 22/04 e, em 2026, recálculo pelo UBP (Uso do Bem Público) a partir de 26/08/2026, tratado como nova vigência. Tabela de tarifas comum a RJ e CE |
| Lista de exceções | — | Arquivo mensal com toda linha sem associação direta de tarifa e toda divergência de desconto, por conta contrato, para o faturamento indicar a tarifa correta; abas `Geradoras` (nome da usina) e `APE` (confirmação) |
| Saídas | `pxq_<ref>.xlsx` e `PxQ_regulacao_<ref>.xlsx` | Mesmas abas e colunas de SP, no arquivo `pxq_<UF>_<ref>.xlsx`; muda o conteúdo de três colunas (`Source` = SAMP; `Qtd. Cta.contrato` no lugar de `Linhas de origem`; `Flag Cli.Liv./ Cat.` com a classificação APE e fonte incentivada) e há uma aba a mais (`APE_Fonte_Incentivada`) |

## 1. Objetivo

Mesmo de SP (README de SP, seção 1): abertura mensal da receita de fornecimento por item tarifário, família
e camada (TA, BE, CVA, FIN), fechando com a receita faturada, entregue no layout de SP.

## 2. Estrutura do projeto

```
p_x_q/
├── common/                       código e tabelas comuns a RJ e CE
│   ├── 3-Py/extrai_pcat.py       PCATs -> tabela de tarifas (todas as distribuidoras e vigências)
│   └── tbl/tarifas/tarifa_long.csv
└── CE_RJ/
    ├── docs/
    │   ├── referencias/          documentação de SP usada como modelo
    │   └── especificacao/        especificação técnica enviada à TI (v1.0)
    ├── FAT/                      este projeto: faturamento
    │   ├── README.md             este documento
    │   ├── sessions.md           diário: milestones, pendências e decisões
    │   ├── 1-Entrega/
    │   │   ├── Abertura_PxQ.ipynb
    │   │   ├── procedimentos/PROCEDIMENTO_PXQ.md
    │   │   ├── pxq/              saída principal (pxq_<UF>_<ref>.xlsx)
    │   │   └── pxq_regulacao/    saída para a Regulação (PxQ_regulacao_<UF>_<ref>.xlsx)
    │   ├── 2-Checks/
    │   │   ├── Informe_Gestao/   Informe de Gestão do mês (conferência da energia)
    │   │   └── lista_excecoes/   lista de exceções do mês (lista_excecoes_<UF>_<ref>.xlsx)
    │   ├── 3-Py/README_py.md     inventário do código de SP reaproveitável
    │   ├── tbl/                  tabelas do processo
    │   └── src/                  arquivos de origem (fora do Git)
    │       ├── LV/               base SAMP LV do mês (livres e geradoras)
    │       ├── PCAT/CE, PCAT/RJ  PCATs e notas técnicas
    │       └── MR_Reajuste/      Mercado Receita dos reajustes
    └── RNF/                      Renda Não Faturada de CE e RJ (projeto próprio)
```

| Pasta / arquivo | Para que serve |
|---|---|
| `common/3-Py/extrai_pcat.py` | Lê os PCATs de RJ e CE e grava a tabela única de tarifas, uma linha por distribuidora, vigência, camada e chave |
| `common/tbl/tarifas/tarifa_long.csv` | Tabela de tarifas usada pelo notebook (e pela RNF de CE e RJ) |
| `1-Entrega/Abertura_PxQ.ipynb` | Abertura do faturamento: leitura, chave, tarifa, blocos 1 e 2, conferências e saídas |
| `1-Entrega/pxq/pxq_<UF>_<ref>.xlsx` | Saída principal, no layout de SP |
| `1-Entrega/pxq_regulacao/PxQ_regulacao_<UF>_<ref>.xlsx` | Entrega item a item nas camadas BE, CVA e FIN, no layout de SP |
| `3-Py/readme/tabelas_readme.ipynb` | Gera as tabelas da seção 5 do README (`2-Checks/tabelas_readme/tabelas_readme_<UF>.md` e `.xlsx`) |
| `2-Checks/lista_excecoes/lista_excecoes_<UF>_<ref>.xlsx` | Contas contrato sem associação direta de tarifa ou com divergência de desconto, e as geradoras e os APE a identificar, para o faturamento |
| `src/LV/<MM>.<AAAA>_SAMP_<UF> LV.XLSX` | Base de livres e geradoras do mês |
| `2-Checks/Informe_Gestao/Informe_de_Gestão_<AAAAMM>_(Enel Distribuição <UF>).xlsb` | Informe de Gestão do mês |

**Rotina mensal:**

1. Gravar a base SAMP LV do mês em `src/LV` e o Informe de Gestão em `2-Checks/Informe_Gestao`, com o nome
   padrão.
2. No notebook, ajustar `DISTRIBUIDORA` e `REFERENCIA` (Passo 0).
3. Rodar o notebook inteiro e verificar que o Passo 10 termina com "TODAS AS CONFERENCIAS OK".
4. Enviar a lista de exceções do mês ao faturamento.
5. A cada nova vigência tarifária: incluir o PCAT em `src/PCAT`, a vigência em `common/3-Py/extrai_pcat.py`
   e rodar o script.

Requisito do ambiente, além dos pacotes de SP: `pyxlsb` (leitura do Informe de Gestão).

### Versões

Nenhuma versão publicada.

## 3. Embasamento regulatório

O mesmo de SP (README de SP, seção 3). Abaixo, o que é específico de CE e RJ: os atos de cada vigência.

| Distribuidora | Vigência | PCAT | REH | Período |
|---|---|---|---|---|
| CE | 2025 | `39 PCAT Enel CE 2025 V02.xlsx` | REH 3.445/2025 | 22/04/2025 a 21/04/2026 |
| CE | 2026 reajuste | `39 PCAT Enel CE 2026 V02_reajuste.xlsx` | REH 3.574/2026 | 22/04/2026 a 25/08/2026 |
| CE | 2026 recálculo (UBP) | `39 PCAT Enel CE 2026 V02_recalculo.xlsx` | Despacho 3.324/2026 | 26/08/2026 a 21/04/2027 |
| RJ | 2025 | `383 PCAT Enel RJ 2025 V02.xlsx` | REH 3.435/2025 | 15/03/2025 a 14/03/2026 |
| RJ | 2026 | `383 PCAT Enel RJ 2026 V02.xlsx` | `[A PREENCHER]` (Nota Técnica 33/2026) | 15/03/2026 a 14/03/2027 |
| CE e RJ | 2024 | Não recebido | | Necessário para refaturamento com referência anterior às vigências de 2025 |

O desconto de fonte incentivada incide só sobre a TUSD Transporte (Fio A e Fio B), como em SP (PRORET
Submódulo 7.1, §38); por isso não há desconto na energia fora ponta, onde a tarifa não tem fio. A tarifa APE
é a TUSD sem os encargos setoriais.

## 4. Implementação

O notebook `Abertura_PxQ.ipynb` segue os passos de SP, com três passos novos (4b, 4c e 7b):

| Passo | O que faz |
|---|---|
| 0 | Parâmetros: distribuidora, mês, tolerâncias da conferência de desconto |
| 1 | Lê a base SAMP LV (só linhas com número de documento) e confere a energia com o Informe de Gestão |
| 2 | Desdobra cada documento em linhas por grandeza; separa tarifável (energia e demanda) de não tarifável (ultrapassagem, bandeira, ICMS, PIS, COFINS, Outros) |
| 3 | Monta a chave tarifária a partir da categoria tarifária e da coluna; classe pela subclasse |
| 4 | Tarifa do mês, da tabela comum de tarifas |
| 4b | Acessantes com tarifa própria (geradoras A2, distribuidoras) identificados pela tarifa aplicada |
| 4c | Classificação dos livres: APE e fonte incentivada (com o percentual) |
| 5 a 7 | Quantidade, parcelas que incidem (no livre, só TUSD) e bloco 1 |
| 7b | Associação direta da tarifa e conferência do desconto contra o `SUBSCARFI`: grava a lista de exceções |
| 8 | Agrupa as linhas, com as chaves de grupo de SP |
| 9 | Bloco 2: 1ª passada (desconto de fonte incentivada no Fio A e Fio B) e 2ª passada |
| 10 | Conferências |
| 15 e 16 | Gravam as saídas no layout de SP |

A Parte II (cativos, base SAMP de documentos) entra no mesmo notebook.

## 5. Resultados obtidos

### 5.1 Do valor pela tarifa ao valor faturado

`[RJ]` Livres e geradoras, outubro de 2026 (4.916 documentos, 2.539 contas contrato):

Tabelas geradas por `3-Py/readme/tabelas_readme.ipynb` (mesmo formato de SP, só com as linhas válidas para RJ e CE).

Valores absolutos:

| Item | Unidade | out |
|---|---|---|
| (a) Receita faturada (base tarifável) | R$ milhões | 107,1 |
| (b) Bloco 1 (tarifa cheia) | R$ milhões | 137,7 |
| (c) Diferença a explicar, (b) − (a) | R$ milhões | 30,7 |
| (d) Tarifa Social (faixa gratuita) | R$ milhões | 0,0 |
| (e) Isenção APE | R$ milhões | 0,0 |
| (f) Fonte incentivada (1ª passada) | R$ milhões | −32,1 |
| (g) 2ª passada, líquida | R$ milhões | 1,5 |
| (h) Resíduo sem item tarifário | R$ milhões | 0,0 |
| (i) 2ª passada, bruta | R$ milhões | 2,1 |

Em relação à receita faturada:

| Item | Unidade | out |
|---|---|---|
| (c) Diferença a explicar, (b) − (a) | % de (a) | 28,63 |
| (d) Tarifa Social (faixa gratuita) | % de (a) | 0,00 |
| (e) Isenção APE | % de (a) | 0,00 |
| (f) Fonte incentivada (1ª passada) | % de (a) | −30,00 |
| (g) 2ª passada, líquida | % de (a) | +1,36 |
| (h) Resíduo sem item tarifário | % de (a) | 0,00 |
| (i) 2ª passada, bruta | % de (a) | 1,97 |

Além da receita tarifável, a base traz 57,74 MR$ não tarifáveis (ICMS, PIS, COFINS, ultrapassagem e
Outros), num total faturado de 164,82 MR$. Nos livres não há Tarifa Social; a linha (d) passa a ter valor com os cativos (Parte II).
A isenção APE é zero porque a base não marca o APE: as contas APE são valorizadas pela tarifa do livre,
e a diferença para a tarifa APE vai para a 2ª passada (−0,32 MR$).

A 2ª passada concentra-se em dois pontos: as distribuidoras com receita sem quantidade (+1,78 MR$) e os
clientes APE valorizados pela tarifa do livre (−0,32 MR$). Nos clientes de fonte incentivada, ela é de
0,002 MR$. Não há resíduo não ajustável: o refaturamento é separado em estorno e conta nova (Passo 2), e
cada parte recebe o próprio desconto.

### 5.2 Validações

`[RJ]` outubro de 2026:

| Validação | Resultado |
|---|---|
| Bloco 2 = receita, grupo a grupo | Diferença zero |
| TA = BE + CVA + FIN | Diferença zero |
| Tarifável + não tarifável = total faturado da base | Diferença zero |
| Energia da base × Informe de Gestão (livres) | 359.390 × 361.861 MWh; a diferença é a energia faturada fora da base SAMP LV |
| Desconto calculado × desconto informado (`SUBSCARFI`) | 31,518 × 31,475 MR$; a diferença está localizada em 41 contas contrato, listadas na lista de exceções |
| Fonte incentivada pela tarifa × `SUBSCARFI` | As duas classificações coincidem em todas as contas |

### 5.3 Comparação com o Mercado Receita do reajuste

Ainda não realizada.

### 5.4 Tarifa homologada × tarifa realizada

Ainda não realizada.

## 6. Pendências

Ver [`sessions.md`](sessions.md).
