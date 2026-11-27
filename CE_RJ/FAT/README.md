# P × Q — Abertura do faturamento de energia por item tarifário (Enel CE e Enel RJ)

> TEMPLATE criado em 26/11/2026 a partir do README de SP (`docs/referencias/`). Cada seção marcada com
> `[A PREENCHER]` é completada à medida que o processo de CE e RJ for construído. Ao preencher, remover
> esta nota e os marcadores.

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
| Fonte do faturamento | Extração do SAP BW (csv por lançamento) + planilha do SD | Base do sistema comercial, conforme a especificação `docs/especificacao/especificacao_base_faturamento_RJ_CE.pdf` (v1.0, 03/11/2026). Atendimento da especificação: `[A PREENCHER]` |
| Granularidade | Lançamento do BW (código ZQC/ZQD), agregado | BT agregado por combinação de atributos; AT cliente a cliente (UC e fatura), se a especificação for atendida |
| Chave tarifária | Montada a partir de subgrupo, modalidade, classe, subclasse, detalhe e posto do BW | `CODIGO_TARIFA` do sistema + tabela de códigos tarifários (Layout 5 da especificação) `[A CONFIRMAR]` |
| Geradoras, suprimento e uso da distribuição | Planilha à parte (Parte II do notebook) | No próprio arquivo AT (`MERCADO` = GER, SUP, USD) `[A CONFIRMAR]` |
| Faixas dos benefícios sociais | Estimadas pela receita (Passo 8) | Informadas no BT (`FAIXA_CONSUMO`) `[A CONFIRMAR]` |
| Fonte incentivada | Só o desconto em R$, sem a UC | Desconto por UC (`DESC_FONTE_RS`), de onde se deduz o percentual `[A CONFIRMAR]` |
| APE | Identificado pelo regime de ajuste | `MERCADO` = APE `[A CONFIRMAR]` |
| SCEE | Códigos de energia compensada no BW | `GD` (NAO, GD1, GD2) e `ENERGIA_COMP*` `[A CONFIRMAR]` |
| Tarifa proporcional | Calendário de leitura por lote | BT: `LOTE_LEITURA` + calendário; AT: datas de leitura da própria fatura `[A CONFIRMAR]` |
| Vigências tarifárias | Reajuste anual (julho), uma vigência por ano | `[RJ]` reajuste em 15/03. `[CE]` reajuste em 22/04 e, em 2026, recálculo pelo UBP (Uso do Bem Público) a partir de 26/08/2026, tratado como nova vigência (com tarifa proporcional). Tabela de tarifas comum a RJ e CE (procedimento, Passo 4) |
| `[A PREENCHER]` | | |

## 1. Objetivo

Mesmo de SP (README de SP, seção 1): abertura mensal da receita de fornecimento por item tarifário, família
e camada (BE, CVA, FIN, TA), fechando com a receita contabilizada, entregue no layout Receita Faturada.
`[A PREENCHER: diferenças de escopo, se houver]`

## 2. Estrutura do projeto

```
CE_RJ/
├── docs/
│   ├── referencias/              documentação de SP usada como modelo (README, procedimento, sessions)
│   ├── especificacao/            especificação técnica da base de faturamento enviada à TI
│   ├── PCAT/CE, PCAT/RJ          PCATs de cada distribuidora
│   └── MR_Reajuste/              Mercado Receita dos reajustes
├── FAT/                          este projeto: faturamento
│   ├── README.md                 este documento
│   ├── sessions.md               diário: milestones, pendências e decisões
│   ├── 1-Entrega/
│   │   ├── procedimentos/PROCEDIMENTO_PXQ.md
│   │   └── [A PREENCHER: notebook e pastas de saída]
│   ├── 2-Checks/                 saídas das conferências
│   ├── 3-Py/README_py.md            inventário do código de SP reaproveitável (sem cópias; estrutura a definir com o app da TI)
│   ├── tbl/                      tabelas do processo (tarifas, de-paras, calendário)
│   └── src/                      arquivos de origem (fora do Git)
└── RNF/                          Renda Não Faturada de CE e RJ (projeto próprio)
```

| Pasta / arquivo | Para que serve |
|---|---|
| `[A PREENCHER]` | |

**Rotina mensal:** `[A PREENCHER]`

### Versões

| Tag | Versão |
|---|---|
| `[A PREENCHER]` | |

## 3. Embasamento regulatório

O mesmo de SP (README de SP, seção 3). Registrar aqui só o que for específico de CE ou RJ
(resoluções homologatórias, vigências, particularidades dos PCATs). `[A PREENCHER]`

| Distribuidora | Vigência | PCAT | REH | Observação |
|---|---|---|---|---|
| CE | 2025 | `39 PCAT Enel CE 2025 V02.xlsx` | REH 3.445/2025 | 22/04/2025 a 21/04/2026 |
| CE | 2026 reajuste | `39 PCAT Enel CE 2026 V02_reajuste.xlsx` | REH 3.574/2026 | 22/04/2026 a 25/08/2026 |
| CE | 2026 recálculo (UBP) | `39 PCAT Enel CE 2026 V02_recalculo.xlsx` | Despacho 3.324/2026 | 26/08/2026 a 21/04/2027 |
| RJ | 2025 | `383 PCAT Enel RJ 2025 V02.xlsx` | REH 3.435/2025 | 15/03/2025 a 14/03/2026 |
| RJ | 2026 | `383 PCAT Enel RJ 2026 V02.xlsx` | `[A PREENCHER]` (Nota Técnica 33/2026) | 15/03/2026 a 14/03/2027 |
| CE e RJ | 2024 | Não recebido (a especificação pede 2024 a 2026) | | Necessário para REFAT antigo |

## 4. Implementação

`[A PREENCHER]` Seguir a ordem do procedimento: leitura e conferência da base (arquivo de controle),
chave tarifária, tabela de tarifas, quantidade, bloco 1, bloco 2, conferências, saídas.

## 5. Resultados obtidos

### 5.1 Do valor pela tarifa ao valor faturado
`[A PREENCHER]`

### 5.2 Validações
`[A PREENCHER]` Contra o arquivo de controle, o informe de gestão e a contabilidade.

### 5.3 Comparação com o Mercado Receita do reajuste
`[A PREENCHER]`

### 5.4 Tarifa homologada × tarifa realizada
`[A PREENCHER]`

## 6. Pendências

Ver [`sessions.md`](sessions.md).
