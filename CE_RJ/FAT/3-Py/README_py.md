# 3-Py — inventário do código de SP reaproveitável em CE e RJ

Este projeto não duplica código. Os scripts abaixo ficam onde estão (`SP/BW/3-Py`, versão de referência:
commit `f928323`) e passam a ser compartilhados por SP, CE e RJ depois de generalizados (parametrizados pela
distribuidora). A estrutura final (pacote comum, pastas, configuração por distribuidora) será definida a partir
do código do app da TI, ao qual tudo será acoplado. Até lá, esta pasta só guarda o inventário.

Acoplamento comum a todos: `BASE` = variável `PXQ_BASE`, com padrão fixo em `SP\BW`; nomes de arquivos e
pastas de SP no código.

## 1. Reaproveitáveis com parametrização

| Script (em `SP/BW/3-Py`) | Função | O que depende de SP |
|---|---|---|
| `1-Entrega/extrai_pcat.py` | PCATs → `tarifa_long.csv` (todas as vigências e camadas; Desconto Social derivado) | Lista `PCATS` (arquivos e vigências); CE tem duas vigências em 2026 (reajuste e recálculo) |
| `1-Entrega/gera_calendario_leitura.py` | Calendário de leitura por lote → `calendario_leitura.csv` | Formato do calendário de SP; em CE e RJ o calendário vale só para o BT |
| `2-Checks/le_mr.py` | Leitura do Mercado Receita do mês | Nomes dos arquivos (`Mercado_Receita_Enel_SP_*`, `MR_SP_periodo_tarifario*`); formato do MR de CE e RJ a confirmar (`CE_RJ/docs/MR_Reajuste`) |
| `2-Checks/gera_de_para_regulacao_MR.py` | De-para item × camada → coluna do MR (dicionário `ITENS`, fonte das famílias) | Itens do PCAT de CE e RJ (conferir se são os mesmos 37) |

## 2. Reaproveitáveis se o pxq de CE e RJ tiver o formato do de SP

Dependem das abas `Componentes`, `Componentes_BE`, `Familias` e `Nao_tarifaveis_linha` do `pxq_<ref>.xlsx` e do
`PxQ_regulacao_<ref>.xlsx`. Manter esse formato é o que permite o reaproveitamento.

| Script | Função | O que depende de SP |
|---|---|---|
| `1-Entrega/gera_receita_faturada.py` | Conversor pxq → layout Receita Faturada (`converte`, base dos demais) | Modelo da área (`tbl/modelo/De-Para_Modelo_PyC_sugerido.xlsx`); classes e códigos do BW (ex.: `ZQC237`, ultrapassagem pelo texto do lançamento) |
| `2-Checks/gera_receita_faturada_bl1_vs_bl2.py` | Tarifa homologada × realizada; descontos por origem (Tarifa Social – DMR, fonte incentivada) | `src/dmr_tarifa_social.csv`; colunas de ajuste do notebook |
| `2-Checks/gera_receita_faturada_mr.py` | Mercado Receita no layout Receita Faturada e check célula a célula | `le_mr.py` |
| `2-Checks/check_mr_bloco2.py`, `check_mr_camadas.py`, `consolida_check_mr_bloco2.py` | Check contra o MR por opção × subgrupo e por família × camada | `le_mr.py` |
| `1-Entrega/receita_faturada_consolidado.py` | Consolidado anual por família | Estrutura de classes da explicação ACT de SP |
| `1-Entrega/receitas_consolidadas.py` | Template Receitas Consolidadas (faturamento + subsídio + RNF) | Template da área; RNF de SP |
| `1-Entrega/extrator_receitas_consolidadas.py` | Base longa (csv) para Power Query / BI | `gera_receita_faturada`, `gera_de_para_regulacao_MR` |
| `1-Entrega/relatorio_energia_bip_fct.py` | Relatório de energia realizada para BIP/FCT | `gera_receita_faturada`, `gera_receita_faturada_bl1_vs_bl2`, RNF de SP (`Componentes_unidade`) |

Dependências entre eles: `gera_receita_faturada` ← `gera_receita_faturada_bl1_vs_bl2`, `extrator_receitas_consolidadas`,
`relatorio_energia_bip_fct`; `le_mr` ← `check_mr_bloco2`, `check_mr_camadas`, `gera_de_para_regulacao_MR`;
`gera_de_para_regulacao_MR` ← `extrator_receitas_consolidadas`, `relatorio_energia_bip_fct`.

## 3. Lógica reaproveitável, leitura nova

| Peça | Situação |
|---|---|
| `SP/BW/1-Entrega/Abertura_PxQ.ipynb` | Bloco 1, bloco 2 (Tarifa Social, APE, 1ª e 2ª passada, regimes), conferências e gravação do pxq valem; a leitura (Passos 1 a 3) e a Parte II (SD) mudam com a base do sistema comercial |
| `SP/BW/3-Py/readme/tabelas_readme.ipynb` | Tabelas da seção 5 do README; reaproveitável depois do pxq |

## 4. Não reaproveitáveis

| Script | Motivo |
|---|---|
| `1-Entrega/prepara_base_bw.py` | Específico do csv do BW; CE e RJ terão um leitor próprio da base (especificação v1.0) |
| `2-Checks/check_subsidios_vs_basefaturamento.py` | Lê o relatório de subsídios do faturamento de SP; em CE e RJ os descontos vêm na própria base (`DESC_SOCIAL_RS`, `DESC_FONTE_RS`) |
