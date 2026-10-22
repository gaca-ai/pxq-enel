# P × Q — Abertura da receita de energia por item tarifário (Enel)

Código e documentação do projeto. Os dados de entrada (extrações do BW, relatórios contábeis,
Mercado Receita, PCATs) e as saídas ficam fora do repositório (`.gitignore`).

| Pasta | Conteúdo |
|---|---|
| [`SP/BW`](SP/BW/README.md) | Fase I — receita faturada (fornecimento, BW + SD): notebook, conversor para o layout Receita Faturada, checks contra o Mercado Receita |
| [`SP/RNF`](SP/RNF/README.md) | Fase II — Renda Não Faturada: notebook a partir do relatório contábil |
| [`SAMP`](SAMP/projeto_samp.md) | P × Q a partir da base SAMP |
| `tools/` | Utilitários do repositório |

A tabela de tarifas homologadas (`SP/BW/tbl/tarifas/tarifa_long.csv`) é versionada: vem dos PCATs
públicos da ANEEL e é a fonte única de tarifa das duas fases.

Os notebooks são versionados sem as saídas das células (filtro `nbstrip` em `.gitattributes`).
Em cada clone, ative o filtro uma vez:

```
git config filter.nbstrip.clean "python tools/nb_strip.py"
```
