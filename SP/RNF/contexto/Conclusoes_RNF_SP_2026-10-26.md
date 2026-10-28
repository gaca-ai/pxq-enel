# Conclusões — RNF Enel SP (26/10/2026)

## Escopo da RNF (Receita Não Faturada)

- Não há demanda (kW) na RNF. A valorização é feita só por energia (kWh): TUSD, TE e bandeira.
- Não há abertura entre ponta e fora ponta.
- O ZUBIL_146 foi descartado como base. Será solicitada uma nova planilha, com especificação própria.

## Lógica contábil

- AT (Alta Tensão, Grupo A): fatura por mês civil, com deslocamento. O faturamento registrado em M é o consumo de M−1, e a RNF de M é o consumo do mês M inteiro, faturado em M+1.
- BT (Baixa Tensão, Grupo B): leitura por lote de faturamento ao longo do mês. A RNF de M é só o pro rata entre a última leitura e o fim de M.
- Estorno: a provisão de M−1 é estornada em M, então não há duplicidade.
- Receita contábil de M = Faturado(M−1) − RNF(M−1) + RNF(M).
- Consequência para o AT: nos meses fechados, o faturamento de M+1 traz o realizado da RNF de M. A comparação por instalação mede o erro da estimativa.

## Regras definidas

- Mês com reajuste: aplica-se a tarifa proporcional pelo calendário de leitura compartilhado pelo faturamento.
- Bandeira: em coluna própria, sem tributos.
- Tributos: valores sem PIS, COFINS e ICMS, com os tributos em colunas separadas.
- Lote de faturamento: obrigatório. A data de leitura sai do calendário de leitura a partir do lote.
- Detalhe de desconto (APE / fonte incentivada): desejável. Se não vier, o desconto será estimado pela receita da tarifa homologada vs receita realizada.
- Ajustes manuais: se houver, em linhas separadas e com o motivo.
- Granularidade AT: por instalação, ou por chave tarifária como alternativa.
- Granularidade BT: por chave tarifária × lote de faturamento.
- Carga inicial: competências de dez/2025 a set/2026, com provisão e estorno.

## Pendências

- Definir a tarifa para o AT: média P/FP do faturamento realizado da mesma classe tarifária, ou somente a tarifa fora ponta.
- Definir o prazo de entrega mensal da base.
- Tirar "demanda" do Objetivo da especificação, que hoje contradiz a regra de que não há demanda na RNF.
- Para reavaliar o ZUBIL_146 com as novas regras, falta o fit-gap campo a campo (depende do arquivo).

## Entregáveis gerados

- Especificação da nova base em PT (documento editável e Especificacao_Base_RNF_SP.md).
- Versão em inglês: Specification_RNF_Base_SP.md.
