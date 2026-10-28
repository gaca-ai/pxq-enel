# Especificação — Nova Base RNF Enel SP

26/10/2026

## 1. Objetivo

Valorizar a Receita Não Faturada por item tarifário (TUSD, TE, demanda e bandeira), no mesmo padrão do PxQ do faturamento.

## 2. Regras de negócio

1. Demanda: não há demanda na RNF. A valorização é feita somente por energia (kWh), em TUSD e TE, mais a bandeira.
2. Postos tarifários: sem abertura entre ponta e fora ponta. A definir: usar a tarifa média entre P e FP obtida do faturamento realizado da mesma classe tarifária, ou somente a tarifa fora ponta.
3. Abrangência: Enel SP, mercados cativo e livre, AT (Alta Tensão, Grupo A) e BT (Baixa Tensão, Grupo B). Clientes livres não pagam TE.
4. AT: o faturamento é por mês civil, com deslocamento. O faturamento registrado em M é o consumo de M−1; a RNF de M é o consumo do mês M inteiro, que será faturado em M+1.
5. BT: a leitura é por lote de faturamento, ao longo do mês. A RNF de M é apenas o pro rata entre a data da última leitura e o último dia de M.
6. Provisão e estorno: a provisão de M−1 é estornada em M, quando o faturamento correspondente é registrado. Receita contábil de M = Faturado(M−1) − RNF(M−1) + RNF(M), sem duplicidade. Provisão e estorno devem vir identificados separadamente.
7. Conciliação AT: a RNF de M é faturada em M+1. Nos meses já fechados, cada instalação da RNF de M será comparada com o seu faturamento em M+1; a diferença mede o erro da estimativa.
8. Mês com reajuste tarifário: será utilizada a regra de tarifa proporcional baseada no calendário de leitura compartilhado pelo faturamento.
9. Bandeira tarifária: em coluna própria, sem tributos e nunca embutida no valor de energia.
10. Tributos: valores sem PIS, COFINS e ICMS, com os tributos em colunas separadas.
11. Ajustes manuais: se houver, identificados em linhas separadas, com descrição do motivo.
12. Totalização: a soma da base deve bater com o saldo contábil da RNF de cada competência.

## 3. Layout necessário

| Campo | Descrição | Obrigatório | Observação |
| --- | --- | --- | --- |
| `competencia` | Mês da RNF (AAAA-MM) | Sim | |
| `tipo_lancamento` | Provisão ou estorno | Sim | |
| `mercado` | Cativo ou livre | Sim | Livres não pagam TE |
| `tensao` | AT ou BT | Sim | Define o tratamento |
| `instalacao` | Código da instalação | Sim (AT) | |
| `subgrupo` | A1, A2, A3, A3a, A4, AS, B1, B2, B3, B4 | Sim | |
| `modalidade` | Azul, Verde, Convencional, Branca | Sim | |
| `classe_subclasse` | Classe e subclasse de consumo | Sim | Mesma codificação do faturamento |
| `detalhe_desconto` | Indicador de APE ou fonte incentivada | Desejável | Se não houver, o desconto será estimado pela receita da tarifa homologada vs receita realizada |
| `data_leitura` | Data da última leitura | Opcional | Sempre pode ser obtida pelo calendário de leitura a partir do lote |
| `lote_faturamento` | Lote de faturamento | Sim | Com o calendário de leitura, define a data da leitura, o pro rata e a tarifa proporcional |
| `kwh_estimado` | Energia provisionada (kWh) | Sim | |
| `valor_energia_rs` | Valor de energia sem tributos (R$) | Sim | |
| `valor_bandeira_rs` | Adicional de bandeira sem tributos (R$) | Sim | |
| `pis_rs` | PIS (R$) | Sim | |
| `cofins_rs` | COFINS (R$) | Sim | |
| `icms_rs` | ICMS (R$) | Sim | |
| `tarifa_aplicada` | Tarifa usada na estimativa (R$/MWh) | Desejável | Permite auditar a valorização |
| `ajuste_manual` | Indicador e motivo do ajuste | Sim, quando houver | |

## 4. Granularidade

- AT: uma linha por competência × tipo de lançamento × instalação, com os atributos tarifários do layout.
- BT: uma linha por competência × tipo de lançamento × chave tarifária (subgrupo, modalidade, classe/subclasse, detalhe de desconto) × lote de faturamento.

## 5. Formato e periodicidade

- Formato: .xlsx ou .csv, em layout tabular (uma linha por registro, sem células mescladas, subtotais ou cabeçalhos múltiplos).
- Carga inicial: competências de dez/2025 a set/2026, com provisão e estorno.
- Periodicidade: mensal, junto com o fechamento contábil.
- Prazo de entrega mensal: a definir.
- Nome do arquivo: RNF_SP_AAAA-MM.
