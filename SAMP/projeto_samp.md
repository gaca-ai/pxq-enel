# Projeto PxQ SAMP — Abertura do faturamento a partir do SAMP da ANEEL

Documento de contexto da frente SAMP, iniciada em 27/09/2026. Complementa o
`SP/BW/README.md` (motor do PxQ, conceitos regulatórios, validações do BW), que
continua sendo a referência do cálculo. Aqui ficam só o que é próprio do SAMP.

---

## 1. Objetivo

Reconstruir o PxQ (abertura do faturamento por item tarifário e por família)
usando o SAMP como fonte de quantidade e receita, para ver se o resultado
melhora — principalmente nos ENCARGOS dos clientes livres, onde o PxQ pelo BW
fica +R$ 13,5 mi (+10,7%) acima do MR em mar/26 (`SP/BW/README.md` 4.10).

Arquitetura de fontes pretendida:

| Período | Fonte |
|---|---|
| Meses com SAMP publicado | SAMP |
| Meses sem SAMP ainda | BW para os cativos (onde o PxQ já está aderente) + Mercado Receita para os livres |

O teste do modelo híbrido (BW + MR) é uma fase posterior — não iniciada.

---

## 2. Fonte: SAMP

SAMP = Sistema de Acompanhamento de Informações de Mercado para Regulação
Econômica (REN ANEEL 1.003/2022). Dados abertos da ANEEL, mercado faturado
declarado pela distribuidora, mensal.

- Arquivo: `SAMP/data/samp-2026.parquet` (todas as distribuidoras, 2026);
  dicionário em `SAMP/data/dd-samp.pdf`.
- Enel SP = CNPJ `61695227000193` (sigla ELETROPAULO). Em 2026 há competências
  de jan a jul (gerado em 15/09/2026).
- **Atenção à leitura**: `NumCNPJAgenteDistribuidora` vem como inteiro. O
  filtro tem de comparar como texto de 14 dígitos
  (`astype(str).str.zfill(14)`). O `note/samp_sp.ipynb` compara com a string
  direto e gera o `samp-SP-2026.parquet` vazio — não usar esse arquivo.
- Formato longo: uma linha por atributo x métrica (`DscDetalheMercado`,
  24 métricas) e valor (`VlrMercado`).
- Atributos equivalentes aos do MR: `NomTipoMercado`, `DscOpcaoEnergia`,
  `DscSubGrupoTarifario`, `DscModalidadeTarifaria`, `DscClasseConsumoMercado`,
  `DscSubClasseConsumidor`, `DscDetalheConsumidor`, `DscPostoTarifario`,
  `NomAgenteAcessante`.
- Data: `DatCompetencia`. Decisão do usuário: usar a competência como a data
  das demais bases (o BW não discrimina refaturamento).

O que o SAMP traz e o BW não:
- `DscDetalheConsumidor` separa os livres em **APE**, **Fonte incentivada**,
  **ERC** e Não se aplica.
- Faixas dos benefícios sociais declaradas (Tarifa Social faixa 01/02,
  Desconto Social faixa 01/02, baixa renda legado faixa 01 a 04).
- Percentual de desconto declarado (`Desconto Demanda %`, `Desconto Energia
  TUSD %`) por linha de fonte incentivada e geradora.
- Tipo de mercado (Regular, Refaturamento, SCEE GD I/GD II).

O que o SAMP não traz: receita separada TUSD x TE (só `Receita Energia` e
`Receita Demanda`), nem receita por componente.

---

## 3. Teste 1 — SAMP x Mercado Receita (mar/26)

Script `SAMP/model/check_samp_mr.py` (`python check_samp_mr.py 2026-03`),
saída `SAMP/saida/check_samp_mr_2026-03.xlsx`. Grão: tipo de mercado x opção x
subgrupo x modalidade x classe x subclasse x detalhe x posto x agente.

Normalizações MR -> SAMP: `LIVRE APE` -> LIVRE (APE fica no detalhe); `ERC` ->
LIVRE com detalhe ERC; "Serviço Público" -> "Serviço público"; nomes de agente
mapeados (`AGENTE_MR` no script).

**Resultado**
- Demanda, energia TUSD e TE, UC, PIS/COFINS e ICMS: **idênticos**, em todas
  as opções.
- Percentuais de desconto de fonte incentivada e de geradoras: **idênticos**
  linha a linha (SAMP declarado = 1 - `Daj`/`D` e 1 - `TUSD_Eaj`/`TUSD_E` do MR).
- Única diferença de forma: o MR cria a linha "Sistema de Compensação GD II -
  SCEE" (13.677 MWh, R$ 1,67 mi), que no SAMP é "Energia compensada" e
  "Receita energia compensada" dentro das linhas GD II. Somando, fecha.
- Conclusão: o MR e o SAMP usam a mesma base de mercado (esperado — a mesma
  área prepara os dois). Nas quantidades, o MR está correto.

**Receita** — SAMP (faturado: Receita Energia + Receita Demanda + Receita
energia compensada) x MR `FinTotAj` (calculado: tarifa de aplicação líquida de
subsídio):

| Opção | SAMP | MR | Delta |
|---|---|---|---|
| Cativo | R$ 1.584,19 mi | R$ 1.574,34 mi | +R$ 9,85 mi (+0,63%) |
| Livre (inclui APE e ERC) | R$ 231,27 mi | R$ 231,52 mi | -R$ 0,25 mi (-0,11%) |
| Distribuição, Geração, Suprimento | | | 0,00 |

- Cativo: B1 Residencial Convencional +R$ 7,6 mi (+0,8%); A4 Verde +R$ 0,92
  mi; A2 Azul +R$ 0,65 mi (+39%); B3 +R$ 0,63 mi. Causa não investigada
  (hipótese: receita faturada sem contrapartida de quantidade x tarifa, como
  custo de disponibilidade).
- Livre, duas diferenças que se compensam: Verde fonte incentivada R$ 1,35 mi
  abaixo do MR (A4 -1,14; A3a -0,21) — o MR só dá desconto na energia de ponta
  da Verde, e o faturado desconta mais (liga com as linhas "TUSD - Desconto
  Consumo" ZQC181/182 do BW, `SP/BW/README.md` 4.10); refaturamento do APE
  R$ 1,10 mi acima do MR (R$ 1,22 mi x R$ 0,12 mi).
- Memos só do SAMP (fora da comparação): ultrapassagem de demanda R$ 6,19 mi;
  bandeiras R$ 0,21 mi.

---

## 4. Teste 2 — PxQ SAMP (quantidade SAMP x tarifa PCAT) x Mercado Receita (mar/26)

Script `SAMP/model/pxq_samp_check_mr.py` (`python pxq_samp_check_mr.py
2026-03`), saída `SAMP/saida/check_pxq_samp_mr_2026-03.xlsx`. Mesmo layout do
`SP/BW/3-Checks/py/check_mr_bloco2.py` (README.md seção 6 item 7): uma aba por
medida e por família, linhas opção x subgrupo, colunas PxQ SAMP [2] | MR |
Delta | Delta % | Participação no delta; aba `Base_PxQ_SAMP` com o detalhe por
grupo (chave, regime, blocos 1 e 2, famílias).

### 4.1 Regras do cálculo

Motor igual ao do `Abertura_PxQ.ipynb` (Passos 4 a 9): tabela
`SP/BW/tbl/tarifas/tarifa_long.csv`, camadas TA = BE + CVA + FIN, vigência
ponderada por dias (mar/26 = 100% vigência 2025), bloco 1 = quantidade x
tarifa, bloco 2 = 1ª passada sobre Fio A + Fio B limitada a [0, 1] + 2ª
passada uniforme, fechando na receita do SAMP.

Diferenças em relação ao PxQ pelo BW, todas possíveis só porque o SAMP traz o
detalhe:
- Chave montada direto dos atributos do SAMP, com o posto por extenso (sem o
  código abreviado do BW).
- TUSD x energia TUSD (ou demanda); TE x energia TE. Cobre cativo (TUSD = TE),
  livre (TE = 0) e suprimento (TUSD inclui PROINFA) sem regra especial.
- **APE usa a tarifa APE do PCAT** (TUSD de energia sem encargos).
- Benefícios sociais pela faixa declarada, sem estimativa pela receita:
  Tarifa Social faixa 01 = 0 (redução de 100% até 80 kWh; receita SAMP também
  0); Tarifa Social faixa 02 e baixa renda legado = chave "Baixa Renda" com a
  isenção de 4.5; Desconto Social faixa 01 e 02 = chaves próprias.
- Regime "fio" só nas linhas com desconto declarado no SAMP (fonte incentivada
  e geradoras do grupo A); todo o resto é "uniforme".
- Energia compensada GD II vira linha própria, detalhe SCEE (como no MR).
- Classe da chave derivada do subgrupo (B1 Residencial, B2 Rural, B4
  Iluminação pública, demais Não se aplica); subclasses não residenciais =
  Não se aplica; detalhe Fonte incentivada, ERC, <500, IRRIG. = Não se aplica.
- Fallbacks de chave (registrados na coluna `REGRA_E`/`REGRA_D`): posto Não
  se aplica -> Fora ponta (18 linhas, energia Azul/APE lançada sem posto, onde
  a TUSD de energia é igual nos postos); posto -> Não se aplica (1 linha, energia
  da Ceris); acessante nomeado -> genérico (Tecipar usa a chave A4 Geração).
- Lado MR idêntico ao `check_mr_bloco2.py` (subsídio distribuído nas famílias
  pelo abatimento de quantidade), exceto que o ERC fica dentro, porque o SAMP
  o traz.

### 4.2 Resultado (bloco 2 x MR)

Quantidades (energia TUSD, TE, demanda): diferença zero.

| Família | Total | CATIVO | LIVRE |
|---|---|---|---|
| Monetário total | +R$ 9,12 mi (+0,50%) | +R$ 9,85 mi | -R$ 0,25 mi |
| FIO A | +0,23% | +0,8% | -1,5% (-R$ 0,64 mi) |
| FIO B | +0,08% | +0,5% | -3,4% (-R$ 1,66 mi) |
| ENCARGOS | +0,74% | +0,6% | **+1,1% (+R$ 1,43 mi)** |
| PERDAS | +1,47% | +0,5% | +7,7% (+R$ 1,23 mi) |
| ENERGIA | +0,66% | +0,7% | - |
| OUTROS | +3,9% | +0,5% | +28% (-R$ 0,62 mi) |

Comparação com o PxQ pelo BW no mesmo mês (`SP/BW/3-Checks/resultados/check_mr_bloco2_2026-03.xlsx`):

| Livres | PxQ BW | PxQ SAMP |
|---|---|---|
| Monetário total | +4,4% (+R$ 10,1 mi) | -0,1% (-R$ 0,25 mi) |
| ENCARGOS | +10,7% (+R$ 13,5 mi) | +1,1% (+R$ 1,43 mi) |
| FIO A | -2,0% | -1,5% |
| FIO B | -4,4% | -3,4% |
| PERDAS | +2,0% | +7,7% |

Memória:
- 1ª passada (desconto de fonte incentivada sobre Fio A + B): R$ -63,47 mi x
  `Subs_CarFI` do MR R$ 63,36 mi — diferença R$ 0,11 mi (0,2%). Pelo BW era
  R$ 0,71 mi.
- 2ª passada total: R$ +3,68 mi.
- Bloco 1 (TA cheia) R$ 1.875,87 mi x MR `Total com Tar de aplicação`
  R$ 1.907,47 mi (-R$ 31,6 mi): esperado — o MR mantém a Tarifa Social faixa
  01 cheia e lança o benefício como subsídio (R$ 32,4 mi); o PxQ zera a faixa.
- Conferências: TA = BE + CVA + FIN nos dois blocos (diferença < 1e-7).

### 4.3 Conclusões

1. **O excesso de ENCARGOS dos livres era a falta de identificação do APE no
   BW.** Com a tarifa APE aplicada pelo detalhe do SAMP, o desvio cai de
   +R$ 13,5 mi para +R$ 1,43 mi. Confirma a hipótese da seção 4.10 do
   `SP/BW/README.md`.
2. **Cativos**: a estrutura por família bate com o MR; todas as famílias ficam
   entre +0,5% e +0,8%, que é o reflexo, via 2ª passada uniforme, dos
   +R$ 9,85 mi de receita faturada acima do `FinTotAj` (teste 1). A diferença
   é de nível de receita, não de abertura.
3. **Livres — resíduos que sobram**, todos rastreáveis:
   - Refaturamento APE: receita SAMP R$ 1,10 mi acima do valor pela tarifa
     APE. A 2ª passada espalha isso pela tarifa APE, que é quase só perdas e
     fio B — explica a maior parte do +R$ 1,23 mi em PERDAS.
   - Verde fonte incentivada: o faturado desconta mais que o desconto sobre o
     fio permite (energia fora ponta Verde não tem fio). A 2ª passada fecha
     com R$ -1,37 mi uniforme, o que empurra FIO B e OUTROS para baixo.
4. **Suprimento (Ceris)**: a tabela de tarifas não tem a TE de suprimento da
   Ceris (Total TE = 0). A receita de energia TE (R$ 482 mil, 1.734 MWh) fica
   como resíduo não ajustável e a família ENERGIA fica zerada. Falta a tarifa
   — não estimada.

### 4.4 BW x SAMP, item a item (mar/26, ambos contra o MR)

Script `SAMP/model/compara_bw_samp_mr.py`, saída
`SAMP/saida/compara_bw_samp_mr_2026-03.xlsx`. PxQ BW = `SP/BW/1-Entrega/pxq/pxq_2026-03.xlsx`
(bloco 2), contra o MR sem ERC (o ERC não passa pelo BW); PxQ SAMP contra o MR com ERC.
Deltas = PxQ - MR.

| Item | PxQ BW - MR | PxQ SAMP - MR | Melhor |
|---|---|---|---|
| Energia TUSD | -12.263 MWh (-0,33%) | 0 | SAMP |
| Energia TE | -12.220 MWh (-0,53%) | 0 | SAMP |
| Demanda | +14.745 kW (+0,18%) | 0 | SAMP |
| Livres - monetário total | +R$ 10,06 mi (+4,4%) | -R$ 0,25 mi (-0,1%) | SAMP |
| Livres - ENCARGOS | +R$ 13,49 mi (+10,7%) | +R$ 1,43 mi (+1,1%) | SAMP |
| Livres - FIO A | -R$ 0,87 mi (-2,0%) | -R$ 0,64 mi (-1,5%) | SAMP |
| Livres - FIO B | -R$ 2,16 mi (-4,4%) | -R$ 1,66 mi (-3,4%) | SAMP |
| Livres - PERDAS | +R$ 0,31 mi (+2,0%) | +R$ 1,23 mi (+7,7%) | BW |
| Livres - OUTROS | -R$ 0,73 mi (+33%) | -R$ 0,62 mi (+28%) | SAMP (marginal) |
| 1ª passada x `Subs_CarFI` | R$ 0,71 mi (1,1%) | R$ 0,11 mi (0,2%) | SAMP |
| Cativos - monetário total | -R$ 1,70 mi (-0,1%) | +R$ 9,85 mi (+0,63%) | BW |
| Cativos - ENCARGOS | -R$ 2,51 mi (-0,7%) | +R$ 2,27 mi (+0,6%) | empate em módulo |
| Cativos - ENERGIA | +R$ 0,75 mi (+0,1%) | +R$ 4,06 mi (+0,7%) | BW |
| Cativos - FIO A / FIO B | +R$ 0,45 mi / -R$ 0,12 mi | +R$ 1,03 mi / +R$ 2,09 mi | BW |
| Cativos - PERDAS | -R$ 0,31 mi (-0,3%) | +R$ 0,46 mi (+0,5%) | BW (marginal) |
| Suprimento - monetário | R$ 0,00 | -R$ 0,48 mi (TE sem tarifa) | BW |
| Geração - monetário | -R$ 3,3 mil | R$ 0,00 | SAMP (marginal) |
| Benefícios sociais | faixas estimadas pela receita | faixas declaradas | SAMP |
| ERC (R$ 0,62 mi) | fora do BW | incluído, fecha | SAMP |
| Disponibilidade | mensal, no fechamento | com defasagem de publicação | BW |

---

## 5. Investigação — receita dos cativos no SAMP (mar/26)

Pergunta: por que o SAMP difere do BW nos cativos se o BW não difere do MR?

**Não é erro de script.** Conferido direto no parquet bruto (B1 Convencional
Residencial Regular, mar/26: Receita Energia R$ 938.760.604, energia 1.283.984
MWh; sem linhas duplicadas no SAMP de SP). As quantidades batem nas três fontes.

**O MR não é faturamento, é quantidade x tarifa.** No MR do período tarifário
2025-26, o B1 Residencial Regular dá exatamente 725,18 R$/MWh em todos os meses
de set/25 a mai/26 — a tarifa TUSD 432,44 + TE 292,74. O BW (faturado) dá
726,06 R$/MWh em mar/26 (+0,12%). Por isso BW e MR ficam próximos: os dois estão
colados em quantidade x tarifa.

**A receita declarada no SAMP tem um adicional que varia no mês.** B1
Residencial Regular, R$/MWh no SAMP: jan 735,67; fev 725,26; mar 731,13; abr
734,19; mai 734,72 (tarifa 725,18 em todos). Em mar/26 o adicional é de
R$ 7,64 mi só nessa linha. A origem não está no SAMP (receita agregada); a
confirmar com a área que prepara o SAMP. Candidatos a checar: custo de
disponibilidade, acertos de faturamento, itens não tarifários somados à
`Receita Energia`.

**Decomposição SAMP - BW, cativos, receita (R$ mi):** B1 Convencional +11,06;
B3 Convencional +1,90; A2 Azul +1,04; A4 Verde -1,89; A4 Azul -0,38; AS Azul
-0,14; demais -0,07. Total +11,52 (SAMP R$ 1.584,19 mi x BW tarifável
R$ 1.572,66 mi).
- A2 Azul cativo: no SAMP só existe como Refaturamento (77.269 kW, R$ 2,31 mi);
  no BW são 40.053 kW, R$ 1,27 mi. Diferença de classificação de
  refaturamento, não de tarifa.
- A4 Verde: o BW tem 43.635 kW e 1.788 MWh a mais que o SAMP.
- BW x MR nos cativos: -R$ 1,70 mi, praticamente a linha SCEE GD II do MR
  (R$ 1,67 mi), que no BW não aparece como receita de fornecimento.

**Qualidade do dado**: no SAMP, jun/26 repete exatamente as receitas, tributos
e UC de mai/26 no B1 Residencial Regular (só a energia muda). Não usar jun/26
do SAMP sem confirmar.

---

## 6. Check reverso — PxQ consolidado (BW cativo + SAMP demais) x PxQ BW (mar/26)

Script `SAMP/model/consolida_bw_samp.py`, saída
`SAMP/saida/consolidado_bw_samp_x_bw_2026-03.xlsx`. Consolidado = linhas do
cativo do PxQ BW (bloco 2 como entregue) + PxQ SAMP de LIVRE, GERAÇÃO,
DISTRIBUIÇÃO e SUPRIMENTO. O SAMP LIVRE já contém o cliente híbrido (energia e
demanda) e as elevatórias da EMAE; nada mais vem do BW.

| Medida | Consolidado | PxQ BW | Delta |
|---|---|---|---|
| Energia TUSD | 3.691.748 MWh | 3.691.705 MWh | +43 MWh (0,00%) |
| Energia TE | 2.305.728 MWh | 2.305.728 MWh | 0 |
| Demanda | 8.425.982 kW | 8.368.067 kW | +57.915 kW (+0,69%) — ERC 58.000 kW, só no SAMP |
| Monetário total | R$ 1.804,52 mi | R$ 1.814,70 mi | -R$ 10,18 mi (-0,56%) |
| FIO A | R$ 164,28 mi | R$ 163,58 mi | +R$ 0,69 mi (+0,42%) |
| FIO B | R$ 501,14 mi | R$ 500,51 mi | +R$ 0,63 mi (+0,13%) |
| ENCARGOS | R$ 495,13 mi | R$ 507,23 mi | -R$ 12,10 mi (-2,38%) |
| PERDAS | R$ 114,46 mi | R$ 113,56 mi | +R$ 0,91 mi (+0,80%) |
| ENERGIA | R$ 548,13 mi | R$ 548,53 mi | -R$ 0,41 mi (-0,07%) |
| OUTROS | R$ -18,62 mi | R$ -18,71 mi | +R$ 0,10 mi |

Por opção: CATIVO e DISTRIBUIÇÃO 0,00; GERAÇÃO +R$ 3,3 mil; SUPRIMENTO
-R$ 0,48 mi (TE da Ceris sem tarifa, seção 4.3); LIVRE -R$ 9,70 mi.

Leitura:
- Quantidades fecham (a única diferença de demanda é o ERC).
- Toda a diferença monetária está nos livres: com a mesma quantidade, a receita
  tarifável dos livres no BW (R$ 240,98 mi) é R$ 10,33 mi maior que a do SAMP
  sem ERC (R$ 230,65 mi). No consolidado, essa diferença sai quase toda de
  ENCARGOS (-R$ 12,06 mi nos livres), coerente com a isenção do APE.
- As linhas "TUSD - Desconto Consumo" do BW (R$ -17,7 mi, fora da base
  tarifável) superam os R$ 10,33 mi em R$ 7,4 mi. Se forem o abatimento que
  leva a receita bruta do BW à receita líquida do SAMP, parte delas é outro
  benefício — mesma pendência da seção 4.10 do `SP/BW/README.md`.

---

## 7. Estrutura e arquivos (desde 29/09/2026)

```
SAMP/
  projeto_samp.md
  data/                         samp-2026.parquet (todas as distribuidoras), dd-samp.pdf
  1-Entrega/
    PxQ_SAMP.ipynb              o cálculo do PxQ SAMP, passo a passo (par do Abertura_PxQ.ipynb do BW)
    pxq/pxq_samp_<ref>.xlsx     saída mensal: abas Componentes e Familias, mesmo desenho do BW
    checks/
      py/                       scripts de check + lib_samp.py (biblioteca comum)
      resultados/               um arquivo por mês e por check
      consolidado_check_pxq_samp_mr_<ini>_a_<fim>.xlsx
  9-Historico/                  model/, saida/ e note/ até 29/09/2026 (motor antigo pxq_samp_engine.py)
```

- **O cálculo existe só no notebook.** Antes, o `pxq_samp_engine.py` recalculava o PxQ dentro de cada
  check. Agora os checks leem o `pxq_samp_<ref>.xlsx` (`lib_samp.carrega_pxq_samp`), como os checks
  do BW leem o `pxq_<ref>.xlsx`. Validação da migração: mar/26 pelo notebook reproduz o motor antigo
  ao centavo (bloco 1 R$ 1.875.870.291,32; 1ª passada R$ -63.473.604,50; 2ª passada
  R$ 3.683.188,54; bloco 2 R$ 1.816.079.875,36), e os checks de jan a mai reproduzem os deltas.
- **Mercado Receita do mês:** todos os checks usam `SP/BW/3-Checks/py/le_mr.py` (arquivo mensal
  se existir; senão o arquivo do período tarifário filtrado por `AnoMes`; mês PROJETADO recusado).
- **Rodar um mês:** notebook com `REFERENCIA` no Passo 0; depois, em `checks/py`,
  `python pxq_samp_check_mr.py <ref>`, `python compara_bw_samp_mr.py <ref>` etc. Consolidado:
  `python SP/BW/3-Checks/py/consolida_check_mr_bloco2.py 2026-01 2026-05 --fonte samp`.

| Script (`1-Entrega/checks/py/`) | Conteúdo |
|---|---|
| `lib_samp.py` | leitura do SAMP bruto e do PxQ SAMP, famílias do MR, leitura do PxQ BW, planilha de comparação |
| `check_samp_mr.py` | Teste 1: SAMP bruto x MR |
| `pxq_samp_check_mr.py` | Teste 2: PxQ SAMP x MR por família |
| `compara_bw_samp_mr.py` | BW e SAMP contra o MR, por opção e família (seção 4.4) |
| `consolida_bw_samp.py` | Check reverso: consolidado x PxQ BW (seção 6) |
| `extrai_linhas_diferenca.py` | linhas SAMP x MR com diferença (seção 5) |

## 7b. Série jan a jun/2026 (29/09/2026)

Notebook rodado de jan a jun; check contra o MR de jan a mai (jun do MR é PROJETADO).
**ERC fora desde 29/09/2026** e em definitivo desde 30/09/2026 (só a cobrança direta aos livres do grupo A; no cativo o ERCAP está embutido na tarifa, componente TE_ESSERR): é o ERCAP (rateio de custo da CCEE pela demanda máxima horária, sem tarifa nem quantidade próprias - `SP/BW/README.md` 4.11): retirado no Passo 1 do notebook e do MR nos checks. Números abaixo já sem ERC; só mai/26 mudou (monetário 18,41 -> 18,89; livre -0,33 -> +0,15).
Todas as conferências do notebook passaram em todos os meses.

Delta PxQ SAMP [2] - MR (R$ milhões):

| | jan | fev | mar | abr | mai | acumulado |
|---|---|---|---|---|---|---|
| Monetário total | 26,31 | -0,61 | 9,12 | 23,64 | 18,89 | 77,35 |
| Cativo | 28,05 | 0,87 | 9,85 | 28,09 | 19,32 | 86,17 |
| Livre | -1,17 | -1,04 | -0,25 | -3,58 | 0,15 | -5,89 |
| Livre - ENCARGOS | 2,12 | 2,61 | 1,43 | -0,24 | 2,80 | 8,72 |
| (BW) Livre - ENCARGOS | 14,76 | 11,99 | 13,49 | 14,95 | 16,27 | 71,45 |

- Quantidades (energia TUSD, TE, demanda) = MR em todos os meses.
- Livres: o SAMP resolve o excesso de ENCARGOS em série (APE com tarifa própria).
- Cativos: a diferença é o nível da receita declarada no SAMP (B1 e B3 Convencional), que oscila
  mês a mês - pendência da seção 5, não da abertura.
- Abr/26: resíduo não ajustável de R$ 1,89 mi (8 grupos: livres A4 Verde R$ 0,93 mi e A2 Azul
  R$ 0,07 mi com receita sem quantidade, além da TE da Ceris). A investigar.
- **Jun/26 do SAMP não é utilizável**: o Passo 2b do notebook detecta 10 grupos com receita idêntica a
  mai/26 e energia diferente (R$ 1,47 bi - B1, B3, B4, A2 Azul, A4 Verde); a receita total de jun é
  igual à de mai. O arquivo `pxq_samp_2026-06.xlsx` é gerado com o aviso.

## 8. Pendências e próximos passos

- ~~Notebook PxQ SAMP~~: feito em 29/09/2026 (`1-Entrega/PxQ_SAMP.ipynb`, seção 7).
- Resíduo não ajustável de abr/26 (R$ 1,89 mi, seção 7b).
- Perguntar à área que prepara o SAMP o que compõe a `Receita Energia` do
  cativo além de quantidade x tarifa (seção 5), e confirmar o jun/26 repetido.
- Refaturamento APE (R$ 1,10 mi acima da tarifa): verificar se é
  refaturamento de meses com outra tarifa ou outro item faturado junto.
- Tarifa TE de suprimento da Ceris: obter a fonte (PCAT ou planilha do SD).
- Teste 2 repetido de jan a mai/26 (seção 7b); teste 1 (SAMP bruto x MR) ainda só em mar/26.
- Fase seguinte (não iniciada): modelo híbrido para meses sem SAMP — BW nos
  cativos, MR nos livres.
