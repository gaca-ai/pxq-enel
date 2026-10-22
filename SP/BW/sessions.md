# Sessions — ata do projeto P × Q (Enel SP)

Registro das reuniões e conversas de trabalho. Tem três partes:

1. **Pendências abertas.** Pontos que dependem de decisão ou de confirmação de alguma área.
2. **Informações a levantar.** Dados que, se obtidos, eliminariam uma aproximação do processo.
3. **Registro de decisões.** O que foi decidido, quando e por quê.

Quando uma pendência é resolvida, ela sai da parte 1 e entra na parte 3. A regra resultante passa a
constar no [README](README.md) ou no
[procedimento](1-Entrega/procedimentos/PROCEDIMENTO_PXQ.md), e não volta a ser descrita aqui.

---

## 1. Pendências abertas

| # | Tema | O que falta | Com quem | Aberta em |
|---|---|---|---|---|
| P1 | Tarifa da Ceris (suprimento) | A tabela homologada traz a chave da Ceris com o lado TE zerado; a planilha do SD traz tarifa de TUSD e de TE. Hoje o lado TE é aberto pela aba `SD tarifas`, só na tarifa de aplicação, e o conversor o aloca na base econômica. Definir qual tabela prevalece. | Regulatório | 19/09/2026 |
| P2 | Camada financeira (FIN) | O FIN é calculado como TA − BE − CVA, e o arredondamento se acumula nessa camada. Verificar se o PCAT original traz a camada financeira. | Regulatório | 19/09/2026 |
| P3 | Base de receita do SD sem PIS/COFINS | Adotada a receita líquida de tributos. Falta a área confirmar que é a base correta para a contabilidade. | Faturamento / Contabilidade | 18/09/2026 |
| P4 | Campo `Encargo Total (EUSD)` do SD | Tem significado diferente entre as abas: inclui PIS/COFINS na Ceris e é líquido no PCH Rasgão. Até a resposta, cada aba é tratada pelo que contém. | Faturamento | 19/09/2026 |
| P5 | PCH Pirapora | Some em junho e volta em julho com o valor dobrado. Confirmar se houve atraso de faturamento e se o caso é recorrente. | Faturamento | 19/09/2026 |
| P6 | Resíduo não ajustável | Grupos com receita e cálculo zero ficam fora do bloco 2 e são informados à parte. Decidir se mantém assim ou se cria um item residual explícito na abertura. | Planejamento | 19/09/2026 |
| P7 | Tributos no layout Receita Faturada | Adotados ICMS na coluna AC e PIS/COFINS na AB, com as subvenções somadas a cada um. A parcela TUSD e os tributos sem parcela são repartidos entre demanda e energia pela receita do bloco 2. Falta a confirmação formal. | Planejamento | 21/09/2026 |
| P8 | SCEE GD1 e GD2 | O BW não distingue as duas modalidades. Em 23/09 ficou definido que os valores virão do Mercado Receita do Sparta; a integração ainda não foi feita. Até lá, o movimento fica em CATIVO no layout. | Planejamento / Tarifas | 21/09/2026 |
| P9 | Energia compensada fora ponta e intermediário | Os códigos do posto convencional são tarifáveis; os de fora ponta e intermediário não. Parece inconsistência de cadastro. | Faturamento | 21/09/2026 |
| P10 | Reativos, ultrapassagem e Res. 166/05 fora do layout | Mantidos fora, por não terem tarifa por item. Confirmar se a área espera vê-los em algum grupo do layout. | Planejamento | 21/09/2026 |
| P11 | Encargo da conta COVID (ZQC237) | Incluído em CDE Contas. Confirmar a origem provável em refaturamento. | Faturamento | 21/09/2026 |
| P12 | Check por grupo de tensão (AT/BT) contra o informe de gestão | O informe só abre por classe e segmento. Depende da aba 2.4 (Categoria de Tarifas). | Controlling | 21/09/2026 |
| P13 | Tecipar e UTE São João | A demanda das duas usinas não consta do informe de gestão. Confirmar se devem compor a base do informe. | Controlling | 21/09/2026 |
| P14 | Interface Gerencial (INTERFACE_ACT) | Segunda saída, com de-para próprio fornecido pela área (o arquivo não está em `aux_files`), ainda não construída. Depende de P8. A abertura de bandeiras e tributos por classe já está disponível na aba `Nao_tarifaveis_linha`. | Planejamento | 22/09/2026 |
| P15 | Grupos cativos com bloco 1 abaixo da receita | Em junho/2026, antes dos tratamentos dos benefícios sociais e da tarifa proporcional, 58 grupos cativos (a maioria de demanda) tinham o bloco 1 de 5% a 48% abaixo da receita. Reavaliar com a versão final antes de investigar a causa. | Interno | 23/09/2026 |
| P16 | Faixa 1 do Desconto Social no Mercado Receita | O MR lança na faixa 1 cerca de um quarto da energia que a receita do BW indica, e tem menos clientes na faixa 1 que na 2. Isso sugere que o MR classifica o cliente inteiro pela faixa do consumo total, em vez de partir o consumo em 120 kWh. Levar à área tarifária. | Tarifas | 24/09/2026 |
| P17 | Demanda A2 no Mercado Receita | Em jan/2026 a demanda dos livres A2 no BW fica 19% abaixo do MR (origem do desvio de Fio A desse mês). A demanda dos cativos A2 fica abaixo do MR em todos os meses. Verificar faturamento fora do mês ou refaturamento. | Interno / Faturamento | 29/09/2026 |
| P18 | Tarifa proporcional | (a) Confirmar com o faturamento o calendário de leitura da alta tensão: a estimativa indica a conta de julho toda na tarifa antiga. (b) Decidir se subgrupos pequenos (A2, B2) herdam o peso do grupo de tensão em vez do calendário. (c) Decidir se a calibração vai para o P × Q pelo SAMP. | Faturamento / Interno | 01/10/2026 |
| P19 | Subclasse da iluminação pública | Os códigos ZQC167 e ZQC168 vão sempre para B4a, mesmo em categoria B4b. Confirmar a regra. | Faturamento | 16/09/2026 |
| P20 | E-mail final às áreas | Texto ainda não fechado. Registrar junto a decisão sobre ultrapassagem. | Interno | 21/09/2026 |

## 2. Informações a levantar

Nenhuma delas bloqueia o processo, mas cada uma substitui uma estimativa por um dado. Estão em
ordem de prioridade.

| # | Informação | O que melhora |
|---|---|---|
| I1 | Data de leitura (ou ciclo de faturamento) na extração do BW | A tarifa proporcional deixa de ser estimada pela receita (procedimento, Passo 7b) |
| I2 | Identificação das instalações APE no BW | A isenção do APE passa a ser aplicada instalação a instalação, sem o descasamento temporal do complemento |
| I3 | Percentual de desconto de fonte incentivada por instalação | A 1ª passada deixa de ser calculada pela diferença |
| I4 | Encontro de contas da SCEE separado do faturamento bruto | Permite separar GD1 e GD2 (P8) |
| I5 | Faixa de consumo dos clientes de Tarifa Social e Desconto Social | As faixas deixam de ser estimadas pela receita (procedimento, Passo 8) |
| I6 | Importe na aba "Usinas Geradoras" do SD | A receita dos quatro geradores sem aba própria deixa de vir do consolidado. Já solicitado à área. |

## 3. Registro de decisões

| Data | Decisão | Motivo | Onde está a regra |
|---|---|---|---|
| 18/09/2026 | Receita do SD usada líquida de PIS/COFINS. A conferência contra a aba consolidada desconta os tributos dos dois lados. | As abas trazem encargo e tributos separados; sem critério único o total varia cerca de 9%. Confirmação pendente em P3. | Procedimento, Passo 12 |
| 21/09/2026 | Layout Receita Faturada: todo o movimento como FAT; híbrido como livre; Residencial Baixa Renda reúne Baixa Renda, BPC, Indígena, Quilombola e Desconto Social; entrega só das linhas com valor; substituição tributária fora (débito e crédito se anulam); conta COVID em CDE Contas. | A extração não distingue FAT, REFAT e CNR; o layout não tem opção para híbrido; Baixa Renda é classe no layout e subclasse no BW. | Procedimento, "Conversão para o layout Receita Faturada" |
| 21/09/2026 | No check contra o informe de gestão, o consumo da Ceris é contado uma vez por chave, com Proinfa. | A contagem anterior somava as pernas TE e TUSD. | — |
| 21/09/2026 | Resíduo de receita no check de junho contra o informe de gestão aceito como explicado (PIS/COFINS do SD e grupos sem parcela ajustável). | — | — |
| 23/09/2026 | Tarifas lidas diretamente dos PCATs numa tabela única, com todas as vigências e camadas. Substitui a planilha de de-para manual. | A planilha manual não tinha rastro do PCAT de origem e gerou erros de vigência. | Procedimento, Passo 4 |
| 23/09/2026 | Isenção de encargos para Baixa Renda e Desconto Social aplicada no bloco 1. | PRORET 7.1, §18 e §23. | README 3.5; procedimento, Passo 7 |
| 23/09/2026 | A saída separa, por grupo, o ajuste da 1ª passada (desconto regulatório) e o da 2ª (fechamento). | Medir quanto do ajuste é desconto e quanto é fechamento de conta. | Procedimento, Passo 9 |
| 23/09/2026 | SCEE GD1 e GD2 não são separáveis pelo BW; os valores virão do Mercado Receita do Sparta. | Os códigos de energia compensada distinguem posto, não tipo de geração. Integração pendente em P8. | — |
| 24/09/2026 | Tarifa Social e Desconto Social tratados como benefícios distintos. A tarifa do Desconto Social é derivada da residencial nas vigências cujo PCAT não a traz. | REN 1.000/2021, art. 179, § 1º e § 4º (redação da REN 1.147/2025). A premissa anterior de que eram o mesmo benefício estava errada. | README 3.5; procedimento, Passo 4 |
| 24/09/2026 | Faixas de consumo da Tarifa Social e do Desconto Social estimadas pela receita no bloco 1. Substitui a decisão de 23/09 de deixar o efeito para a 2ª passada. | O BW não traz a faixa; deixar para a 2ª passada distorcia a abertura por item. | Procedimento, Passo 8 |
| 25/09/2026 | Tributos identificados pela operação principal `R/#` e bandeiras pelos oito códigos ZQD da área. ICMS, PIS e COFINS separados pelo campo de montante. | Regra informada pela área de faturamento. | Procedimento, Passo 2 |
| 25/09/2026 | Regras da conferência contra o Mercado Receita: subsídio do MR devolvido aos itens pelo abatimento de quantidade; energia TUSD do cativo igual à TE; comparação pelo bloco 2. | Os dois lados passam a carregar os descontos dentro dos itens. | Procedimento, "Conferência anual contra o Mercado Receita" |
| 29/09/2026 | Parâmetro para desligar a 2ª passada e listar as diferenças linha a linha do BW. | Investigar a origem das diferenças sem o rateio de fechamento. | Procedimento, Passos 0 e 9 |
| 30/09/2026 | ERCAP fora do P × Q para os livres do grupo A; nos cativos continua no item `TE_ESSERR`. O valor dos livres vem do Mercado Receita do Sparta. | É rateio de custo da CCEE pela demanda máxima, sem preço nem quantidade próprios. Nos cativos já está na tarifa. | README 3.7 |
| 01/10/2026 | Linhas "TUSD - Desconto Consumo" e "Complemento Res. 166/05" tratadas como isenção de encargos do APE (regime APE). O descasamento mês a mês do complemento é aceito, e SP fica finalizado assim. | Confirmado pela área com o detalhe por instalação e faturas: 301 de 305 instalações são APE. | README 3.4; procedimento, Passo 9 |
| 01/10/2026 | Papel de cada fonte de validação: Mercado Receita do reajuste uma vez por ano; SAMP como referência futura; informe de gestão mensal; simulador Sparta não é referência. | O MR só existe uma vez por ano e o simulador é o que o P × Q substitui. | README 5.2 |
| 01/10/2026 | Tarifa proporcional estimada pela receita nos meses de reajuste, com régua padrão dos 3 últimos meses normais. | Escolhida entre as opções apresentadas; o calendário do mês distorcia a abertura. Pontos em aberto em P18. | Procedimento, Passo 7b |
| 04/10/2026 | Plano de migração para Databricks retirado do escopo. | Foco no notebook e no procedimento. | — |
| 04/10/2026 | Conversor para o layout Receita Faturada criado a partir do de-para do próprio modelo. O notebook passa a gravar a base econômica por item e os não tarifáveis linha a linha. | Não havia script de conversão; a abertura de tributos por classe exigia o dado linha a linha. Premissas pendentes em P1, P7 e P8. | Procedimento, "Conversão para o layout Receita Faturada" |
| 04/10/2026 | Pacote final da parte BW: pastas `1-Entrega`, `2-Apoio`, `3-Checks`, `tbl` e `src`; versões anteriores e material de investigação no arquivo morto. O Mercado Receita mensal de março foi excluído, por ser idêntico ao recorte do arquivo do período tarifário. | Fechar SP (BW) para avançar na Renda Não Faturada. | README 2 |
| 04/10/2026 | Documentação reorganizada: o README é o índice explicado, o procedimento é o manual do notebook e este arquivo é a ata. | Eliminar redundâncias e contradições entre os três documentos. | — |
| 21/10/2026 | Ultrapassagem de demanda do cativo (`ZQD021`, `ZQD022`, `ZQD023`) passa a não tarifável, como a do livre (`ZQD019`, `ZQD020`, `ZQD037`). Os códigos de diferença de demanda contratada (`ZQD058` a `ZQD063`) continuam tarifáveis. Total de códigos tarifáveis: 30. Janeiro a agosto reprocessados. | A ultrapassagem é multa, sem item na tabela homologada; no BW cada código tem linha própria com quantidade e R$. Antes, a multa entrava em Fio A e Fio B e os mesmos kW eram valorizados duas vezes no bloco 1. | Procedimento, Passo 2; README 5.1 |
