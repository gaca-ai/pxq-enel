# Sessions — ata do projeto RNF (Enel SP)

Registro das reuniões e conversas de trabalho da fase II (Renda Não Faturada). Tem três partes:

1. **Pendências abertas.** Pontos que dependem de decisão ou de confirmação de alguma área.
2. **Informações a levantar.** Dados que, se obtidos, eliminariam uma aproximação do processo.
3. **Registro de decisões.** O que foi decidido, quando e por quê.

Quando uma pendência é resolvida, ela sai da parte 1 e entra na parte 3. A regra resultante passa a
constar no [README](README.md) ou no notebook, e não volta a ser descrita aqui. Este arquivo
substitui o antigo `projeto_RNF.md` (análise inicial de 23/09/2026), guardado no arquivo morto.

---

## 1. Pendências abertas

| # | Tema | O que falta | Com quem | Aberta em |
|---|---|---|---|---|
| P1 | Fluxograma do processo completo | Desenhar um fluxograma que ilustre o processo como um todo, do início (extração do BW, PCATs, tabela de tarifas) até a abertura da RNF, com entradas, transformações, saídas e conferências de cada fase. Uma das últimas etapas do projeto. | Interno | 14/10/2026 |
| P2 | Demanda do livre no ZUBIL | O ZUBIL traz duas colunas de demanda não faturada: AW ("Demanda Ativa Não Faturado") e AY ("Demanda Ativa Não Faturado TUSD"). No cativo, a AW é usada no bloco 1 e o resultado fecha com o valor líquido com menos de 1,5% de diferença em todos os subgrupos do grupo A, o que valida a coluna. No livre, nenhuma das duas bate com a demanda faturada no BW (A4 Azul: AW 393 mil e AY 3.254 mil kW, contra 1.540 mil faturados); até a confirmação, a demanda do livre entra pelo regime fio. Pergunta à área: como o ZUBIL calcula AW e AY (demanda da última fatura proporcional aos dias não faturados?) e por que AY é cerca do dobro do faturado no livre. | Contabilidade / Faturamento | 04/10/2026 |
| P4 | Valor dos ajustes manuais por classe | O relatório traz o volume por classe e o valor total de cada ajuste. O valor por classe é estimado por volume × tarifa média do ZUBIL, e o ajuste herda o mix do SAP da mesma classe. Confirmar o critério com a contabilidade, que pode ter o valor por classe nas abas de cada ajuste. | Contabilidade | 14/10/2026 |
| P11 | PIS/COFINS no ajuste de iluminação pública | Depois de retirado o PIS/COFINS, a iluminação pública fica 5,1% abaixo da tarifa B4a, contra no máximo 1,6% nos demais subgrupos. O valor do ajuste de IP (aba `RNF - IP`, faturamento do mês seguinte) parece já vir sem tributo. Confirmar e, se for o caso, não retirar o tributo dessa linha. | Contabilidade | 14/10/2026 |
| P5 | Valores cativos lançados em categorias livres | O ZUBIL tem R$ 18,7 milhões (brutos) nas colunas de TE (S) e "TUSD TE" (U) em linhas de categoria livre (coluna K com "LIVRE"), com só 21 MWh de energia. Seguindo o relatório, entram como cativo. Valorizados pela demanda da coluna AW, dão o dobro do valor registrado, de forma uniforme: compatível com demanda de cliente livre com 50% de desconto de fonte incentivada. Confirmar o que é esse valor e se deve ser tratado como livre. Baixa prioridade: não altera o total nem a abertura por item (a valorização é só de demanda, em Fio A e Fio B, nos dois mercados), apenas a classificação cativo/livre, que não é necessária por ora (usuário, 21/10/2026). | Contabilidade | 14/10/2026 |
| P6 | Tarifa Social faixa 02 | Valorada pela chave `Baixa Renda`, como no Faturamento. O PCAT 2026 traz uma chave própria, com TE de 305,60 R$/MWh contra 316,78 da Baixa Renda. Decidir a regra para as duas fases. | Tarifas | 04/10/2026 |
| P7 | RNF do uso da distribuição e dos geradores | O ZUBIL não traz linhas desses clientes. Confirmar se há RNF para eles ou se são sempre faturados dentro do mês. | Faturamento | 23/09/2026 |
| P8 | Layout Receita Faturada | Conversão da RNF para o layout ainda não construída. Propostos `TipoFaturamento = RNF` e `OpçãoCompra` CATIVO / LIVRE a partir do mercado. | Planejamento | 23/09/2026 |
| P10 | Procedimento | Escrever `PROCEDIMENTO_RNF.md` (manual gerencial) quando o método estabilizar. Hoje o porquê de cada passo está nas células de texto do notebook. | Interno | 04/10/2026 |
| P12 | Pivôs desatualizados no relatório | De janeiro a maio, a aba `RNF_POR_CLASSE` (e, em fevereiro e março, o pivô do `RESUMO ZUBIL_146`) não reflete o detalhe do ZUBIL: em abril e maio traz o volume de março. Os valores em R$ conciliam com o contabilizado; só os volumes por classe dessas abas ficam inconsistentes. A conferência de volume do notebook passou a ser informativa. Avisar a contabilidade. | Contabilidade | 21/10/2026 |
| P13 | Julho: `Comparativo Novo` diferente do contabilizado | O `Comparativo Novo` de julho traz R$ 1.702,3 milhões; o contabilizado (`Conciliação contábil Novo`) é R$ 1.621,4 milhões, e o próprio relatório acusa −R$ 81,0 milhões no "Check". O relatório de agosto usa os R$ 1.702,3 milhões como reversão de julho. Entender qual valor vale. | Contabilidade | 21/10/2026 |
| P14 | Setembro/2026 | O relatório de setembro está em `src/`, mas o notebook precisa do P × Q do Faturamento do mesmo mês (posto horário). A extração do BW disponível vai até agosto. | Faturamento / Interno | 21/10/2026 |
| P15 | Abril: valor contabilizado ilegível | Na aba `Conciliação contábil Novo` de abril, a célula do valor contabilizado está com erro de fórmula. A conciliação do mês foi feita contra o `Comparativo Novo`. | Contabilidade | 21/10/2026 |
| P16 | Gratuidade da Tarifa Social na RNF | A estimativa por faixa coloca 98,2% da energia da Tarifa Social na faixa 02 (agosto/2026): o ZUBIL valoriza quase todo o consumo pela tarifa `Baixa Renda`, inclusive os primeiros 80 kWh, que são gratuitos para o cliente. Hipótese: a gratuidade é reembolsada pela CDE e registrada como receita. Confirmar. Não afeta o total nem a regra adotada. | Contabilidade | 21/10/2026 |

## 2. Informações a levantar

Em ordem de prioridade.

| # | Informação | O que melhora |
|---|---|---|
| I3 | Significado das colunas de demanda do livre no ZUBIL | Resolve P2: a demanda do livre passa ao bloco 1 |
| I4 | Valor por classe de cada ajuste manual (abas próprias do relatório) | Substitui a estimativa por tarifa média (P4) |
| I5 | Posto horário no ZUBIL, se a transação tiver essa abertura | A divisão por posto deixa de vir do Faturamento do mês |

## 3. Registro de decisões

| Data | Decisão | Motivo | Onde está a regra |
|---|---|---|---|
| 23/09/2026 | Baixa Renda, Tarifa Social e Desconto Social entram no escopo da RNF. | As subclasses existem na RNF. | README 3.2 |
| 04/10/2026 | Estrutura de pastas espelhando o BW: `1-Entrega` (notebook e `rnf/`) e `src/`. A tabela de tarifas é a do BW, sem cópia. | Uma única fonte de tarifa para as duas fases. | README 2 |
| 04/10/2026 | Chave de 8 atributos, igual à da tabela de tarifas. O atributo que o `projeto_RNF.md` chamava de "UC" é o acessante. | Resolve a pendência 4 do `projeto_RNF.md`. | Notebook, Passo 5 |
| 04/10/2026 | Protótipo com a base BIP (`SP_Base_RNF.xlsx`): valorização direta, sem fechamento. **Substituído em 14/10.** | — | Arquivo morto (`versao_base_BIP/`) |
| 14/10/2026 | A base BIP sai do processo. Volume e valor passam a vir do relatório contábil. | A base diverge do relatório em julho (cativo +19%) e agosto (+15%); é de planejamento, não de fechamento. Fecha as antigas pendências de base de fechamento e volume livre de junho. | README 0 e 2 |
| 14/10/2026 | A abertura parte do valor contabilizado: bloco 1 (quantidade × tarifa) levado ao valor do relatório pela mesma regra do bloco 2 do BW. Livre do grupo A no regime fio (diferença em Fio A e Fio B, na composição da tarifa de demanda); demais linhas no rateio proporcional. | A contabilidade valoriza a RNF pelo SAP e pela tarifa média do ZUBIL; o processo deve decompor esse valor, não recalculá-lo. Proposta do usuário. Absorve as antigas pendências de percentual da fonte incentivada e de SCEE. | README 4; notebook, Passo 9 |
| 14/10/2026 | Linhas do SAP: ZUBIL_146 somado por classe, subclasse e categoria tarifária, com tabelas explícitas de categoria (22) e classe (12); uma categoria nova interrompe o processo. Demanda do cativo do grupo A no bloco 1. | O ZUBIL traz energia, demanda e valor no mesmo recorte da tarifa. | Notebook, Passo 3 |
| 14/10/2026 | Ajustes manuais: valor por classe = volume × tarifa média do ZUBIL, ajustado ao total do ajuste; mix herdado das linhas do SAP da mesma classe e mercado; iluminação pública em linha própria (B4a), na tarifa média do ajuste de IP. | É o critério da contabilidade. Pendência P4. | Notebook, Passo 4 |
| 14/10/2026 | Posto horário pela proporção do consumo e da demanda faturados no P × Q do Faturamento do mesmo mês: subgrupo, modalidade, mercado e classe; depois sem classe; depois só subgrupo e modalidade. | Proposta do usuário. O faturado do mês cobre um período um pouco diferente do não faturado, mas a proporção entre postos muda pouco de um mês para o outro. | Notebook, Passo 6 |
| 14/10/2026 | Benefícios sociais: faixa de consumo estimada pelo valor de cada linha, como no Faturamento (Tarifa Social: faixa 01 gratuita e faixa 02 `Baixa Renda`; Desconto Social: chaves próprias). | O ZUBIL traz o valor por subclasse, mas não a faixa. Fecha as antigas pendências de gratuidade e de subclasses legadas. | Notebook, Passo 8 |
| 14/10/2026 | Reativo e bandeira (do ZUBIL e dos ajustes) fora da abertura, na aba `Nao_tarifaveis`, e de volta na conciliação. | Não têm tarifa por item; o total precisa fechar com o relatório. | Notebook, Passos 3, 4 e 11 |
| 21/10/2026 | Cativo e livre seguem a convenção do relatório em todo o processo: colunas de TE e "TUSD TE" do ZUBIL são cativo, a coluna "TUSD" é livre, qualquer que seja a categoria. Categoria livre com valor nas colunas do cativo gera linha cativa própria, com a demanda da coluna AW. | Pedido do usuário: cada valor do README pode ser conferido diretamente no relatório (`Comparativo Novo` D9, F9, G9 e I9). Pendência P5. | README 5.1; notebook, Passos 3 e 11 |
| 14/10/2026 | Fluxograma do processo completo, do BW à RNF, registrado como uma das últimas etapas do projeto. | Pedido do usuário. | Pendência P1 |
| 14/10/2026 | Os valores do ZUBIL (e, por consequência, os ajustes manuais) contêm PIS/COFINS. O tributo é retirado antes do bloco 2 pela alíquota do próprio relatório (aba `RNF TOI SP`, "Alíquota de COFINS" + "Alíquota de PIS"): 5,35% em agosto (0,96% + 4,39%). Informado pelo usuário; substitui a alíquota estimada pelo BW (5,44%). O tributo vai para os não tarifáveis e volta na conciliação. | Confirmado pelo usuário. A tarifa homologada não tem tributos; sem a retirada, o rateio espalhava o tributo pelos itens. Com a retirada, o cativo fecha com menos de 1,5% de diferença em todos os subgrupos. | README 3.2 e 5.1; notebook, Passo 4b |
| 14/10/2026 | Demanda não faturada do cativo (coluna AW do ZUBIL) mantida no bloco 1. | A demanda não é energia acumulada, mas a receita de demanda dos dias não faturados é receita do mês. Com a AW no bloco 1, o grupo A cativo fecha com o valor líquido tão bem quanto o grupo B. | Notebook, Passo 8 |
| 21/10/2026 | Série de janeiro a agosto/2026 processada com os relatórios mensais. O nome do relatório segue o padrão `RNF CCS <mmaaaa> Fator K Nova Metodologia.xlsb`; todo mês muda só `REFERENCIA`. | Relatórios fornecidos pelo usuário. | README 2 e 5 |
| 21/10/2026 | A conciliação final é feita contra o valor contabilizado (conta 1120110910, `Conciliação contábil Novo`, coluna I), não contra o `Comparativo Novo`, que fica como informação. Quando o contabilizado está ilegível, vale o `Comparativo Novo`. | O total aberto bate com o contabilizado em todos os meses; o `Comparativo Novo` diverge em julho e o relatório acusa a diferença. Pendências P13 e P15. | Notebook, Passo 11 |
| 21/10/2026 | Conferência de volume contra `RNF_POR_CLASSE` passa a ser informativa. | O pivô dessa aba não foi atualizado em vários meses (P12); o controle de valor é a conciliação. | Notebook, Passo 10 |
| 21/10/2026 | Tarifa proporcional no mês do reajuste: peso de cada vigência por subgrupo estimado pelo valor líquido das linhas cativas do SAP sem benefício, `w = (R − A) / (B − A)`; calendário quando as tarifas diferem menos de 1% ou não há linha de calibração. Substitui a regra "tarifa do último dia do mês". | Em julho, com a tarifa nova, as categorias cativas ficavam 5% a 11% abaixo do bloco 1; o ZUBIL valoriza a maior parte de julho na tarifa antiga (peso da nova: 0% a 12% no grupo A, cerca de 25% no B), o mesmo padrão do Faturamento (P18 do BW). Com a regra, julho fecha como os outros meses. | README 3.2 e 5.1; notebook, Passos 7 e 8 |
| 21/10/2026 | Tabela 5.1 do README no mesmo formato da tabela 5.1 do README do BW (linhas (a) a (h), mês a mês). APE (d) não se aplica; o regime fio ocupa a linha (e). | Pedido do usuário: comparar Faturamento e RNF lado a lado. | README 5.1 |
| 21/10/2026 | Aproximação das faixas dos benefícios sociais (participação da faixa 02 estimada pelo valor registrado) validada pelo usuário e descrita na seção 4 do README. | Mesma regra do Faturamento; o ZUBIL não traz a faixa de consumo. | README 4; notebook, Passo 8 |
