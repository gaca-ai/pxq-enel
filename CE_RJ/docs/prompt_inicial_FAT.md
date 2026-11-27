Vamos iniciar o P × Q do faturamento da Enel CE e da Enel RJ, replicando o método já construído e validado para a Enel SP.

## Contexto
- Repositório: `C:\python_projects\claude\p_x_q` (o mesmo de SP). O projeto novo fica em `CE_RJ/FAT`. A RNF de CE e RJ é outro projeto, em `CE_RJ/RNF`, conduzido em outro chat; não altere nada lá sem me perguntar.
- CE e RJ usam o mesmo sistema de faturamento: um só processo para as duas, com as diferenças marcadas no texto como [CE] ou [RJ].
- Não há BW. A base virá do sistema comercial, conforme a especificação que enviamos à TI (`CE_RJ/docs/especificacao/especificacao_base_faturamento_RJ_CE.pdf`, v1.0). Ainda não sei se ela será atendida. A alta tensão deve vir cliente a cliente (UC e fatura); a baixa tensão, agregada.
- A documentação de SP (README, procedimento e sessions do BW e da RNF) está em `CE_RJ/docs/referencias`. É a referência do método, das regras e das decisões. O código de SP está em `SP/BW/3-Py` e `SP/BW/1-Entrega/Abertura_PxQ.ipynb`.
- Já criei os templates da documentação nova, para você preencher: `CE_RJ/FAT/README.md`, `CE_RJ/FAT/1-Entrega/procedimentos/PROCEDIMENTO_PXQ.md` e `CE_RJ/FAT/sessions.md`. Eles mostram, passo a passo, o que foi feito em SP e o que se espera mudar em CE e RJ.
- Todo o código será enviado à TI para ser acoplado ao app que já existe (o app de geração dos layouts, que hoje recebe as releases de SP). Vou pedir à TI o código do app para desenharmos uma estrutura fácil de acoplar. Não começamos a codificar antes disso.
- Não duplicar código: o código de SP será generalizado (parametrizado pela distribuidora) e compartilhado por SP, CE e RJ. O inventário do que pode ser reaproveitado, com o que depende de SP em cada script e as dependências entre eles, está em `CE_RJ/FAT/3-Py/README_py.md`. Entre eles: `extrai_pcat.py` (tabela de tarifas), o conversor do layout Receita Faturada, o Bl1 × Bl2 e os checks contra o Mercado Receita.
- Insumos já disponíveis: PCATs de CE (2025, 2026 reajuste e 2026 recálculo) e de RJ (2025 e 2026) em `CE_RJ/docs/PCAT`, e o Mercado Receita dos reajustes em `CE_RJ/docs/MR_Reajuste`. O PCAT 2024 não foi recebido.

## Como trabalhar
- Valores monetários sempre em MR$; energia em MWh; demanda em MW.
- Código Python comentado em inglês. Sem gambiarras: prefiro a solução definitiva.
- Faça perguntas de alinhamento antes de tarefas complexas, e só as necessárias.
- O `sessions.md` é um diário scrum (milestones, pendências e decisões importantes). Peça autorização antes de alterá-lo.
- Eu faço os commits pelo GitHub Desktop; você me passa a mensagem de commit.
- Ambiente: venv `C:\python_projects\.venv`, executado pelo Desktop Commander (PowerShell; use `PYTHONIOENCODING=utf-8`).
- Toda mudança no código compartilhado não pode alterar os resultados de SP: confira SP antes e depois (os valores de referência estão nas releases, `SP/releases/`).
- O que já foi decidido em SP vale para CE e RJ, salvo decisão em contrário registrada no sessions.

## Primeira etapa
1. Leia a documentação de SP em `docs/referencias`, a especificação, os templates e o inventário `CE_RJ/FAT/3-Py/README_py.md`.
2. Quando eu trouxer o código do app da TI, analise a estrutura dele e proponha como organizar o código comum (SP, CE e RJ): pacote compartilhado, configuração por distribuidora e o que fica específico de cada uma. O objetivo é que a TI acople com o mínimo de adaptação.
3. Só depois disso, comece pela tabela de tarifas de CE e RJ (generalizando o `extrai_pcat.py`), que serve também à RNF de CE e RJ. Trate as duas vigências de 2026 do CE (reajuste e recálculo).
4. Em cada etapa, apresente o plano com as perguntas de alinhamento antes de codificar.
