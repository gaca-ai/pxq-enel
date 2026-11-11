# -*- coding: utf-8 -*-
"""
Receitas consolidadas: o template Receitas Consolidadas (abas Receita, RI e Parcela B), com os meses
lado a lado, e abas adicionais para as demais familias e para o subsidio.

Uso:
    python receitas_consolidadas.py 2026
Le
    1-Entrega/receita_faturada/receita_faturada_<ano>-<mm>.xlsx            Faturamento (FAT + REFAT + CNR)
    2-Checks/receita_faturada_Bl1_vs_Bl2/receita_faturada_Bl1_vs_Bl2_<ano>-<mm>.xlsx
                                                                          Subsidio (abas desconto_*)
    ../RNF/1-Entrega/receita_nao_faturada/receita_nao_faturada_<ano>-<mm>.xlsx
                                                                          RNF (aba Saldo)
Grava
    1-Entrega/receitas_consolidadas/receitas_consolidadas_<ano>.xlsx

Abas:
- Leia-me: fontes, regras e correcoes de formula em relacao ao template.
- Receita: Por Componente (Faturamento, Subsidio e RNF por componente), Demais ingressos (Covid Migrantes, Bandeira,
  UD/ER, Outros, PIS/COFINS e ICMS, faturado e nao faturado) e Por Cliente (Cativo e Livre por classe).
- RI, Parcela B, Compra de Energia, Perdas, Transmissao, Encargos, CVA e Ativos e Passivos
  Regulatorios: Por Cliente so com a familia da aba.
- Receita P&C: uma linha por coluna do layout Receita Faturada (Parcela B a ICMS), com os subtotais
  Tarifavel e Demais ingressos; blocos Receita = Faturamento + Subsidio + RNF.
- Uma aba por coluna do layout sem aba de familia (Rede Basica a TFSEE e Demais ingressos), Por Cliente.
- Subsidio: Tarifa Social - DMR e Fonte incentivada, por componente e por cliente.
- Conferencia: cada bloco contra a sua fonte, Por Componente contra Por Cliente e P&C contra
  Por Componente + Demais ingressos.

Regras:
- Valores em MR$ (milhoes de reais). Fontes em R$ mil.
- Componentes (colunas do layout Receita Faturada): Parcela B; RI; Compra de Energia; Perdas;
  Transmissao = Rede Basica + Transporte Itaipu + CUSD + Conexao T; Encargos = CDE USO + CDE CONTAS +
  CDE ELET + PROINFA + ESS/ERR/ERCAP + ONS + P&D + TFSEE; CVA = Base CVA; Ativos e Passivos
  Regulatorios = Base Financeira. Receita Liquida = soma dos componentes (sem Demais ingressos).
- Demais ingressos: Covid Migrantes (ZQC237, encargo da Conta-COVID dos migrantes para o livre, nao
  tarifavel), Bandeira, UD/ER, Outros, PIS/COFINS e ICMS.
- Subsidio = Tarifa Social - DMR + Fonte incentivada (tarifa homologada - tarifa realizada do
  Bl1 x Bl2; positivo = desconto). RNF = aba Saldo (provisao do mes + estorno da provisao anterior).
- Classes: mercado CATIVO e SCEE nas classes do Cativo; SUPRIMENTO, USO D e GERACAO no Cativo, nas
  linhas Suprimento, Uso Distribuicao e Uso Geracao; LIVRE nas classes do Livre.
- Meses sem fonte ficam em branco. Linha 3: YTD ate o ultimo mes com receita_faturada, YTG depois;
  coluna YTD = SUMIFS dos meses YTD. Totais e YTD sao formulas.
"""
import os
import sys
import datetime as dt
from pathlib import Path

import pandas as pd
import openpyxl
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.utils import get_column_letter as L

BASE = Path(os.environ.get("PXQ_BASE", r"C:\python_projects\claude\p_x_q\SP\BW"))
DIR_RF = BASE / "1-Entrega" / "receita_faturada"
DIR_BL = BASE / "2-Checks" / "receita_faturada_Bl1_vs_Bl2"
DIR_RNF = BASE.parent / "RNF" / "1-Entrega" / "receita_nao_faturada"
ANO = int(sys.argv[1]) if len(sys.argv) > 1 else max(int(p.stem[-7:-3]) for p in DIR_RF.glob("receita_faturada_????-??.xlsx"))
DIR_SAIDA = BASE / "1-Entrega" / "receitas_consolidadas"
SAIDA = DIR_SAIDA / f"receitas_consolidadas_{ANO}.xlsx"
MESES = [f"{ANO}-{m:02d}" for m in range(1, 13)]
UN = "MR$"

# template component -> layout columns
COMPONENTES = [
    ("Parcela B", ["Parcela B"]),
    ("RI", ["RI"]),
    ("Compra de Energia", ["Compra de Energia"]),
    ("Perdas", ["Perdas"]),
    ("Transmissão", ["Rede Básica", "Transporte Itaipu", "CUSD", "Conexão T"]),
    ("Encargos", ["CDE USO", "CDE CONTAS", "CDE ELET", "PROINFA", "ESS/ERR/ERCAP", "ONS", "P&D", "TFSEE"]),
    ("CVA", ["Base CVA"]),
    ("Ativos e Passivos Regulatórios", ["Base Financeira"]),
]
DEMAIS = [("Covid Migrantes", ["Covid Migrantes"]), ("Bandeira", ["Bandeira"]), ("UD/ER", ["UD/ER"]), ("Outros", ["Outros"]),
          ("PIS/COFINS", ["PIS/COFINS"]), ("ICMS", ["ICMS"])]
COLS_COMP = [c for _, cs in COMPONENTES for c in cs]
COLS_DEMAIS = [c for _, cs in DEMAIS for c in cs]
COLS = COLS_COMP + COLS_DEMAIS

CLASSES = ["Residencial", "Residencial Baixa Renda", "Industrial", "Comercial", "Rural", "Poder Público",
           "Iluminação Pública", "Serviço Público", "Consumo Próprio", "Suprimento", "Uso Distribuição",
           "Uso Geração"]
GRUPOS = ["Cativo", "Livre"]
# layout class -> template class
DE_PARA_CLASSE = {"Residencial": "Residencial", "Residencial Baixa Renda": "Residencial Baixa Renda",
                  "Industrial": "Industrial", "Comercial": "Comercial", "Rural": "Rural",
                  "Poderes Públicos": "Poder Público", "Iluminação Pública": "Iluminação Pública",
                  "Serviços Públicos": "Serviço Público", "Consumo Próprio": "Consumo Próprio"}
MERCADO_CATIVO = {"SUPRIMENTO": "Suprimento", "USO D": "Uso Distribuição", "GERAÇÃO": "Uso Geração"}
OPCOES_CATIVO = {"CATIVO", "SCEE - GD1", "SCEE - GD2"}

# Por Cliente blocks: (title, source key)
FONTES = [("Faturamento", "FAT"), ("Subsídio", "SUB"), ("RNF", "RNF")]


def chave(op, cl):
    """(group, template class) of a layout option x class."""
    if op in OPCOES_CATIVO:
        return "Cativo", DE_PARA_CLASSE.get(cl)
    if op in MERCADO_CATIVO:
        return "Cativo", MERCADO_CATIVO[op]
    if op == "LIVRE":
        return "Livre", DE_PARA_CLASSE.get(cl)
    return None, None


def agrega(d, origem):
    """Layout frame (R$ mil) -> MR$ by (group, class) for the value columns."""
    d = d[d["OpcaoCompra"].notna()].copy()
    falta = [c for c in COLS if c not in d.columns]
    if falta:
        raise KeyError(f"{origem}: colunas ausentes (regerar o arquivo): {falta}")
    k = [chave(o, c) for o, c in zip(d["OpcaoCompra"], d["Classe"])]
    d["GRUPO"], d["LINHA"] = [g for g, _ in k], [c for _, c in k]
    sem = d[d["LINHA"].isna()]
    if len(sem):
        raise ValueError(f"{origem}: opcao x classe sem linha: {sorted(set(zip(sem.OpcaoCompra, sem.Classe)))}")
    for c in COLS:
        d[c] = pd.to_numeric(d[c], errors="coerce").fillna(0.0)
    return d.groupby(["GRUPO", "LINHA"])[COLS].sum() / 1000.0


def le_fat(ref):
    arq = DIR_RF / f"receita_faturada_{ref}.xlsx"
    if not arq.exists():
        return None
    return agrega(pd.read_excel(arq, sheet_name="Receita Faturada", header=2), arq.name)


def le_sub(ref, aba):
    arq = DIR_BL / f"receita_faturada_Bl1_vs_Bl2_{ref}.xlsx"
    if not arq.exists():
        return None
    return agrega(pd.read_excel(arq, sheet_name=aba, header=2), f"{arq.name}/{aba}")


def le_rnf(ref):
    arq = DIR_RNF / f"receita_nao_faturada_{ref}.xlsx"
    if not arq.exists():
        return None
    return agrega(pd.read_excel(arq, sheet_name="Saldo", header=2), arq.name)


DADOS = {"FAT": {}, "TS": {}, "FI": {}, "RNF": {}}
for ref in MESES:
    DADOS["FAT"][ref] = le_fat(ref)
    DADOS["TS"][ref] = le_sub(ref, "desconto_tarifa_social_dmr")
    DADOS["FI"][ref] = le_sub(ref, "desconto_fonte_incentivada")
    DADOS["RNF"][ref] = le_rnf(ref)
DADOS["SUB"] = {r: (None if DADOS["TS"][r] is None else DADOS["TS"][r].add(DADOS["FI"][r], fill_value=0.0))
                for r in MESES}
com_fat = [r for r in MESES if DADOS["FAT"][r] is not None]
if not com_fat:
    sys.exit(f"nenhum receita_faturada de {ANO} em {DIR_RF}")
ULTIMO = com_fat[-1]

# ------------------------------------------------ workbook layout (template columns)
CD, CU, CG, C0 = 4, 5, 7, 11                      # Descricao (D), Unid (E), YTD (G), first month (K)
CMES = [C0 + k for k in range(12)]                # K..V
AZUL = PatternFill("solid", fgColor="1F4E79")
CINZA = PatternFill("solid", fgColor="D9E1F2")
FMT = '#,##0.00;-#,##0.00;"-"'


def prepara(ws, titulo):
    ws.cell(1, CD, titulo).font = Font(bold=True, size=12)
    ws.cell(2, CD, "MR$ = milhões de reais").font = Font(italic=True, size=9, color="808080")
    ws.cell(3, CG, ANO)
    for k, c in enumerate(CMES):
        ws.cell(3, c, "YTD" if MESES[k] <= ULTIMO else "YTG").alignment = Alignment(horizontal="center")
    for c, t in ((CD, "Descrição"), (CU, "Unid"), (CG, "YTD")):
        ws.cell(4, c, t)
    for k, c in enumerate(CMES):
        ws.cell(4, c, dt.datetime(ANO, k + 1, 1)).number_format = "mmm/yyyy"
    for c in range(CD, CMES[-1] + 1):
        ws.cell(4, c).font = Font(bold=True, color="FFFFFF")
        ws.cell(4, c).fill = AZUL
        ws.cell(4, c).alignment = Alignment(horizontal="center")
    ws.column_dimensions["A"].width = 2
    for x in "BCFHIJ":
        ws.column_dimensions[x].width = 2
    ws.column_dimensions[L(CD)].width = 42
    ws.column_dimensions[L(CU)].width = 7
    ws.column_dimensions[L(CG)].width = 13
    for c in CMES:
        ws.column_dimensions[L(c)].width = 12
    ws.freeze_panes = ws.cell(5, CG)


def linha(ws, r, rotulo, nivel, mes_val, unidade=UN):
    """Writes one row: label, unit, YTD (SUMIFS) and the month cells; mes_val(k, col_letter) returns a
    value, a formula or None. nivel: 0 section, 1 block total, 2 subtotal, 3 item."""
    ws.cell(r, CD, rotulo)
    if nivel == 0:
        ws.cell(r, CD).font = Font(bold=True, size=11, color="1F4E79")
    elif nivel in (1, 2):
        ws.cell(r, CD).font = Font(bold=True)
    ws.cell(r, CD).alignment = Alignment(indent={0: 0, 1: 0, 2: 1, 3: 2}[nivel])
    if nivel == 1:
        for c in range(CD, CMES[-1] + 1):
            ws.cell(r, c).fill = CINZA
    if nivel > 0:
        ws.cell(r, CU, unidade)
    g = ws.cell(r, CG, f"=SUMIFS($K{r}:$V{r},$K$3:$V$3,$G$4)")
    g.number_format = FMT
    g.font = Font(bold=nivel < 3)
    for k, c in enumerate(CMES):
        v = mes_val(k, L(c))
        if v is None:
            continue
        cel = ws.cell(r, c, v)
        cel.number_format = FMT
        if nivel < 3:
            cel.font = Font(bold=True)


def valor(fonte, cols, grupo=None, classe=None):
    """Month value getter from a source: sum of cols, optionally for one (group, class)."""
    def f(k, _x):
        fr = DADOS[fonte][MESES[k]]
        if fr is None:
            return None
        if grupo is not None:
            fr = fr.loc[[i for i in fr.index if i == (grupo, classe)]]
        return round(float(fr[cols].sum().sum()), 9)
    return f


def soma(rows):
    """Formula getter: sum of given rows of the same month (only months up to the last closed one)."""
    def f(k, x):
        if MESES[k] > ULTIMO:
            return None
        return "=" + "+".join(f"{x}{r}" for r in rows)
    return f


def intervalo(r1, r2):
    def f(k, x):
        return None if MESES[k] > ULTIMO else f"=SUM({x}{r1}:{x}{r2})"
    return f


def bloco_cliente(ws, r, titulo, getter):
    """Total + Cativo/Livre with the 12 classes. getter(grupo, classe) -> month getter. Returns
    ({(grupo, classe) or 'TOTAL' or grupo: row}, next row)."""
    pos = {"TOTAL": r}
    rr = r + 1
    for g in GRUPOS:
        pos[g] = rr
        for i, cl in enumerate(CLASSES):
            pos[(g, cl)] = rr + 1 + i
        rr += 1 + len(CLASSES)
    linha(ws, r, titulo, 1, soma([pos[g] for g in GRUPOS]))
    for g in GRUPOS:
        linha(ws, pos[g], g, 2, intervalo(pos[(g, CLASSES[0])], pos[(g, CLASSES[-1])]))
        for cl in CLASSES:
            linha(ws, pos[(g, cl)], cl, 3, getter(g, cl))
    return pos, rr + 1


def secao_cliente(ws, r, titulo, cols, fontes=FONTES, rotulo="Receita"):
    """Receita Liquida | Por Cliente: Receita (= sum of the source blocks) and one block per source."""
    r_sec = r
    r_rec = r + 2
    n = 1 + len(GRUPOS) * (1 + len(CLASSES))
    blocos = {}
    rb = r_rec + n + 1
    for tit, fonte in fontes:
        blocos[tit], rb = bloco_cliente(ws, rb, tit, lambda g, cl, f=fonte: valor(f, cols, g, cl))
    pos_rec, _ = bloco_cliente(ws, r_rec, rotulo,
                               lambda g, cl: soma([blocos[t][(g, cl)] for t, _ in fontes]))
    linha(ws, r_sec, titulo, 0, soma([pos_rec["TOTAL"]]))
    return pos_rec, blocos, rb + 1


def bloco_componente(ws, r, titulo, itens, fonte):
    """Block total + one row per (label, cols) item from a source."""
    linha(ws, r, titulo, 1, intervalo(r + 1, r + len(itens)))
    for i, (rot, cols) in enumerate(itens):
        linha(ws, r + 1 + i, rot, 3, valor(fonte, cols))
    return r, r + len(itens) + 2


wb = openpyxl.Workbook()
wb.remove(wb.active)
CONF = {}                                         # rows used by the Conferencia sheet

# ------------------------------------------------ Receita
ws = wb.create_sheet("Receita")
prepara(ws, f"Receitas consolidadas {ANO} – Receita")
r_sec = 6
r_rec = 8
r = r_rec + len(COMPONENTES) + 2
fonte_rows = {}
for tit, fonte in FONTES:
    fonte_rows[tit], r = bloco_componente(ws, r, tit, COMPONENTES, fonte)
linha(ws, r_rec, "Receita", 1, intervalo(r_rec + 1, r_rec + len(COMPONENTES)))
for i, (rot, _) in enumerate(COMPONENTES):
    linha(ws, r_rec + 1 + i, rot, 2, soma([fonte_rows[t] + 1 + i for t, _ in FONTES]))
linha(ws, r_sec, "Receita Líquida | Por Componente", 0, soma([r_rec]))
CONF["comp_total"] = r_rec
CONF["comp_fontes"] = dict(fonte_rows)

# Demais ingressos: Faturado (FAT) and Nao Faturado (RNF)
r_di = r + 1
r_dr = r_di + 2
r = r_dr + len(DEMAIS) + 2
r_df, r = bloco_componente(ws, r, "Faturamento", DEMAIS, "FAT")
r_dn, r = bloco_componente(ws, r, "Não Faturado", DEMAIS, "RNF")
linha(ws, r_dr, "Receita", 1, intervalo(r_dr + 1, r_dr + len(DEMAIS)))
for i, (rot, _) in enumerate(DEMAIS):
    linha(ws, r_dr + 1 + i, rot, 2, soma([r_df + 1 + i, r_dn + 1 + i]))
linha(ws, r_di, "Demais ingressos", 0, soma([r_dr]))
CONF["demais_fat"], CONF["demais_rnf"] = r_df, r_dn
CONF["demais_total"] = r_dr

pos_rec, blocos, r = secao_cliente(ws, r + 1, "Receita Líquida | Por Cliente", COLS_COMP)
CONF["cli_total"] = pos_rec["TOTAL"]

# ------------------------------------------------ family sheets
for nome, cols in [COMPONENTES[1], COMPONENTES[0]] + COMPONENTES[2:]:
    wsf = wb.create_sheet(nome[:31])
    prepara(wsf, f"Receitas consolidadas {ANO} – {nome}")
    secao_cliente(wsf, 6, f"Receita Líquida {nome} | Por Cliente", cols)

# ------------------------------------------------ Subsidio
ws_s = wb.create_sheet("Subsídio")
prepara(ws_s, f"Receitas consolidadas {ANO} – Subsídio (Bl1 × Bl2: tarifa homologada - tarifa realizada)")
ORIGENS = [("Tarifa Social – DMR", "TS"), ("Fonte incentivada", "FI")]
r_rec = 8
r = r_rec + len(COMPONENTES) + 2
orig_rows = {}
for tit, fonte in ORIGENS:
    orig_rows[tit], r = bloco_componente(ws_s, r, tit, COMPONENTES, fonte)
linha(ws_s, r_rec, "Subsídio", 1, intervalo(r_rec + 1, r_rec + len(COMPONENTES)))
for i, (rot, _) in enumerate(COMPONENTES):
    linha(ws_s, r_rec + 1 + i, rot, 2, soma([orig_rows[t] + 1 + i for t, _ in ORIGENS]))
linha(ws_s, 6, "Subsídio | Por Componente", 0, soma([r_rec]))
secao_cliente(ws_s, r + 1, "Subsídio | Por Cliente", COLS_COMP, fontes=ORIGENS, rotulo="Subsídio")

# ------------------------------------------------ Receita P&C: one row per layout column
# tariff columns in the order of the Receita Faturada layout (Base Financeira before Base CVA)
COLS_PC = [c for c in COLS_COMP if c not in ("Base Financeira", "Base CVA")] + ["Base Financeira", "Base CVA"]


def bloco_pc(ws, r, titulo, getter):
    """Block total + Tarifavel subtotal (Parcela B to Base CVA) + its columns + Demais ingressos subtotal +
    its columns. getter(col) -> month getter. Returns ({col or subtotal key: row}, next row)."""
    pos = {"TOTAL": r, "TAR": r + 1}
    for i, c in enumerate(COLS_PC):
        pos[c] = r + 2 + i
    pos["DEM"] = r + 2 + len(COLS_PC)
    for i, c in enumerate(COLS_DEMAIS):
        pos[c] = pos["DEM"] + 1 + i
    linha(ws, r, titulo, 1, soma([pos["TAR"], pos["DEM"]]))
    linha(ws, pos["TAR"], "Tarifável (Parcela B a Base CVA)", 2, intervalo(pos[COLS_PC[0]], pos[COLS_PC[-1]]))
    for c in COLS_PC:
        linha(ws, pos[c], c, 3, getter(c))
    linha(ws, pos["DEM"], "Demais ingressos", 2, intervalo(pos[COLS_DEMAIS[0]], pos[COLS_DEMAIS[-1]]))
    for c in COLS_DEMAIS:
        linha(ws, pos[c], c, 3, getter(c))
    return pos, pos[COLS_DEMAIS[-1]] + 2


ws_p = wb.create_sheet("Receita P&C")
prepara(ws_p, f"Receitas consolidadas {ANO} – Receita por coluna do layout (P&C)")
n_pc = 2 + len(COLS_PC) + 1 + len(COLS_DEMAIS)
r_pc = 8
r = r_pc + n_pc + 1
pc = {}
for tit, fonte in FONTES:
    pc[tit], r = bloco_pc(ws_p, r, tit, lambda c, f=fonte: valor(f, [c]))
pos_pc, _ = bloco_pc(ws_p, r_pc, "Receita", lambda c: soma([pc[t][c] for t, _ in FONTES]))
linha(ws_p, 6, "Receita | Por coluna do layout Receita Faturada", 0, soma([pos_pc["TOTAL"]]))
CONF["pc_total"] = pos_pc["TOTAL"]

# one sheet per layout column that has no sheet yet, Por Cliente
ITENS_ABA = ["Rede Básica", "Transporte Itaipu", "CUSD", "Conexão T", "CDE USO", "CDE CONTAS", "CDE ELET",
             "PROINFA", "ESS/ERR/ERCAP", "ONS", "P&D", "TFSEE"] + COLS_DEMAIS
ABAS_ITEM = []
for c in ITENS_ABA:
    wsi = wb.create_sheet(c.replace("/", "-")[:31])
    ABAS_ITEM.append(wsi.title)
    prepara(wsi, f"Receitas consolidadas {ANO} – {c}")
    secao_cliente(wsi, 6, f"Receita {c} | Por Cliente", [c])

# ------------------------------------------------ Leia-me
ws_l = wb.create_sheet("Leia-me", 0)
TEXTO = [
    f"Receitas consolidadas {ANO}",
    "",
    "Fontes (por mês)",
    "Faturamento: 1-Entrega/receita_faturada/receita_faturada_<mês>.xlsx, aba Receita Faturada, FAT + REFAT + CNR.",
    "Subsídio: 2-Checks/receita_faturada_Bl1_vs_Bl2/receita_faturada_Bl1_vs_Bl2_<mês>.xlsx, abas desconto_tarifa_social_dmr "
    "e desconto_fonte_incentivada (tarifa homologada - tarifa realizada; positivo = desconto).",
    "RNF: RNF/1-Entrega/receita_nao_faturada/receita_nao_faturada_<mês>.xlsx, aba Saldo (provisão do mês + estorno da provisão anterior).",
    "",
    "Regras",
    "Valores em MR$ (milhões de reais).",
    "Componentes: Transmissão = Rede Básica + Transporte Itaipu + CUSD + Conexão T; Encargos = CDE USO + CDE CONTAS + CDE ELET + "
    "PROINFA + ESS/ERR/ERCAP + ONS + P&D + TFSEE; CVA = Base CVA; Ativos e Passivos Regulatórios = Base Financeira.",
    "Receita Líquida = Faturamento + Subsídio + RNF dos componentes; Covid Migrantes, Bandeira, UD/ER, Outros, PIS/COFINS e ICMS "
    "ficam em Demais ingressos.",
    "Covid Migrantes = ZQC237, encargo da Conta-COVID cobrado à parte dos consumidores que migraram para o mercado livre (não tarifável).",
    "Outros = itens não tarifários da coluna Outros da receita_faturada (detalhados na aba Conferencia de cada mês).",
    "Classes: CATIVO e SCEE nas classes do Cativo; SUPRIMENTO, USO D e GERAÇÃO no Cativo (Suprimento, Uso Distribuição e Uso Geração); "
    "LIVRE nas classes do Livre.",
    "Linha 3: YTD até o último mês com receita_faturada, YTG depois. Coluna YTD = SUMIFS dos meses marcados YTD.",
    "Meses sem fonte ficam em branco.",
    "",
    "Correções em relação ao template",
    "Totais de bloco Por Cliente = Cativo + Livre (no template, SUM das 9 primeiras linhas, que somava Cativo com classes e "
    "deixava de fora parte do Livre).",
    "Cativo e Livre = soma das 12 classes.",
    "Receita Líquida | Por Componente = linha Receita (template: SUM de Faturamento, Subsídio e RNF, mesmo resultado).",
    "Receita Líquida | Por Cliente = linha Receita do bloco (template apontava para linhas que não eram totais).",
    "Demais ingressos | Receita = soma das linhas; cada linha = Faturamento + Não Faturado (template: SUM de um intervalo errado "
    "e vínculo externo [1]Base).",
    "Linhas Covid Migrantes e Outros acrescentadas em Demais ingressos; Não Faturado com as mesmas 6 linhas do Faturamento.",
    "YTD (coluna G) com SUMIFS em todas as linhas, inclusive totais.",
    "",
    "Abas adicionais",
    "Compra de Energia, Perdas, Transmissão, Encargos, CVA e Ativos e Passivos Regulatórios: mesmo layout das abas RI e Parcela B.",
    "Subsídio: Tarifa Social – DMR e Fonte incentivada, por componente e por cliente.",
    "Receita P&C: uma linha por coluna do layout Receita Faturada (Parcela B a ICMS), com os subtotais Tarifável e "
    "Demais ingressos; blocos Receita = Faturamento + Subsídio + RNF.",
    "Uma aba por coluna do layout que não tem aba de família (Rede Básica a TFSEE e Demais ingressos), no layout "
    "Por Cliente; nos nomes das abas, '/' vira '-'.",
    "Conferencia: cada bloco contra a sua fonte e Por Componente contra Por Cliente.",
]
for i, t in enumerate(TEXTO, start=2):
    ws_l.cell(i, 2, t)
    if i == 2 or t in ("Fontes (por mês)", "Regras", "Correções em relação ao template", "Abas adicionais"):
        ws_l.cell(i, 2).font = Font(bold=True)
ws_l.column_dimensions["A"].width = 2
ws_l.column_dimensions["B"].width = 160

# ------------------------------------------------ Conferencia
ws_c = wb.create_sheet("Conferencia")
cab = ["Mês",
       "Faturamento: layout (MR$)", "Faturamento: Conferencia do arquivo (MR$)", "Diferença",
       "Subsídio: origens (MR$)", "Subsídio: homologada - realizada (MR$)", "Diferença",
       "Subsídio em Demais ingressos (deve ser zero)",
       "RNF: Saldo (MR$)", "RNF: Provisão + Estorno (MR$)", "Diferença",
       "Por Componente - Por Cliente (deve ser zero)",
       "P&C - (Por Componente + Demais ingressos) (deve ser zero)"]
ws_c.append(cab)
for k, ref in enumerate(MESES):
    lin = [ref]
    fat = DADOS["FAT"][ref]
    if fat is not None:
        conf = pd.read_excel(DIR_RF / f"receita_faturada_{ref}.xlsx", sheet_name="Conferencia", header=None)
        grav = float(conf.loc[conf[0].astype(str).str.startswith("Gravado no layout"), 1].iloc[0]) / 1e6
        s = float(fat[COLS].sum().sum())
        lin += [s, grav, s - grav]
    else:
        lin += [None] * 3
    if DADOS["SUB"][ref] is not None:
        arq = DIR_BL / f"receita_faturada_Bl1_vs_Bl2_{ref}.xlsx"
        h = agrega(pd.read_excel(arq, sheet_name="tarifa_homologada", header=2), "homologada")
        p = agrega(pd.read_excel(arq, sheet_name="tarifa_realizada", header=2), "realizada")
        s = float(DADOS["SUB"][ref][COLS_COMP].sum().sum())
        dif = float(h[COLS_COMP].sum().sum() - p[COLS_COMP].sum().sum())
        lin += [s, dif, s - dif, float(DADOS["SUB"][ref][COLS_DEMAIS].abs().sum().sum())]
    else:
        lin += [None] * 4
    if DADOS["RNF"][ref] is not None:
        arq = DIR_RNF / f"receita_nao_faturada_{ref}.xlsx"
        pe = sum(float(agrega(pd.read_excel(arq, sheet_name=a, header=2), a)[COLS].sum().sum())
                 for a in ("Provisao", "Estorno"))
        s = float(DADOS["RNF"][ref][COLS].sum().sum())
        lin += [s, pe, s - pe]
    else:
        lin += [None] * 3
    x = L(CMES[k])
    lin.append(f"=Receita!{x}{CONF['comp_total']}-Receita!{x}{CONF['cli_total']}" if ref <= ULTIMO else None)
    lin.append(f"='Receita P&C'!{x}{CONF['pc_total']}-Receita!{x}{CONF['comp_total']}-Receita!{x}{CONF['demais_total']}"
               if ref <= ULTIMO else None)
    ws_c.append(lin)
    for i in (3, 6, 10):
        if lin[i] is not None and abs(lin[i]) > 0.001:
            print(f"ATENCAO {ref}: {cab[i]} {lin[i]:,.3f} MR$")
for row in ws_c.iter_rows(min_row=2):
    for cel in row[1:]:
        cel.number_format = "#,##0.000;-#,##0.000;\"-\""
for c in ws_c[1]:
    c.font = Font(bold=True)
    c.alignment = Alignment(wrap_text=True, vertical="top")
ws_c.column_dimensions["A"].width = 10
for c in range(2, len(cab) + 1):
    ws_c.column_dimensions[L(c)].width = 18
ws_c.row_dimensions[1].height = 60

# sheet order: Leia-me, Receita, Receita P&C, RI, Parcela B, other families, item sheets, Subsidio, Conferencia
ordem = (["Leia-me", "Receita", "Receita P&C"] + [n[:31] for n, _ in [COMPONENTES[1], COMPONENTES[0]] + COMPONENTES[2:]]
         + ABAS_ITEM + ["Subsídio", "Conferencia"])
wb._sheets = [wb[n] for n in ordem]
DIR_SAIDA.mkdir(parents=True, exist_ok=True)
wb.save(SAIDA)
falta = {f: [r for r in com_fat if DADOS[f][r] is None] for f in ("SUB", "RNF")}
print(f"meses: {', '.join(com_fat)} | sem subsidio: {falta['SUB'] or '-'} | sem RNF: {falta['RNF'] or '-'}")
print(f"saida: {SAIDA}")
