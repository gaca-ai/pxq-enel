# -*- coding: utf-8 -*-
"""
Extrator da base das receitas consolidadas: uma tabela longa (camada Gold) para Power Query / BI.

Uso:
    python extrator_receitas_consolidadas.py 2026
Le, por mes com receita_faturada:
    1-Entrega/pxq/pxq_<ref>.xlsx                    Faturamento linha a linha (pelo conversor do layout)
    1-Entrega/pxq_regulacao/PxQ_regulacao_<ref>.xlsx camadas CVA e FIN de cada item tarifario
    1-Entrega/receita_faturada/receita_faturada_<ref>.xlsx    conferencia do Faturamento
    2-Checks/receita_faturada_Bl1_vs_Bl2/receita_faturada_Bl1_vs_Bl2_<ref>.xlsx    Subsidio
    ../RNF/1-Entrega/receita_nao_faturada/receita_nao_faturada_<ref>.xlsx (aba Saldo)  RNF
Grava
    1-Entrega/receitas_consolidadas/receitas_consolidadas_base_<ano>.csv
    1-Entrega/receitas_consolidadas/receitas_consolidadas_base_<ano>_conferencia.xlsx

Regras:
- Uma so tabela para os tres blocos. Cada linha e preenchida ate o nivel que a fonte tem; os niveis
  abaixo ficam "(não aberto)".
  * Faturamento: abertura maxima. Linha do BW/SD (codigo e texto do lancamento, subclasse, categoria
    tarifaria) x item tarifario x camada (BE, CVA, FIN); nao tarifaveis por lancamento. A
    classificacao (opcao, classe, unidade, tipo de faturamento), o de-para item -> coluna, o residuo
    e o rateio D/E dos nao tarifaveis sao os do proprio conversor (gera_receita_faturada.converte,
    parametro detalhe): somada no nivel do layout, a base e igual a receita_faturada.
  * Subsidio (Tarifa Social - DMR e Fonte incentivada) e RNF (aba Saldo, por tipo de RNF): nivel do
    layout (opcao x classe x unidade x coluna).
- Camada: BE de cada item vai para a coluna do layout do item; CVA e FIN vao para Base CVA e Base
  Financeira. No Subsidio e no RNF, pela coluna: Base Financeira = FIN, Base CVA = CVA, demais
  tarifaveis = BE; nao tarifaveis = "Não tarifável". Linhas do SD so com TA (sem camadas) entram em BE,
  como no layout.
- Colunas de agregacao: coluna da receita_faturada, componente do template Receitas Consolidadas,
  familia do consolidado por familia, natureza (Tarifável ou Demais ingressos), grupo e classe do
  template; no Faturamento tambem familia PxQ, parcela (TUSD/TE) e coluna do Mercado Receita, pelo
  de-para PxQ_regulacao x MR (3-Py/2-Checks/gera_de_para_regulacao_MR.py, ITENS).
- Medidas: Valor (MR$); Energia (MWh) nas linhas E e Demanda (MW) nas linhas D (= CONSUMO do layout);
  Demanda de ultrapassagem (MW). Valores zerados nao sao gravados.
- CSV: separador ";", virgula decimal, UTF-8 (com BOM), Competencia em dd/mm/aaaa.
- Conferencia: por mes e bloco, base no nivel do layout x fonte, celula a celula (MR$).
"""
import os
import sys
from pathlib import Path

import numpy as np
import pandas as pd

BASE = Path(os.environ.get("PXQ_BASE", r"C:\python_projects\claude\p_x_q\SP\BW"))
sys.path.insert(0, str(BASE / "3-Py" / "1-Entrega"))
sys.path.insert(0, str(BASE / "3-Py" / "2-Checks"))
import gera_receita_faturada as grf                                   # noqa: E402
from gera_de_para_regulacao_MR import ITENS as DE_PARA_MR             # noqa: E402

ANO = int(sys.argv[1]) if len(sys.argv) > 1 else pd.Timestamp.today().year
DIR_RF = BASE / "1-Entrega" / "receita_faturada"
DIR_REG = BASE / "1-Entrega" / "pxq_regulacao"
DIR_BL = BASE / "2-Checks" / "receita_faturada_Bl1_vs_Bl2"
DIR_RNF = BASE.parent / "RNF" / "1-Entrega" / "receita_nao_faturada"
DIR_SAIDA = BASE / "1-Entrega" / "receitas_consolidadas"
SAIDA = DIR_SAIDA / f"receitas_consolidadas_base_{ANO}.csv"
SAIDA_CONF = DIR_SAIDA / f"receitas_consolidadas_base_{ANO}_conferencia.xlsx"
NA = "(não aberto)"
MR = 1e6                                                              # R$ -> MR$

# layout column -> (component of the Receitas Consolidadas template, family of the family consolidation)
COLUNAS = {
    "Parcela B": ("Parcela B", "Parcela B"), "RI": ("RI", "RI"),
    "Compra de Energia": ("Compra de Energia", "Compra de Energia"), "Perdas": ("Perdas", "Perdas"),
    "Rede Básica": ("Transmissão", "Transmissão"), "Transporte Itaipu": ("Transmissão", "Transmissão"),
    "CUSD": ("Transmissão", "Transmissão"), "Conexão T": ("Transmissão", "Transmissão"),
    "CDE USO": ("Encargos", "Encargos"), "CDE CONTAS": ("Encargos", "Encargos"),
    "CDE ELET": ("Encargos", "Encargos"), "PROINFA": ("Encargos", "Encargos"),
    "ESS/ERR/ERCAP": ("Encargos", "Encargos"), "ONS": ("Encargos", "Encargos"),
    "P&D": ("Encargos", "Encargos"), "TFSEE": ("Encargos", "Encargos"),
    "Base Financeira": ("Ativos e Passivos Regulatórios", "Financeiros"), "Base CVA": ("CVA", "CVA"),
    "Covid Migrantes": ("Demais ingressos", "UD-ER e Outros"), "Bandeira": ("Demais ingressos", "Bandeira"),
    "UD/ER": ("Demais ingressos", "UD-ER e Outros"), "Outros": ("Demais ingressos", "UD-ER e Outros"),
    "PIS/COFINS": ("Demais ingressos", "Tributos"), "ICMS": ("Demais ingressos", "Tributos"),
}
COLS = list(COLUNAS)
TARIFAVEIS = COLS[:COLS.index("Base CVA") + 1]
# template group and class (same rules as receitas_consolidadas.py)
DE_PARA_CLASSE = {"Residencial": "Residencial", "Residencial Baixa Renda": "Residencial Baixa Renda",
                  "Industrial": "Industrial", "Comercial": "Comercial", "Rural": "Rural",
                  "Poderes Públicos": "Poder Público", "Iluminação Pública": "Iluminação Pública",
                  "Serviços Públicos": "Serviço Público", "Consumo Próprio": "Consumo Próprio"}
MERCADO_CATIVO = {"SUPRIMENTO": "Suprimento", "USO D": "Uso Distribuição", "GERAÇÃO": "Uso Geração"}
OPCOES_CATIVO = {"CATIVO", "SCEE - GD1", "SCEE - GD2"}
CAMPOS = ["Competência", "Bloco", "Origem", "Fonte", "Opção", "Grupo", "Classe layout", "Classe template",
          "Subclasse", "Categoria tarifa", "Código lançamento", "Lançamento", "Unidade", "Item", "Parcela",
          "Camada", "Coluna receita_faturada", "Componente template", "Família consolidado", "Família PxQ",
          "Coluna MR", "Natureza", "Medida", "Unid. medida", "Valor"]


def grupo_classe(op, cl):
    """(template group, template class) of a layout option x class."""
    if op in OPCOES_CATIVO:
        return "Cativo", DE_PARA_CLASSE.get(cl, NA)
    if op in MERCADO_CATIVO:
        return "Cativo", MERCADO_CATIVO[op]
    if op == "LIVRE":
        return "Livre", DE_PARA_CLASSE.get(cl, NA)
    return NA, NA


def camada_coluna(col):
    if col == "Base Financeira":
        return "FIN"
    if col == "Base CVA":
        return "CVA"
    return "BE" if col in TARIFAVEIS else "Não tarifável"


def completa(df):
    """Adds the aggregation columns that depend only on the layout column and the option x class."""
    gc = [grupo_classe(o, c) for o, c in zip(df["Opção"], df["Classe layout"])]
    df["Grupo"], df["Classe template"] = [g for g, _ in gc], [c for _, c in gc]
    q = df["Medida"] != "Valor"
    df["Componente template"] = np.where(q, NA, df["Coluna receita_faturada"].map(lambda c: COLUNAS.get(c, (NA,))[0]))
    df["Família consolidado"] = np.where(q, NA, df["Coluna receita_faturada"].map(lambda c: COLUNAS.get(c, (NA, NA))[1]))
    df["Natureza"] = np.where(q, NA, np.where(df["Coluna receita_faturada"].isin(TARIFAVEIS), "Tarifável",
                                              "Demais ingressos"))
    return df


# ------------------------------------------------ Faturamento (row level, through the converter)
def faturamento(ref, modelo):
    det = {}
    grf.converte(ref, *modelo, grava=False, detalhe=det)
    _, cab, de_para, _ = modelo
    t, nt, cols_comp, so_ta, tar_ta = det["tar"], det["nt"], det["cols_comp"], det["so_ta"], det["tar_ta"]
    pre = "[2] "
    reg = pd.read_excel(DIR_REG / f"PxQ_regulacao_{ref}.xlsx")
    if not (reg.loc[t.index, "Chave"].astype(str).values == t["Chave"].astype(str).values).all():
        raise ValueError(f"{ref}: PxQ_regulacao fora da ordem das linhas do pxq (regerar o notebook)")
    ident = pd.DataFrame({
        "Origem": t["TIPO"], "Fonte": t["Source"], "Opção": t["OPCAO"], "Classe layout": t["CLASSE"],
        "Subclasse": t["Classe/Sub-classe"].fillna(NA), "Categoria tarifa": t["Categoria tarifa"].fillna(NA),
        "Código lançamento": t["Grp.estatísts.quant."].fillna(NA), "Lançamento": t["Grp.estatísts.mont."].fillna(NA),
        "Unidade": t["UNIDADE"]}, index=t.index)
    partes = []
    itens = [c[len(pre):] for c in cols_comp]
    sota = so_ta.loc[t.index]
    for cam in ("BE", "CVA", "FIN"):
        if cam == "BE":
            v = t[cols_comp].fillna(0.0).copy()
            v.loc[sota] = tar_ta.loc[sota, cols_comp].fillna(0.0).values     # SD rows with only TA
            v.columns = itens
        else:
            v = reg.loc[t.index, [f"{cam} {i}" for i in itens]].fillna(0.0)
            v.columns = itens
            v.loc[sota] = 0.0
        lg = v.stack()
        lg = lg[lg.abs() > 1e-9]
        lg.index.names = ["_i", "Item"]
        lg = lg.rename("Valor").reset_index()
        lg["Camada"] = cam
        partes.append(lg)
    it = pd.concat(partes, ignore_index=True)
    it = it.join(ident, on="_i").drop(columns="_i")
    esp = {i: DE_PARA_MR.get(i) for i in it["Item"].unique()}
    pos = {"BE": 2, "CVA": 3, "FIN": 4}
    it["Coluna receita_faturada"] = [de_para.get(i, "(sem coluna no modelo)") if c == "BE" else
                                     ("Base CVA" if c == "CVA" else "Base Financeira")
                                     for i, c in zip(it["Item"], it["Camada"])]
    it["Parcela"] = [esp[i][1] if esp[i] else NA for i in it["Item"]]
    it["Família PxQ"] = [esp[i][0] if esp[i] else NA for i in it["Item"]]
    it["Coluna MR"] = [(esp[i][pos[c]] or "(sem coluna no MR)") if esp[i] else "(sem item no de-para)"
                       for i, c in zip(it["Item"], it["Camada"])]
    it["Valor"] = it["Valor"] / MR

    # residue (revenue - TA of the delivery) -> Outros, only groups above R$ 0.005, as in the layout
    chave = ["TIPO", "OPCAO", "CLASSE", "UNIDADE"]
    ok = det["residuo"][det["residuo"].abs() > 0.005].index
    no_grupo = pd.Series([k in ok for k in zip(*[t[c] for c in chave])], index=t.index)
    rs = t[no_grupo & (t["RESIDUO"].abs() > 1e-9)]
    res = ident.loc[rs.index].assign(Item="(resíduo sem item tarifário)", Camada="Não tarifável",
                                     **{"Coluna receita_faturada": grf.COL_OUTROS, "Parcela": NA,
                                        "Família PxQ": NA, "Coluna MR": "(sem coluna no MR)",
                                        "Valor": rs["RESIDUO"] / MR})

    # non-tariff lines, split between D and E with the converter's own shares
    lin = []
    for _, r in nt.iterrows():
        for u, p in det["reparte"](r["TIPO"], r["OPCAO"], r["CLASSE"], r["PARCELA"]).items():
            lin.append({"Origem": r["TIPO"], "Fonte": "SD" if r["OPCAO"] in grf.OPCOES_SD else "BW",
                        "Opção": r["OPCAO"], "Classe layout": r["CLASSE"],
                        "Subclasse": r["Classe/Sub-classe"] if pd.notna(r["Classe/Sub-classe"]) else NA,
                        "Categoria tarifa": r["Categoria tarifa"] if pd.notna(r["Categoria tarifa"]) else NA,
                        "Código lançamento": r["Grp.estatísts.quant."], "Lançamento": r["Grp.estatísts.mont."],
                        "Unidade": u, "Item": "(não tarifável)", "Parcela": NA, "Camada": "Não tarifável",
                        "Coluna receita_faturada": r["COLUNA"], "Família PxQ": NA,
                        "Coluna MR": "(sem coluna no MR)", "Valor": r["RECEITA DE FORNECIMENTO"] * p / MR})
    ntl = pd.DataFrame(lin)

    # quantities: CONSUMO of the layout (E rows MWh, D rows MW) and demand overrun (MW)
    q = ident.assign(Item=NA, Parcela=NA, Camada=NA, **{"Coluna receita_faturada": NA, "Família PxQ": NA,
                     "Coluna MR": NA}, Valor=t["CONSUMO"])
    q["Medida"] = np.where(q["Unidade"] == "D", "Demanda (MW)", "Energia (MWh)")
    q["Unid. medida"] = np.where(q["Unidade"] == "D", "MW", "MWh")
    ud = nt[nt["Grp.estatísts.mont."].astype(str).str.upper().str.contains(grf.PADRAO_UD_MW)]
    qud = pd.DataFrame({"Origem": ud["TIPO"], "Fonte": np.where(ud["OPCAO"].isin(grf.OPCOES_SD), "SD", "BW"),
                        "Opção": ud["OPCAO"], "Classe layout": ud["CLASSE"],
                        "Subclasse": ud["Classe/Sub-classe"].fillna(NA), "Categoria tarifa": ud["Categoria tarifa"].fillna(NA),
                        "Código lançamento": ud["Grp.estatísts.quant."], "Lançamento": ud["Grp.estatísts.mont."],
                        "Unidade": "D", "Item": NA, "Parcela": NA, "Camada": NA, "Coluna receita_faturada": NA,
                        "Família PxQ": NA, "Coluna MR": NA, "Medida": "Demanda de ultrapassagem (MW)",
                        "Unid. medida": "MW", "Valor": ud.get("Demanda (kW)", 0.0) / 1000})
    val = pd.concat([it, res, ntl], ignore_index=True).assign(Medida="Valor", **{"Unid. medida": "MR$"})
    out = pd.concat([val, q, qud], ignore_index=True)
    out = out[out["Valor"].abs() > 1e-12]
    return out.assign(Bloco="Faturamento")


# ------------------------------------------------ Subsidio and RNF (layout level)
def layout_longo(arq, aba, bloco, origem=None):
    d = pd.read_excel(arq, sheet_name=aba, header=2)
    d = d[d["OpcaoCompra"].notna() & d["TipoFaturamento"].notna()]
    falta = [c for c in COLS if c not in d.columns]
    if falta:
        raise KeyError(f"{arq.name}/{aba}: colunas ausentes (regerar o arquivo): {falta}")
    ids = ["TipoFaturamento", "OpcaoCompra", "Classe", "Unidade"]
    for c in COLS + ["CONSUMO"]:
        d[c] = pd.to_numeric(d[c], errors="coerce").fillna(0.0)
    v = d.melt(id_vars=ids, value_vars=COLS, var_name="Coluna receita_faturada", value_name="Valor")
    v["Valor"] = v["Valor"] / 1000.0                                     # R$ mil -> MR$
    v["Medida"], v["Unid. medida"] = "Valor", "MR$"
    v["Camada"] = v["Coluna receita_faturada"].map(camada_coluna)
    q = d[ids + ["CONSUMO"]].rename(columns={"CONSUMO": "Valor"})
    q["Medida"] = np.where(q["Unidade"] == "D", "Demanda (MW)", "Energia (MWh)")
    q["Unid. medida"] = np.where(q["Unidade"] == "D", "MW", "MWh")
    q["Camada"] = q["Coluna receita_faturada"] = NA
    out = pd.concat([v, q], ignore_index=True)
    out = out[out["Valor"].abs() > 1e-12].rename(columns={"OpcaoCompra": "Opção", "Classe": "Classe layout"})
    out["Origem"] = origem if origem else out["TipoFaturamento"]
    out = out.drop(columns="TipoFaturamento")
    for c in ("Fonte", "Subclasse", "Categoria tarifa", "Código lançamento", "Lançamento", "Item", "Parcela",
              "Família PxQ", "Coluna MR"):
        out[c] = NA
    return out.assign(Bloco=bloco)


def layout_fonte(arq, aba, tipos=None):
    """Layout of a source file at option x class x unit x column (MR$), for the reconciliation."""
    d = pd.read_excel(arq, sheet_name=aba, header=2)
    d = d[d["OpcaoCompra"].notna() & d["TipoFaturamento"].notna()]
    if tipos:
        d = d[d["TipoFaturamento"].isin(tipos)]
    m = d.melt(id_vars=["OpcaoCompra", "Classe", "Unidade"], value_vars=COLS, var_name="Coluna", value_name="v")
    m["v"] = pd.to_numeric(m["v"], errors="coerce").fillna(0.0) / 1000.0
    return m.groupby(["OpcaoCompra", "Classe", "Unidade", "Coluna"])["v"].sum()


def nivel_layout(b):
    v = b[b["Medida"] == "Valor"]
    return v.groupby(["Opção", "Classe layout", "Unidade", "Coluna receita_faturada"])["Valor"].sum()


def compara(base, fonte):
    k = base.index.union(fonte.index)
    d = base.reindex(k, fill_value=0.0) - fonte.reindex(k, fill_value=0.0)
    return float(base.sum()), float(fonte.sum()), float(d.abs().max()) if len(d) else 0.0


modelo = grf.le_modelo()
meses = [f"{ANO}-{m:02d}" for m in range(1, 13) if (DIR_RF / f"receita_faturada_{ANO}-{m:02d}.xlsx").exists()]
if not meses:
    sys.exit(f"nenhum receita_faturada de {ANO} em {DIR_RF}")
blocos, conf = [], []
for ref in meses:
    comp = pd.Timestamp(f"{ref}-01")
    fat = faturamento(ref, modelo).assign(Competência=comp)
    conf.append([ref, "Faturamento", *compara(nivel_layout(fat), layout_fonte(DIR_RF / f"receita_faturada_{ref}.xlsx",
                                                                            "Receita Faturada"))])
    blocos.append(fat)
    arq = DIR_BL / f"receita_faturada_Bl1_vs_Bl2_{ref}.xlsx"
    if arq.exists():
        sub = pd.concat([layout_longo(arq, "desconto_tarifa_social_dmr", "Subsídio", "Tarifa Social – DMR"),
                         layout_longo(arq, "desconto_fonte_incentivada", "Subsídio", "Fonte incentivada")])
        sub = sub[sub["Medida"] == "Valor"].assign(Competência=comp)   # the discount has no quantity of its own
        fonte = layout_fonte(arq, "desconto_tarifa_social_dmr").add(layout_fonte(arq, "desconto_fonte_incentivada"),
                                                                    fill_value=0.0)
        conf.append([ref, "Subsídio", *compara(nivel_layout(sub), fonte)])
        blocos.append(sub)
    else:
        print(f"ATENCAO {ref}: sem Bl1 x Bl2 ({arq.name})")
    arq = DIR_RNF / f"receita_nao_faturada_{ref}.xlsx"
    if arq.exists():
        rnf = layout_longo(arq, "Saldo", "RNF").assign(Competência=comp)
        conf.append([ref, "RNF", *compara(nivel_layout(rnf), layout_fonte(arq, "Saldo"))])
        blocos.append(rnf)
    else:
        print(f"ATENCAO {ref}: sem RNF ({arq.name})")
    print(f"{ref}: Faturamento {len(fat):,} linhas")

base = pd.concat(blocos, ignore_index=True)
base = completa(base)[CAMPOS]
conf = pd.DataFrame(conf, columns=["Mês", "Bloco", "Base (MR$)", "Fonte (MR$)", "Maior diferença por célula (MR$)"])
conf["Diferença total (MR$)"] = conf["Base (MR$)"] - conf["Fonte (MR$)"]
for r in conf.itertuples(index=False):
    if abs(r[4]) > 1e-6 or abs(r[5]) > 1e-6:
        print(f"ATENCAO {r[0]} {r[1]}: base x fonte, maior diferença {r[4]:.6f} MR$, total {r[5]:.6f} MR$")

DIR_SAIDA.mkdir(parents=True, exist_ok=True)
base.to_csv(SAIDA, sep=";", decimal=",", index=False, encoding="utf-8-sig", date_format="%d/%m/%Y",
            float_format="%.10f")
with pd.ExcelWriter(SAIDA_CONF, engine="openpyxl") as w:
    conf.to_excel(w, sheet_name="Conferencia", index=False)
    resumo = (base[base["Medida"] == "Valor"].assign(Mês=base["Competência"].dt.strftime("%Y-%m"))
              .pivot_table(index="Mês", columns="Bloco", values="Valor", aggfunc="sum"))
    resumo.to_excel(w, sheet_name="Total por bloco (MR$)")
print(f"{len(base):,} linhas | {SAIDA}")
print(f"conferencia: {SAIDA_CONF}")
