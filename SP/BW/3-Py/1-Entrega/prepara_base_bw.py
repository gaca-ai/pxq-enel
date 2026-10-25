"""
Prepara a extracao mensal de receita do BW: csv original -> parquet limpo.

Uso (a partir de SP/BW):
    python 3-Py/1-Entrega/prepara_base_bw.py              # todos os meses pendentes
    python 3-Py/1-Entrega/prepara_base_bw.py --forcar     # refaz tambem os parquets ja existentes

Pastas:
    src/csv_original/   entrada: MM-AAAA_ENEL SP_RECEITA.csv (novo) ou .zip (ja compactado)
    src/parquet/        saida:   MM-AAAA_ENEL SP_RECEITA.parquet

Para cada mes pendente (csv novo, ou zip sem parquet):
    1. le o csv (separador ";", UTF-8, 4 linhas de cabecalho, numeros no formato brasileiro);
    2. grava o parquet limpo: nomes de coluna unicos, numeros e datas convertidos, so a coluna
       "Resultado" de cada medida, sem a linha de total;
    3. confere as cinco medidas do parquet contra a linha "Result.global" do csv (montante ao
       centavo; quantidades com tolerancia de 1 parte em 10 mil, pelo arredondamento do BW). Quando o BW
       imprime "*" no Resultado (unidades misturadas), o valor e a soma das colunas kW, kW.h e
       Nao atribuido da medida;
    4. so se a conferencia passar: compacta o csv em src/csv_original/<mesmo nome>.zip e apaga o csv.
       Se a conferencia falhar, o csv fica onde esta e o parquet nao e gravado.
"""
import sys
import zipfile
from pathlib import Path
import os

import numpy as np
import pandas as pd
import pyarrow as pa
import pyarrow.parquet as pq

BASE = Path(os.environ.get("PXQ_BASE", r"C:\python_projects\claude\p_x_q\SP\BW"))
DIR_CSV = BASE / "src" / "csv_original"
DIR_PARQUET = BASE / "src" / "parquet"

LINHAS_CABECALHO = 4                        # 3 header rows of the BW export + the column names row
TAMANHO_BLOCO = 200_000                     # rows per chunk (the csv has 200+ MB)

# column position in the csv -> clean name (dimensions, as text)
DIMENSOES = {
    0: "Classe principal",
    1: "Categoria tarifa",
    2: "Grp.estatísts.mont.",
    3: "Grp.estatísts.quant.",
    4: "Grp.estatísts.quant. TEXTO",
    5: "Operação principal",
    6: "Operação principal TEXTO",
    7: "Suboperação",
    8: "Suboperação TEXTO",
    9: "Flag Cli.Liv./ Cat.",
    10: "Classe/Sub-classe",
    11: "Conjunto de Contrato",
    12: "Mês de Referência",
    13: "Flag de Refaturamento",
    14: "Data de Faturamento",
}
# "Resultado" column of each measure (the kW / kW.h / Não atribuído columns are its breakdown)
MEDIDAS = {
    18: "TUSD Faturado (kWh)",
    22: "TUSD Faturado (kW)",
    26: "Demanda Faturada (kW)",
    30: "Consumo TE Faturado (kWh)",
    34: "Montante (R$)",
}
# each "Resultado" is the sum of the three columns before it (kW, kW.h, Não atribuído); the BW prints
# "*" in "Resultado" when the units are mixed, and then the value is rebuilt from those columns
DETALHE = {c: [c - 3, c - 2, c - 1] for c in MEDIDAS}
TOTAL = "Result.global"


def numero_br(s):
    """'1.234,56' -> 1234.56; '6.481.782 KW' -> 6481782.0 (the BW appends the unit to some cells);
    blank -> 0.0."""
    s = s.fillna("").astype(str).str.replace(r"[^0-9,\-]", "", regex=True).str.replace(",", ".", regex=False)
    return pd.to_numeric(s.replace("", "0"), errors="coerce").fillna(0.0)


def nao_numericos(s):
    """Non-blank cells that are not a number even after removing the unit (data loss guard)."""
    t = s.fillna("").astype(str).str.strip()
    limpo = t.str.replace(r"[^0-9,\-]", "", regex=True)
    return int(((t != "") & (pd.to_numeric(limpo.str.replace(",", ".", regex=False), errors="coerce").isna())).sum())


def limpa(bloco):
    """Raw chunk (text) -> clean frame, plus the total row when present in the chunk."""
    for col, nome in MEDIDAS.items():
        estrela = bloco[col].fillna("").astype(str).str.strip() == "*"
        valor = numero_br(bloco[col])
        if estrela.any():
            valor = valor.where(~estrela, sum(numero_br(bloco[d]) for d in DETALHE[col]))
        bloco[col] = np.where(estrela, valor.astype(str), bloco[col].fillna("").astype(str))
        bloco[f"__{nome}"] = valor
    bloco = bloco.drop(columns=[d for ds in DETALHE.values() for d in ds])
    bloco = bloco.rename(columns={**DIMENSOES, **MEDIDAS})
    invalidos = sum(nao_numericos(bloco[c]) for c in MEDIDAS.values())
    for c in MEDIDAS.values():
        bloco[c] = bloco.pop(f"__{c}")
    eh_total = bloco["Classe principal"] == TOTAL
    total = bloco[eh_total]
    bloco = bloco[~eh_total].copy()
    bloco["Mês de Referência"] = pd.to_datetime(bloco["Mês de Referência"], format="%m.%Y", errors="coerce")
    bloco["Data de Faturamento"] = pd.to_datetime(bloco["Data de Faturamento"], format="%d.%m.%Y",
                                                  errors="coerce")
    for c in DIMENSOES.values():
        if c not in ("Mês de Referência", "Data de Faturamento"):
            bloco[c] = bloco[c].fillna("")
    return bloco, total, invalidos


def converte(origem, destino):
    """Csv (or the csv inside a zip) -> parquet. Returns (ok, message); writes only if the check passes."""
    if origem.suffix.lower() == ".zip":
        z = zipfile.ZipFile(origem)
        arquivo = z.open(z.namelist()[0])
    else:
        z, arquivo = None, open(origem, "rb")
    temporario = destino.with_suffix(".parquet.tmp")
    escritor, soma, total, linhas, invalidos = None, dict.fromkeys(MEDIDAS.values(), 0.0), None, 0, 0
    try:
        leitor = pd.read_csv(arquivo, sep=";", encoding="utf-8", skiprows=LINHAS_CABECALHO, header=None,
                             usecols=list(DIMENSOES) + sorted(set(MEDIDAS) | {d for ds in DETALHE.values() for d in ds}),
                             dtype=str, chunksize=TAMANHO_BLOCO)
        for bruto in leitor:
            bloco, t, inv = limpa(bruto)
            invalidos += inv
            if len(t):
                total = t.iloc[0]
            for c in MEDIDAS.values():
                soma[c] += bloco[c].sum()
            linhas += len(bloco)
            tabela = pa.Table.from_pandas(bloco[list(DIMENSOES.values()) + list(MEDIDAS.values())],
                                          preserve_index=False)
            if escritor is None:
                escritor = pq.ParquetWriter(temporario, tabela.schema, compression="zstd")
            escritor.write_table(tabela)
    finally:
        if escritor is not None:
            escritor.close()
        arquivo.close()
        if z is not None:
            z.close()

    if invalidos:
        temporario.unlink(missing_ok=True)
        return False, f"{invalidos} células de medida não numéricas"
    if total is None:
        temporario.unlink(missing_ok=True)
        return False, "linha de total (Result.global) não encontrada"
    alvo = {c: float(total[c]) for c in MEDIDAS.values()}
    difs = {c: soma[c] - alvo[c] for c in MEDIDAS.values()}
    # R$ must match to the cent; quantities are printed rounded line by line in the BW export, so
    # their sum may differ from the total by some units (tolerance: 1 part in 10 thousand)
    ruins = {c: d for c, d in difs.items()
             if abs(d) > (0.5 if c == "Montante (R$)" else max(0.5, 1e-4 * abs(alvo[c])))}
    if ruins:
        temporario.unlink(missing_ok=True)
        return False, "total diferente do Result.global: " + "; ".join(f"{c} {d:,.2f}" for c, d in ruins.items())
    temporario.replace(destino)
    return True, f"{linhas:,} linhas | montante R$ {soma['Montante (R$)']:,.2f} (confere com o Result.global)"


def compacta(csv):
    """Csv -> zip with the same name in the same folder, then delete the csv."""
    zip_final = csv.with_suffix(".zip")
    temporario = csv.with_suffix(".zip.tmp")
    with zipfile.ZipFile(temporario, "w", compression=zipfile.ZIP_DEFLATED, compresslevel=9) as z:
        z.write(csv, arcname=csv.name)
    with zipfile.ZipFile(temporario) as z:                         # the zip must be readable before deleting
        if z.testzip() is not None or z.getinfo(csv.name).file_size != csv.stat().st_size:
            temporario.unlink(missing_ok=True)
            raise RuntimeError(f"zip de {csv.name} não conferiu; o csv foi mantido")
    temporario.replace(zip_final)
    csv.unlink()
    return zip_final


def pendentes(forcar):
    """Csv files always; zip files only when their parquet is missing (or --forcar)."""
    fila = {}
    for arq in sorted(DIR_CSV.glob("*.zip")):
        if forcar or not (DIR_PARQUET / f"{arq.stem}.parquet").exists():
            fila[arq.stem] = arq
    for arq in sorted(DIR_CSV.glob("*.csv")):
        fila[arq.stem] = arq                                       # a new csv wins over its old zip
    return fila


if __name__ == "__main__":
    forcar = "--forcar" in sys.argv[1:]
    DIR_PARQUET.mkdir(parents=True, exist_ok=True)
    fila = pendentes(forcar)
    if not fila:
        print("nada a converter: todos os meses de src/csv_original já têm parquet")
    for nome, origem in fila.items():
        ok, msg = converte(origem, DIR_PARQUET / f"{nome}.parquet")
        print(f"{nome}: {'OK' if ok else 'FALHOU'} - {msg}")
        if ok and origem.suffix.lower() == ".csv":
            print(f"   csv compactado em {compacta(origem).name} e apagado")
