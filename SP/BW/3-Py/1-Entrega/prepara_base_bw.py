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
       "Resultado" de cada medida, sem a linha de total. As colunas sao localizadas pelo NOME no
       cabecalho, nunca pela posicao: coluna nova ou retirada pelo BW (ex.: a data de faturamento,
       retirada em 2026) nao quebra a leitura; falta de coluna obrigatoria para com erro;
    3. confere as cinco medidas do parquet contra a linha "Result.global" do csv (montante ao
       centavo; quantidades com tolerancia de 1 parte em 10 mil, pelo arredondamento do BW). Quando o BW
       imprime "*" no Resultado (unidades misturadas), o valor e a soma das colunas kW, kW.h e
       Nao atribuido da medida;
    4. so se a conferencia passar: compacta o csv em src/csv_original/<mesmo nome>.zip e apaga o csv.
       Se a conferencia falhar, o csv fica onde esta e o parquet nao e gravado.
"""
import csv
import io
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

# Dimensions kept in the parquet, by their clean name. A text column repeats the name of its code
# column in the csv header; the second occurrence gets the " TEXTO" suffix. Any other dimension
# column of the export (e.g. "Data de Faturamento") is ignored.
DIMENSOES = [
    "Classe principal", "Categoria tarifa", "Grp.estatísts.mont.", "Grp.estatísts.quant.",
    "Grp.estatísts.quant. TEXTO", "Operação principal", "Operação principal TEXTO", "Suboperação",
    "Suboperação TEXTO", "Flag Cli.Liv./ Cat.", "Classe/Sub-classe", "Conjunto de Contrato",
    "Mês de Referência", "Flag de Refaturamento",
]
APELIDOS = {"Flag de Refaturament": "Flag de Refaturamento"}   # names truncated by the BW export
# Measure label in the "Índices" header row -> clean name. Only its "Resultado" column is kept;
# the three columns before it (kW, kW.h, Não atribuído) are its breakdown.
MEDIDAS_BW = {
    "TUSD Faturado / KWH":       "TUSD Faturado (kWh)",
    "TUSD Faturado / KW":        "TUSD Faturado (kW)",
    "Demanda Faturado / KW":     "Demanda Faturada (kW)",
    "Consumo TE Faturado / MWH": "Consumo TE Faturado (kWh)",
    "Montante R$":               "Montante (R$)",
}
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


def abre(origem):
    """Binary handle of the csv (or of the csv inside a zip) and the zip to close afterwards."""
    if origem.suffix.lower() == ".zip":
        z = zipfile.ZipFile(origem)
        return z.open(z.namelist()[0]), z
    return open(origem, "rb"), None


def le_layout(origem):
    """Locate the columns by name in the header rows. Returns (dimensions {position: clean name},
    measures {position of "Resultado": clean name}, breakdown {position: [3 positions]})."""
    arquivo, z = abre(origem)
    try:
        texto = io.TextIOWrapper(arquivo, encoding="utf-8")
        cab = [next(csv.reader([texto.readline()], delimiter=";")) for _ in range(LINHAS_CABECALHO)]
    finally:
        arquivo.close()
        if z is not None:
            z.close()
    indices, unidades, nomes = cab[1], cab[2], cab[3]
    primeira = min((j for j, v in enumerate(indices) if v.strip() in MEDIDAS_BW), default=None)
    if primeira is None:
        raise ValueError(f"{origem.name}: no measure found in the header ({sorted(MEDIDAS_BW)})")

    dimensoes, vistos = {}, set()
    for j, nome in enumerate(nomes[:primeira]):
        nome = APELIDOS.get(nome.strip(), nome.strip())
        if not nome:
            continue
        limpo = f"{nome} TEXTO" if nome in vistos else nome
        vistos.add(nome)
        if limpo in DIMENSOES:
            dimensoes[j] = limpo
    faltam = [d for d in DIMENSOES if d not in dimensoes.values()]
    if faltam:
        raise ValueError(f"{origem.name}: dimension columns missing in the header: {faltam}")

    medidas, detalhe = {}, {}
    for j in range(primeira, len(indices)):
        rotulo = indices[j].strip()
        if rotulo in MEDIDAS_BW and unidades[j].strip() == "Resultado":
            if [indices[k].strip() for k in (j - 3, j - 2, j - 1)] != [rotulo] * 3:
                raise ValueError(f"{origem.name}: unexpected breakdown columns for {rotulo}")
            medidas[j], detalhe[j] = MEDIDAS_BW[rotulo], [j - 3, j - 2, j - 1]
    faltam = [m for m in MEDIDAS_BW.values() if m not in medidas.values()]
    if faltam:
        raise ValueError(f"{origem.name}: measure columns missing in the header: {faltam}")
    return dimensoes, medidas, detalhe


def limpa(bloco, dimensoes, medidas, detalhe):
    """Raw chunk (text) -> clean frame, plus the total row when present in the chunk."""
    for col, nome in medidas.items():
        estrela = bloco[col].fillna("").astype(str).str.strip() == "*"
        valor = numero_br(bloco[col])
        if estrela.any():
            valor = valor.where(~estrela, sum(numero_br(bloco[d]) for d in detalhe[col]))
        bloco[col] = np.where(estrela, valor.astype(str), bloco[col].fillna("").astype(str))
        bloco[f"__{nome}"] = valor
    bloco = bloco.drop(columns=[d for ds in detalhe.values() for d in ds])
    bloco = bloco.rename(columns={**dimensoes, **medidas})
    invalidos = sum(nao_numericos(bloco[c]) for c in medidas.values())
    for c in medidas.values():
        bloco[c] = bloco.pop(f"__{c}")
    eh_total = bloco["Classe principal"] == TOTAL
    total = bloco[eh_total]
    bloco = bloco[~eh_total].copy()
    bloco["Mês de Referência"] = pd.to_datetime(bloco["Mês de Referência"], format="%m.%Y", errors="coerce")
    for c in DIMENSOES:
        if c != "Mês de Referência":
            bloco[c] = bloco[c].fillna("")
    return bloco, total, invalidos


def converte(origem, destino):
    """Csv (or the csv inside a zip) -> parquet. Returns (ok, message); writes only if the check passes."""
    dimensoes, medidas, detalhe = le_layout(origem)
    arquivo, z = abre(origem)
    temporario = destino.with_suffix(".parquet.tmp")
    escritor, soma, total, linhas, invalidos = None, dict.fromkeys(MEDIDAS_BW.values(), 0.0), None, 0, 0
    try:
        leitor = pd.read_csv(arquivo, sep=";", encoding="utf-8", skiprows=LINHAS_CABECALHO, header=None,
                             usecols=list(dimensoes) + sorted(set(medidas) | {d for ds in detalhe.values() for d in ds}),
                             dtype=str, chunksize=TAMANHO_BLOCO)
        for bruto in leitor:
            bloco, t, inv = limpa(bruto, dimensoes, medidas, detalhe)
            invalidos += inv
            if len(t):
                total = t.iloc[0]
            for c in MEDIDAS_BW.values():
                soma[c] += bloco[c].sum()
            linhas += len(bloco)
            tabela = pa.Table.from_pandas(bloco[DIMENSOES + list(MEDIDAS_BW.values())],
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
    alvo = {c: float(total[c]) for c in MEDIDAS_BW.values()}
    difs = {c: soma[c] - alvo[c] for c in MEDIDAS_BW.values()}
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
