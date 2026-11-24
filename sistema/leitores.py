"""Leitores dos 3 relatórios do ERP (1174, 1211 e Vendas Perdidas).

Cada função devolve um DataFrame já limpo e padronizado:
  - filial: int (1..13)
  - produto: str (código sem espaços, maiúsculo)
"""
from __future__ import annotations

import re
from pathlib import Path

import openpyxl
import pandas as pd


def _cod(v) -> str:
    if v is None:
        return ""
    if isinstance(v, float) and v.is_integer():
        v = int(v)
    return str(v).strip().upper()


def _num_br(s: str) -> float:
    s = (s or "").strip()
    if not s:
        return 0.0
    return float(s.replace(".", "").replace(",", "."))


# ---------------------------------------------------------------- 1174 vendas
COLS_1174 = {
    "Filial": "filial_txt",
    "Grupo": "grupo",
    "Subgrupo": "subgrupo",
    "Marca": "marca",
    "Produto": "produto",
    "Descrição": "descricao",
    "Cód. Int. NF ": "cod_nf",
    "Local de venda": "local_venda",
    "Operação": "operacao",
    "Tipo de Lançamento": "tipo_lanc",
    "Data": "data",
    "Quantidade": "qtd",
    "Valor líquido": "valor",
    "Custo Médio Venda": "custo",
}


def ler_1174(caminho: Path, extras: dict | None = None) -> pd.DataFrame:
    """Relatório 1174 (vendas item a item). Arquivo grande: leitura em streaming.

    extras: colunas adicionais {nome no relatório: nome novo} (ex.: {"Nome": "cliente"}),
    devolvidas como vieram do Excel.
    """
    wb = openpyxl.load_workbook(caminho, read_only=True, data_only=True)
    ws = wb.worksheets[0]
    it = ws.iter_rows(values_only=True)
    header = [str(h) if h is not None else "" for h in next(it)]
    idx = {}
    for orig, novo in {**COLS_1174, **(extras or {})}.items():
        if orig not in header:
            # tolera variação de espaços no cabeçalho
            cand = [i for i, h in enumerate(header) if h.strip() == orig.strip()]
            if not cand:
                raise ValueError(f"1174: coluna '{orig}' não encontrada")
            idx[novo] = cand[0]
        else:
            idx[novo] = header.index(orig)
    cols = list(idx.keys())
    pos = [idx[c] for c in cols]
    dados = []
    for r in it:
        if r is None or r[idx["produto"]] is None:
            continue
        dados.append([r[p] for p in pos])
    wb.close()
    df = pd.DataFrame(dados, columns=cols)
    df["filial"] = df["filial_txt"].astype(str).str.extract(r"^(\d+)")[0].astype(int)
    df["produto"] = df["produto"].map(_cod)
    df["data"] = pd.to_datetime(df["data"]).dt.normalize()
    for c in ("qtd", "valor", "custo"):
        df[c] = pd.to_numeric(df[c], errors="coerce").fillna(0.0)
    for c in ("grupo", "subgrupo", "cod_nf"):
        df[c] = pd.to_numeric(df[c], errors="coerce").fillna(0).astype("int64")
    for c in ("filial_txt", "marca", "descricao", "local_venda", "operacao", "tipo_lanc"):
        df[c] = df[c].astype(str).str.strip().astype("category")
    return df


# ---------------------------------------------------------------- 1211 estoque
_RE_COD = re.compile(r'=\("(.*)"\)')


def ler_1211(caminho: Path) -> pd.DataFrame:
    """Relatório 1211 (estoque). Hierarquia: FILIAL -> GRUPO -> PRODUTO.

    Linhas de grupo/filial têm quantidade com 2 casas decimais; produtos com 4.
    """
    linhas = []
    filial = None
    grupo = None
    grupo_desc = None
    with open(caminho, encoding="utf-8-sig") as f:
        next(f)  # cabeçalho
        for ln in f:
            p = ln.rstrip("\r\n").split(";")
            if len(p) < 6:
                continue
            m = _RE_COD.search(p[1])
            if not m:
                continue  # TOTAL GERAL
            cod = m.group(1).strip()
            desc = p[2].strip()
            qtd_txt = p[3].strip()
            if desc.startswith("FILIAL"):
                filial = int(cod) % 1000 - 100  # "000101" -> 1 (mesmo nº do 1174)
                grupo = None
                continue
            casas = len(qtd_txt.split(",")[-1]) if "," in qtd_txt else 0
            if casas == 2:
                grupo = int(cod) if cod.isdigit() else cod
                grupo_desc = desc
                continue
            linhas.append(
                (filial, grupo, grupo_desc, cod.upper(), desc,
                 _num_br(qtd_txt), _num_br(p[4]), _num_br(p[5]))
            )
    return pd.DataFrame(
        linhas,
        columns=["filial", "subgrupo", "subgrupo_desc", "produto", "descricao",
                 "estoque", "custo_unit", "valor_estoque"],
    )


# ---------------------------------------------------------------- vendas perdidas
def ler_vendas_perdidas(caminho: Path) -> pd.DataFrame:
    wb = openpyxl.load_workbook(caminho, read_only=True, data_only=True)
    ws = wb.worksheets[0]
    rows = list(ws.iter_rows(values_only=True))
    wb.close()
    # acha a linha de cabeçalho (começa com "Cód. Venda Perdida")
    h = next(i for i, r in enumerate(rows) if r and r[0] and str(r[0]).startswith("C"))
    header = [str(x).strip() if x else "" for x in rows[h]]
    df = pd.DataFrame([r for r in rows[h + 1:] if r and r[0] is not None], columns=header)
    ren = {
        "Cód. Venda Perdida": "cod_vp", "Filial": "filial", "Data": "data",
        "Nome do vendedor": "vendedor", "Nome do cliente": "cliente",
        "Produto": "produto", "Descrição do produto": "descricao",
        "Quantidade": "qtd", "Preço Unitário": "preco", "Total": "valor",
        "Descrição": "motivo", "Origem": "origem", "Grupo": "subgrupo",
        "Observação": "obs",
    }
    df = df.rename(columns=ren)[list(ren.values())]
    df["cod_vp"] = pd.to_numeric(df["cod_vp"]).astype("int64")
    df["filial"] = pd.to_numeric(df["filial"]).astype(int)
    df["subgrupo"] = pd.to_numeric(df["subgrupo"], errors="coerce").fillna(0).astype(int)
    df["produto"] = df["produto"].map(_cod)
    df["data"] = pd.to_datetime(df["data"]).dt.normalize()
    df["qtd"] = pd.to_numeric(df["qtd"], errors="coerce").fillna(0.0)
    df["valor"] = pd.to_numeric(df["valor"], errors="coerce").fillna(0.0)
    df["obs"] = df["obs"].astype(str).str.strip()
    return df
