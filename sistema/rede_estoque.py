"""Cálculos da rede inteira para a Consulta de Peças: demanda por filial × produto no
período-base, sugestões de transferência entre lojas e saúde do estoque por loja.

Regras (combinadas com o usuário):
- Demanda = venda líquida (Vendas + Devolução de vendas) + vendas perdidas no período-base;
  média mensal = soma (com cada mês limitado a >= 0) ÷ nº de meses do período — a mesma regra
  da aba Previsão no modo "Período do 1174".
- Transferência: a ORIGEM guarda a própria previsão de (Atender + 1) meses; só o que passar
  disso pode ir para outra loja. O DESTINO recebe até cobrir a previsão de Atender meses.
- Saúde: Ruptura (sem estoque, com demanda) · Abaixo (estoque < previsão de Atender) ·
  Em giro · Excesso (cobertura acima do limite escolhido) · Parada (estoque e nenhuma demanda).
- "Recorrência": falta (ruptura / abaixo / destino de transferência) só conta para a loja que
  teve demanda em pelo menos N meses do período — evita sugerir peça vendida uma única vez.
"""
from __future__ import annotations

import re

import numpy as np
import pandas as pd

STATUS = ["Ruptura", "Abaixo", "Em giro", "Excesso", "Parada"]
# venda perdida acima de 10× a maior venda (uma linha do 1174) da peça é descartada; teto mínimo:
LIMITE_PERDIDA_MIN = 50.0
_LV = ["filial", "produto"]


def _cod_sub(x) -> str:
    s = "" if x is None or (isinstance(x, float) and pd.isna(x)) else str(x).strip()
    return re.sub(r"\.0$", "", s) if s not in ("nan", "None") else ""


def _vazio(cols):
    idx = pd.MultiIndex.from_arrays([[], []], names=_LV)
    return pd.DataFrame({c: pd.Series(dtype=float) for c in cols}, index=idx)


def _info(base):
    """Informações por produto que não dependem do período (calculadas 1 vez por carga)."""
    info = getattr(base, "_rede_info", None)
    if info is not None:
        return info
    v, e = base.vendas, base.estoque
    ref = v["data"].max() if not v.empty else pd.Timestamp.today()
    if not v.empty:
        vp = v[(v["tipo_lanc"].astype(str) == "Vendas") & (v["qtd"] > 0)]
        v12 = vp[vp["data"] > ref - pd.Timedelta(days=365)]

        def unit(df):
            g = df.groupby("produto")[["qtd", "custo"]].sum()
            return (g["custo"] / g["qtd"]).where((g["qtd"] > 0) & (g["custo"] > 0)).dropna()
        custo = unit(v12).combine_first(unit(vp))
        fp = pd.DataFrame({"cmv12": v12.groupby(_LV)["custo"].sum(),
                           "ult_venda": vp.groupby(_LV)["data"].max()})
        sub_v = v.drop_duplicates("produto", keep="last").set_index("produto")["subgrupo"].map(_cod_sub)
        maior = vp.groupby("produto")["qtd"].max()
    else:
        custo, fp, sub_v = pd.Series(dtype=float), _vazio(["cmv12", "ult_venda"]), pd.Series(dtype=str)
        maior = pd.Series(dtype=float)
    if not e.empty:
        ce = e[e["custo_unit"] > 0].groupby("produto")["custo_unit"].mean()
        custo = custo.combine_first(ce)
        ee = e.drop_duplicates("produto", keep="last").set_index("produto")
        sub_e = ee["subgrupo"].map(_cod_sub)
        d = e.drop_duplicates("subgrupo")
        nomes = dict(zip(d["subgrupo"].map(_cod_sub), d["subgrupo_desc"].astype(str).str.strip()))
    else:
        sub_e, nomes = pd.Series(dtype=str), {}
    sub = sub_e[sub_e != ""].combine_first(sub_v)
    info = dict(custo=custo, fp=fp, sub=sub, nomes_sub=nomes, ref=ref, maior_venda=maior)
    base._rede_info = info
    return info


def lista_subgrupos(base) -> list[tuple[str, str]]:
    """[(código, descrição)] de todos os subgrupos que aparecem no estoque ou nas vendas."""
    info = _info(base)
    cods = sorted({c for c in info["sub"].unique() if c}, key=lambda c: (len(c), c))
    return [(c, info["nomes_sub"].get(c, "")) for c in cods]


def painel_rede(base, de: pd.Period, ate: pd.Period) -> pd.DataFrame:
    """Uma linha por filial × produto (com estoque ou com demanda no período-base)."""
    chave = (str(de), str(ate))
    cache = base.__dict__.setdefault("_painel_cache", {})
    if chave in cache:
        return cache[chave]
    info = _info(base)
    v, p, e = base.vendas, base.perdidas, base.estoque
    ini, fim = de.start_time, ate.end_time
    n_meses = (ate.year - de.year) * 12 + ate.month - de.month + 1

    partes, pq = [], _vazio(["perd_qtd", "perd_valor"])
    if not v.empty:
        m = v["tipo_lanc"].astype(str).isin(["Vendas", "Devolução de vendas"]) & \
            (v["data"] >= ini) & (v["data"] <= fim)
        partes.append(v.loc[m, ["filial", "produto", "data", "qtd"]])
    if not p.empty:
        pp = p.loc[(p["data"] >= ini) & (p["data"] <= fim), ["filial", "produto", "data", "qtd", "valor"]].copy()
        pp["qtd"] = pd.to_numeric(pp["qtd"], errors="coerce").fillna(0.0)
        pp["valor"] = pd.to_numeric(pp["valor"], errors="coerce").fillna(0.0)
        # descarta venda perdida absurda (erro de digitação, ex.: 825.270 un de uma correia)
        teto = pp["produto"].map(info["maior_venda"]).fillna(0.0).mul(10).clip(lower=LIMITE_PERDIDA_MIN)
        pp = pp[pp["qtd"] <= teto]
        partes.append(pp[["filial", "produto", "data", "qtd"]])
        pq = pp.groupby(_LV)[["qtd", "valor"]].sum().rename(columns={"qtd": "perd_qtd", "valor": "perd_valor"})
    if partes:
        d = pd.concat(partes, ignore_index=True)
        d["qtd"] = pd.to_numeric(d["qtd"], errors="coerce").fillna(0.0)
        d["m"] = d["data"].dt.year * 12 + d["data"].dt.month
        mm = d.groupby(_LV + ["m"])["qtd"].sum().clip(lower=0)
        dem = pd.DataFrame({"soma": mm.groupby(level=_LV).sum(),
                            "meses_dem": (mm > 0).groupby(level=_LV).sum()})
    else:
        dem = _vazio(["soma", "meses_dem"])
    if not e.empty:
        eg = e.groupby(_LV).agg(est=("estoque", "sum"), valor_est=("valor_estoque", "sum"),
                                custo_e=("custo_unit", "max"))
    else:
        eg = _vazio(["est", "valor_est", "custo_e"])

    pn = eg.join(dem, how="outer").join(pq, how="outer").join(info["fp"], how="left").reset_index()
    for c in ["est", "valor_est", "custo_e", "soma", "meses_dem", "perd_qtd", "perd_valor", "cmv12"]:
        pn[c] = pd.to_numeric(pn[c], errors="coerce").fillna(0.0)
    pn["filial"] = pn["filial"].astype(int)
    pn["media"] = pn["soma"] / n_meses
    pn = pn[(pn["est"] > 0) | (pn["media"] > 0)].reset_index(drop=True)
    custo = pn["custo_e"].where(pn["custo_e"] > 0, pn["produto"].map(info["custo"]))
    pn["custo"] = pd.to_numeric(custo, errors="coerce").fillna(0.0)
    pn["valor_est"] = pn["valor_est"].where(pn["valor_est"] > 0, pn["est"].clip(lower=0) * pn["custo"])
    pn["subgrupo"] = pn["produto"].map(info["sub"]).fillna("")
    pn["descricao"] = pn["produto"].map(base.descricoes).fillna("")
    pn.attrs["n_meses"] = n_meses
    cache[chave] = pn
    return pn


def _recorr(pn, recorr):
    return max(1, min(int(recorr), int(pn.attrs.get("n_meses", recorr))))


def transferencias(pn: pd.DataFrame, n: int, recorr: int = 2) -> pd.DataFrame:
    """Sugestões origem → destino, peça a peça (casamento guloso)."""
    cols = ["produto", "descricao", "subgrupo", "de", "est_de", "media_de", "para", "est_para",
            "media_para", "qtd", "custo", "valor", "motivo"]
    if pn.empty:
        return pd.DataFrame(columns=cols)
    est = pn["est"].clip(lower=0)
    media = pn["media"]
    sobra = (est - np.ceil(media * (n + 1) - 1e-9)).clip(lower=0)
    destino_ok = (media > 0) & (pn["meses_dem"] >= _recorr(pn, recorr))
    falta = np.where(destino_ok, np.ceil(media * n - 1e-9) - est, 0).clip(min=0)
    tab = pn.assign(sobra=sobra, falta=falta, est0=est)
    prods = set(tab.loc[tab["sobra"] > 0, "produto"]) & set(tab.loc[tab["falta"] > 0, "produto"])
    if not prods:
        return pd.DataFrame(columns=cols)
    tab = tab[tab["produto"].isin(prods) & ((tab["sobra"] > 0) | (tab["falta"] > 0))]
    # ordem: por produto; origens (parada primeiro, depois maior sobra) e destinos (maior falta)
    tab = tab.assign(papel=np.where(tab["sobra"] > 0, 0, 1), ativo=(tab["media"] > 0).astype(int),
                     ordem=-np.maximum(tab["sobra"], tab["falta"]))
    tab = tab.sort_values(["produto", "papel", "ativo", "ordem"], kind="stable")
    A = {c: tab[c].to_numpy() for c in ["produto", "descricao", "subgrupo", "filial", "est0", "media",
                                         "custo", "sobra", "falta", "papel"]}
    inicio = np.flatnonzero(np.r_[True, A["produto"][1:] != A["produto"][:-1]])
    fim_g = np.r_[inicio[1:], len(tab)]
    linhas = []
    for a, z in zip(inicio, fim_g):
        idx = range(a, z)
        origens = [[i, float(A["sobra"][i])] for i in idx if A["papel"][i] == 0]
        for j in idx:
            if A["papel"][j] != 1:
                continue
            precisa = float(A["falta"][j])
            for o in origens:
                if precisa <= 0:
                    break
                if o[1] <= 0:
                    continue
                q = min(precisa, o[1])
                o[1] -= q
                precisa -= q
                i = o[0]
                mo, eo, md, ed = A["media"][i], A["est0"][i], A["media"][j], A["est0"][j]
                m_o = "origem sem venda no período" if mo <= 0 else f"origem cobre {eo / mo:.1f} meses"
                m_d = "destino sem estoque" if ed <= 0 else f"destino cobre {ed / md:.1f} meses"
                linhas.append((A["produto"][j], A["descricao"][j], A["subgrupo"][j], int(A["filial"][i]),
                               float(eo), float(mo), int(A["filial"][j]), float(ed), float(md), q,
                               float(A["custo"][i]), q * float(A["custo"][i]), f"{m_o} · {m_d}"))
    out = pd.DataFrame(linhas, columns=cols)
    return out.sort_values("valor", ascending=False, kind="stable").reset_index(drop=True)


def classificar(pn: pd.DataFrame, n: int, limite: float, recorr: int = 2) -> pd.DataFrame:
    """Situação de cada filial × produto + valores de excesso e parado."""
    if pn.empty:
        return pn.assign(status=pd.Series(dtype=str), cob=pd.Series(dtype=float))
    est = pn["est"].clip(lower=0)
    media = pn["media"]
    nec = np.ceil(media * n - 1e-9)
    recorrente = pn["meses_dem"] >= _recorr(pn, recorr)
    cob = np.where(media > 0, est / media.where(media > 0, 1.0), np.inf)
    st = np.select([(est <= 0) & recorrente, est <= 0, media <= 0, (est < nec) & recorrente, cob > limite],
                   ["Ruptura", "", "Parada", "Abaixo", "Excesso"], "Em giro")
    out = pn.assign(status=st, cob=cob, nec=nec, est0=est)
    out = out[out["status"] != ""].copy()
    cu = np.where(out["est0"] > 0, out["valor_est"] / out["est0"].where(out["est0"] > 0, 1.0), out["custo"])
    out["excesso_valor"] = np.where(out["status"] == "Excesso",
                                    (out["est0"] - limite * out["media"]).clip(lower=0) * cu, 0.0)
    out["parado_valor"] = np.where(out["status"] == "Parada", out["valor_est"], 0.0)
    out["dem_valor"] = out["media"] * out["custo"]
    return out


def resumo_lojas(cl: pd.DataFrame, transf: pd.DataFrame) -> pd.DataFrame:
    """Uma linha por loja + 'total' (índice -1)."""
    t = cl.assign(saldo=cl["est0"] > 0, **{f"n_{s}": cl["status"] == s for s in STATUS})
    soma = ["valor_est", "saldo", "dem_valor", "cmv12", "parado_valor", "excesso_valor",
            "perd_qtd", "perd_valor"] + [f"n_{s}" for s in STATUS]
    g = t.groupby("filial")[soma].sum()
    env = transf.groupby("de")["valor"].sum() if not transf.empty else pd.Series(dtype=float)
    g["transf"] = env.reindex(g.index).fillna(0.0)
    g.loc[-1] = g.sum()
    g["cob"] = g["valor_est"] / g["dem_valor"].where(g["dem_valor"] > 0)
    g["giro"] = g["cmv12"] / g["valor_est"].where(g["valor_est"] > 0)
    g["em_giro_valor"] = (g["valor_est"] - g["parado_valor"] - g["excesso_valor"]).clip(lower=0)
    return g
