"""Motor de previsão de falta de peças (BETA).

Para cada (filial, produto):
  demanda diária prevista = mistura do ritmo recente (vendas + vendas perdidas)
                            com o mesmo período do ano anterior (sazonalidade)
  cobertura (dias)        = estoque / demanda diária
  status                  = RUPTURA / CRÍTICO / ATENÇÃO / OK / SEM GIRO
"""
from __future__ import annotations

import numpy as np
import pandas as pd

import config as C


def _soma_janela(df, ini, fim, col="qtd"):
    m = (df["data"] > ini) & (df["data"] <= fim)
    return df.loc[m].groupby(["filial", "produto"])[col].sum()


def calcular(vendas: pd.DataFrame, estoque: pd.DataFrame, perdidas: pd.DataFrame,
             data_base: pd.Timestamp | None = None,
             data_estoque: pd.Timestamp | None = None,
             modelo_R: dict | None = None) -> tuple[pd.DataFrame, dict]:
    """modelo_R: resultado de modelo.rodar(). Se informado, a demanda diária vem do modelo."""
    if data_base is None:
        data_base = vendas["data"].max()
    hoje = pd.Timestamp(data_base).normalize()          # fim das janelas de venda
    dia_est = pd.Timestamp(data_estoque or hoje).normalize()  # a cobertura parte daqui
    H = C.LEAD_TIME_DIAS + C.SEGURANCA_DIAS          # horizonte de proteção
    R = C.JANELA_RECENTE_DIAS

    # venda líquida: o 1174 já traz devolução com qtd negativa.
    # Garantia ("Outras saídas") é consumo real de peça e entra como demanda.
    v = vendas[vendas["tipo_lanc"].astype(str).isin(C.TIPOS_DEMANDA)].copy()

    k = ["filial", "produto"]
    d = pd.DataFrame({
        "vendido_30d": _soma_janela(v, hoje - pd.Timedelta(days=30), hoje),
        "vendido_recente": _soma_janela(v, hoje - pd.Timedelta(days=R), hoje),
        "vendido_12m": _soma_janela(v, hoje - pd.Timedelta(days=365), hoje),
        # ano anterior: próximos H dias, a partir de hoje-365
        "vendido_ano_ant_prox": _soma_janela(
            v, hoje - pd.Timedelta(days=365), hoje - pd.Timedelta(days=365 - H)),
        "valor_12m": _soma_janela(v, hoje - pd.Timedelta(days=365), hoje, "valor"),
    })

    v12 = v[(v["data"] > hoje - pd.Timedelta(days=365)) & (v["qtd"] > 0)]
    d["meses_com_venda_12m"] = v12.assign(m=v12["data"].dt.to_period("M")) \
        .groupby(k)["m"].nunique()
    d["ultima_venda"] = v[v["qtd"] > 0].groupby(k)["data"].max()
    # preço médio de venda (12m) para valorizar o risco
    pos = v12.groupby(k)[["qtd", "valor"]].sum()
    d["preco_medio"] = pos["valor"] / pos["qtd"].replace(0, np.nan)

    # vendas perdidas
    p = perdidas[perdidas["data"] > hoje - pd.Timedelta(days=R)]
    pg = p.groupby(k).agg(perdido_qtd=("qtd", "sum"), perdido_valor=("valor", "sum"),
                          perdido_ocorrencias=("qtd", "size"),
                          ultima_perda=("data", "max"))
    d = d.join(pg, how="outer")

    # estoque (o 1211 só traz saldo > 0; ausência = 0)
    e = estoque.set_index(k)[["estoque", "custo_unit"]]
    d = d.join(e, how="outer")

    # descrições / subgrupo: estoque > vendas > perdidas
    desc = pd.concat([
        estoque[k + ["descricao", "subgrupo"]],
        vendas.drop_duplicates(k, keep="last")[k + ["descricao", "subgrupo"]],
        perdidas[k + ["descricao", "subgrupo"]],
    ]).drop_duplicates(k).set_index(k)
    d = d.join(desc, how="left")

    num = ["vendido_30d", "vendido_recente", "vendido_12m", "vendido_ano_ant_prox",
           "valor_12m", "meses_com_venda_12m", "perdido_qtd", "perdido_valor",
           "perdido_ocorrencias", "estoque"]
    d[num] = d[num].fillna(0)
    d.loc[d["estoque"] < 0, "estoque"] = 0
    d[["vendido_recente", "vendido_30d", "vendido_12m", "vendido_ano_ant_prox"]] = \
        d[["vendido_recente", "vendido_30d", "vendido_12m", "vendido_ano_ant_prox"]].clip(lower=0)

    # só interessa o que teve movimento em 12m, venda perdida ou estoque
    d = d[(d["vendido_12m"] > 0) | (d["perdido_qtd"] > 0) | (d["estoque"] > 0)].copy()
    prods = d.index.get_level_values("produto")
    excl = np.zeros(len(d), dtype=bool)
    for pre in C.EXCLUIR_PREFIXOS:
        excl |= prods.str.startswith(pre)
    d = d[~excl]

    # ---- demanda prevista (unid/dia)
    dem_recente = (d["vendido_recente"] + C.PESO_VENDA_PERDIDA * d["perdido_qtd"]) / R
    dem_sazonal = d["vendido_ano_ant_prox"] / H
    tem_hist = d["vendido_12m"] > 0
    w = np.where(tem_hist, C.PESO_SAZONAL, 0.0)
    d["demanda_dia"] = (1 - w) * dem_recente + w * dem_sazonal
    if modelo_R is not None:
        import modelo
        dm = modelo.demanda_diaria_horizonte(modelo_R, dia_est, H)
        mi = modelo_R["info"].set_index(["filial", "produto"])
        hit = d.index.isin(dm.index)
        d.loc[hit, "demanda_dia"] = dm.reindex(d.index[hit]).to_numpy()
        d["classe"] = mi["classe"].reindex(d.index).fillna("—").to_numpy()
        d["metodo"] = mi["metodo"].reindex(d.index).fillna("—").to_numpy()
        d["acur_item"] = mi["acur_item"].reindex(d.index).to_numpy()
        d["codigos_antigos"] = mi["codigos_antigos"].reindex(d.index).fillna("").to_numpy()
        # código antigo já substituído: a demanda está no código novo
        antigos = set(modelo_R["equivalencias"]["codigo_antigo"])
        sub = d.index.get_level_values("produto").isin(antigos)
        d.loc[sub, "demanda_dia"] = 0.0
        d.loc[sub, "classe"] = "SUBSTITUÍDO"
    d["demanda_mes"] = d["demanda_dia"] * 30
    d["demanda_horizonte"] = d["demanda_dia"] * H

    d["cobertura_dias"] = np.where(d["demanda_dia"] > 0,
                                   d["estoque"] / d["demanda_dia"].replace(0, np.nan),
                                   np.inf)
    d["data_prevista_falta"] = np.where(
        np.isfinite(d["cobertura_dias"]),
        dia_est + pd.to_timedelta(np.minimum(d["cobertura_dias"].replace(np.inf, 0), 3650), unit="D"),
        pd.NaT)
    d["data_prevista_falta"] = pd.to_datetime(d["data_prevista_falta"]).dt.normalize()

    # sugestão de compra: cobrir LT + segurança + ciclo
    alvo = d["demanda_dia"] * (H + C.CICLO_COMPRA_DIAS)
    d["sugestao_compra"] = np.ceil((alvo - d["estoque"]).clip(lower=0))
    d.loc[d["demanda_dia"] <= 0, "sugestao_compra"] = 0

    # estoque das outras filiais (possível transferência)
    rede = d.groupby(level="produto")["estoque"].sum()
    d["estoque_outras_filiais"] = d.index.get_level_values("produto").map(rede).values - d["estoque"]
    # filiais com sobra: estoque acima da própria necessidade no horizonte
    d["sobra"] = (d["estoque"] - np.ceil(d["demanda_dia"] * H)).clip(lower=0)
    sobra = d[d["sobra"] > 0].reset_index()
    sobra_map: dict[str, list] = {}
    for f, prod, s in zip(sobra["filial"], sobra["produto"], sobra["sobra"]):
        sobra_map.setdefault(prod, []).append((f, s))

    # ---- status
    recorrente = (d["meses_com_venda_12m"] >= C.MIN_MESES_RECORRENCIA) | (d["perdido_ocorrencias"] >= 2)
    # demanda relevante: espera-se vender ao menos MIN_UNID_HORIZONTE no horizonte, ou houve venda perdida
    relevante = (d["demanda_dia"] * H >= C.MIN_UNID_HORIZONTE) | (d["perdido_qtd"] > 0)
    st = np.select(
        [
            (d["demanda_dia"] * H < C.MIN_UNID_HORIZONTE) & (d["estoque"] > 0) & (d["perdido_qtd"] <= 0),
            (d["estoque"] <= 0) & relevante,
            (d["estoque"] <= 0),
            d["cobertura_dias"] < C.LEAD_TIME_DIAS,
            d["cobertura_dias"] < H,
        ],
        ["SEM GIRO", "RUPTURA", "OK", "CRÍTICO", "ATENÇÃO"],
        default="OK",
    )
    d["status"] = st
    d.loc[~relevante, "sugestao_compra"] = 0
    if "classe" in d:
        d.loc[d["classe"] == "SUBSTITUÍDO", "status"] = "SUBSTITUÍDO"
    d["recorrencia"] = np.where(recorrente, "RECORRENTE", "BAIXA RECORRÊNCIA")
    d["dias_sem_venda"] = (dia_est - d["ultima_venda"]).dt.days
    d["evidencia"] = np.select(
        [d["perdido_qtd"] > 0, d["vendido_recente"] > 0, d["vendido_ano_ant_prox"] > 0],
        ["VENDA PERDIDA", "VENDA RECENTE", "SÓ SAZONAL (ano ant.)"], default="")

    preco = d["preco_medio"].fillna(d["custo_unit"])
    falta_qtd = (d["demanda_horizonte"] - d["estoque"]).clip(lower=0)
    d["valor_em_risco"] = (falta_qtd * preco).fillna(0)

    d = d.reset_index()
    d["filial_nome"] = d["filial"].map(lambda f: f"{f:02d} {C.NOMES_FILIAIS.get(f, '')}")
    def _outras(f, prod):
        lst = [(ff, s) for ff, s in sobra_map.get(prod, []) if ff != f]
        txt = ", ".join(f"{C.NOMES_FILIAIS.get(ff, ff)}({int(s)})"
                        for ff, s in sorted(lst, key=lambda x: -x[1]))
        return txt, sum(s for _, s in lst)
    precisa = d["status"].isin(["RUPTURA", "CRÍTICO", "ATENÇÃO"])
    res = [(_outras(f, p) if q else ("", 0.0))
           for f, p, q in zip(d["filial"], d["produto"], precisa)]
    d["sobra_outras_filiais"] = [r[0] for r in res]
    d["sobra_rede_qtd"] = [r[1] for r in res]

    prioridade = {"RUPTURA": 0, "CRÍTICO": 1, "ATENÇÃO": 2, "OK": 3, "SEM GIRO": 4, "SUBSTITUÍDO": 5}
    d["_p"] = d["status"].map(prioridade)
    d["_r"] = (d["recorrencia"] != "RECORRENTE").astype(int)
    d = d.sort_values(["_p", "_r", "valor_em_risco"], ascending=[True, True, False]) \
         .drop(columns=["_p", "_r"])

    meta = dict(data_base=hoje, data_estoque_ts=dia_est, horizonte=H, janela=R)
    return d, meta
