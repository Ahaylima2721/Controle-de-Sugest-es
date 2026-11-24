"""Modelo de previsão de vendas (demanda mensal por filial × produto).

Como funciona
-------------
1. Demanda mensal = venda líquida (vendas − devoluções + garantia) + vendas perdidas.
   Códigos substituídos (ex.: 81331R61BR -> 81331R61BRDSH) têm o histórico unido.
2. Vários métodos competem num BACKTEST: simulamos a previsão feita em cada um dos
   últimos 9 meses e comparamos com o que realmente aconteceu nos 2 meses seguintes.
   Cada método roda em duas versões: direto na filial, ou na rede inteira e depois
   distribuído entre as filiais pela participação de cada uma ("TD_", top-down).
3. Para cada perfil de demanda (SUAVE, ERRÁTICO, INTERMITENTE, IRREGULAR, RARO,
   INATIVO) fica o método com menor ERRO entre os que têm viés de no máximo ±30%.
   Limitar o viés evita um modelo que "acerta" prevendo sempre de menos (o que
   causaria ruptura). Picos isolados do histórico são cortados antes de prever.
4. Reconciliação: as previsões das filiais são escaladas para que o total de cada
   produto na rede bata com a previsão feita direto na rede (RECONCILIAR_COM).
   A filial define a DISTRIBUIÇÃO; a rede define o TOTAL.
"""
from __future__ import annotations

import time
from pathlib import Path

import numpy as np
import pandas as pd

import config as C

# ----------------------------------------------------------------- parâmetros
N_ORIGENS = 9          # meses passados simulados no backtest
H_AVAL = 2             # horizonte avaliado: soma dos próximos 2 meses (≈ lead time + segurança)
H_FINAL = 4            # meses previstos à frente
SAZ_SHRINK = 0.7       # 1 = índice sazonal puro; 0 = ignora sazonalidade
SAZ_MIN_VOLUME = 300   # volume mínimo (24m) para o subgrupo ter índice próprio
OUTLIER_K = 2.0        # mês acima de K × mediana dos meses com venda é "pico" e é cortado
VIES_MAX = 0.30        # na escolha do método, só valem os que erram para mais/menos até 30% no total
# Reconciliação: depois de prever cada filial, escala as filiais para que o total do produto
# na rede bata com a previsão feita direto na rede por este método (mais estável no agregado).
# Validado em 17/11/2026: melhora item×filial e iguala o BETA nos níveis rede e subgrupo,
# inclusive nos 3 meses que não participaram da escolha. None = desliga.
RECONCILIAR_COM = "TD_BETA_V1"
SUFIXOS_SUBST = ["DSH", "OF", "BR", "US", "FP", "PA", "OL", "ZF", "A"]
DIAS_ANTIGO_PARADO = 90    # código antigo sem venda há pelo menos isso
DIAS_TRANSICAO = 180       # antigo parou até X dias depois de o novo começar

BASE_METODOS = {
    "MM3": "Média dos últimos 3 meses",
    "MM6": "Média dos últimos 6 meses",
    "MM12": "Média dos últimos 12 meses",
    "SNAIVE": "Mesmo mês do ano anterior",
    "SBA10": "Croston/SBA α=0,10 (demanda intermitente)",
    "SBA20": "Croston/SBA α=0,20 (demanda intermitente)",
    "SAZ12": "Nível 12 meses × sazonalidade do subgrupo",
    "SAZ6": "Nível 6 meses (dessazonalizado) × sazonalidade do subgrupo",
    "COMB": "Média de SAZ12, SBA10 e MM6",
    "BETA_V1": "Método da 1ª versão BETA (50% últimos 3m + 50% ano anterior)",
}
METODOS = dict(BASE_METODOS)
METODOS.update({f"TD_{k}": f"{v} — calculado na rede e distribuído por filial"
                for k, v in BASE_METODOS.items()})

CLASSES = ["SUAVE", "ERRÁTICO", "INTERMITENTE", "IRREGULAR", "RARO", "INATIVO"]
CLASSES_GIRO = ["SUAVE", "ERRÁTICO", "INTERMITENTE", "IRREGULAR"]
CLASSES_DESC = {
    "SUAVE": "Vende quase todo mês, quantidades estáveis",
    "ERRÁTICO": "Vende quase todo mês, quantidades muito variáveis",
    "INTERMITENTE": "Tem meses sem venda; quantidade estável quando vende",
    "IRREGULAR": "Tem meses sem venda e quantidades variáveis",
    "RARO": "Vendeu em só 1 ou 2 meses dos últimos 12",
    "INATIVO": "Sem venda nos últimos 12 meses",
}


# ----------------------------------------------------------------- equivalências
def detectar_equivalencias(vendas: pd.DataFrame, arq_manual: Path | None = None) -> pd.DataFrame:
    v = vendas[vendas["tipo_lanc"].astype(str) == "Vendas"]
    fim = v["data"].max()
    fl = v.groupby("produto")["data"].agg(["min", "max"])
    qtd = v.groupby("produto")["qtd"].sum()
    prods = set(fl.index)
    cand = []
    for suf in SUFIXOS_SUBST:
        for p in prods:
            if not p.endswith(suf) or len(p) <= len(suf) + 3:
                continue
            b = p[: -len(suf)]
            if b not in prods:
                continue
            bmin, bmax = fl.at[b, "min"], fl.at[b, "max"]
            pmin = fl.at[p, "min"]
            if (bmax < fim - pd.Timedelta(days=DIAS_ANTIGO_PARADO) and pmin > bmin
                    and bmax <= pmin + pd.Timedelta(days=DIAS_TRANSICAO)):
                cand.append((b, p, qtd.get(p, 0)))
    eq = pd.DataFrame(cand, columns=["codigo_antigo", "codigo_novo", "qtd_novo"])
    # um antigo com vários candidatos -> fica o novo que mais vende
    eq = eq.sort_values("qtd_novo", ascending=False).drop_duplicates("codigo_antigo")
    eq = eq.assign(origem="AUTOMÁTICO")[["codigo_antigo", "codigo_novo", "origem"]]
    if arq_manual and Path(arq_manual).exists():
        man = pd.read_csv(arq_manual, sep=";", comment="#", dtype=str).dropna(how="all")
        if len(man):
            man.columns = [c.strip().lower() for c in man.columns]
            man = man.apply(lambda s: s.str.strip().str.upper())
            bloq = set(zip(man.loc[man["acao"] == "NAO_UNIR", "codigo_antigo"],
                           man.loc[man["acao"] == "NAO_UNIR", "codigo_novo"]))
            eq = eq[[(a, b) not in bloq for a, b in zip(eq["codigo_antigo"], eq["codigo_novo"])]]
            uni = man[man["acao"] == "UNIR"][["codigo_antigo", "codigo_novo"]].assign(origem="MANUAL")
            eq = pd.concat([eq[~eq["codigo_antigo"].isin(uni["codigo_antigo"])], uni], ignore_index=True)
    return eq.reset_index(drop=True)


def _resolver(eq: pd.DataFrame) -> dict:
    m = dict(zip(eq["codigo_antigo"], eq["codigo_novo"]))
    out = {}
    for a in m:
        b, vistos = m[a], {a}
        while b in m and b not in vistos:
            vistos.add(b); b = m[b]
        out[a] = b
    return out


# ----------------------------------------------------------------- séries
def montar_series(vendas, perdidas, estoque=None, equiv: dict | None = None):
    equiv = equiv or {}
    v = vendas[vendas["tipo_lanc"].astype(str).isin(C.TIPOS_DEMANDA)].copy()
    v["produto"] = v["produto"].astype(str)
    if equiv:
        v["produto"] = v["produto"].map(lambda p: equiv.get(p, p))
        perdidas = perdidas.assign(produto=perdidas["produto"].map(lambda p: equiv.get(p, p)))
    fim = v["data"].max()
    mes_fechado = (fim + pd.Timedelta(days=1)).month != fim.month
    ult_completo = fim.to_period("M") if mes_fechado else fim.to_period("M") - 1
    meses = pd.period_range(v["data"].min().to_period("M"), ult_completo, freq="M")
    T = len(meses)
    v["m"] = v["data"].dt.to_period("M")
    k = ["filial", "produto"]

    pares = pd.concat([v[k], perdidas[k]] + ([estoque[k]] if estoque is not None else [])) \
        .drop_duplicates().reset_index(drop=True)
    pares["i"] = np.arange(len(pares))
    n = len(pares)
    pos = {p: j for j, p in enumerate(meses)}

    def matriz(df, col="qtd"):
        g = df.groupby(k + ["m"], observed=True)[col].sum().reset_index().merge(pares, on=k)
        g["t"] = g["m"].map(pos)
        g = g.dropna(subset=["t"])
        M = np.zeros((n, T))
        np.add.at(M, (g["i"].to_numpy(), g["t"].astype(int).to_numpy()), g[col].to_numpy())
        return M

    def vetor(df):
        out = np.zeros(n)
        if len(df):
            g = df.groupby(k)["qtd"].sum().reset_index().merge(pares, on=k)
            out[g["i"].to_numpy()] = g["qtd"].to_numpy()
        return out

    pp = perdidas.assign(m=perdidas["data"].dt.to_period("M"))
    Y = np.clip(matriz(v[v["m"] <= ult_completo]), 0, None)
    L = np.clip(matriz(pp[pp["m"] <= ult_completo]), 0, None)
    parcial = np.clip(vetor(v[v["m"] > ult_completo]), 0, None)
    parcial_perdida = vetor(pp[pp["m"] > ult_completo])

    # atributos: subgrupo, descrição, preço médio 12m
    ult = v.sort_values("data").drop_duplicates(k, keep="last")[k + ["subgrupo", "descricao"]]
    ult = ult.assign(descricao=ult["descricao"].astype(str), subgrupo=ult["subgrupo"].astype(float))
    info = pares.merge(ult, on=k, how="left")
    if estoque is not None:
        e = estoque[k + ["subgrupo", "descricao", "custo_unit"]].rename(
            columns={"subgrupo": "sg_e", "descricao": "desc_e"})
        info = info.merge(e, on=k, how="left")
        info["subgrupo"] = info["subgrupo"].fillna(info["sg_e"])
        info["descricao"] = info["descricao"].fillna(info["desc_e"])
    else:
        info["custo_unit"] = np.nan
    pd_ = perdidas.drop_duplicates(k, keep="last")[k + ["subgrupo", "descricao"]] \
        .rename(columns={"subgrupo": "sg_p", "descricao": "desc_p"})
    info = info.merge(pd_, on=k, how="left")
    info["subgrupo"] = info["subgrupo"].fillna(info["sg_p"]).fillna(0).astype(int)
    info["descricao"] = info["descricao"].fillna(info["desc_p"]).fillna("").astype(str)
    v12 = v[(v["data"] > fim - pd.Timedelta(days=365)) & (v["qtd"] > 0)]
    pr = v12.groupby(k)[["qtd", "valor"]].sum()
    pr = (pr["valor"] / pr["qtd"]).rename("preco").reset_index()
    info = info.merge(pr, on=k, how="left")
    # preço do produto na rede quando a filial não vendeu
    prr = v12.groupby("produto")[["qtd", "valor"]].sum()
    info["preco"] = info["preco"].fillna(info["produto"].map(prr["valor"] / prr["qtd"]))
    info["preco"] = info["preco"].fillna(info["custo_unit"]).fillna(1.0).clip(lower=0.01)
    info = info.sort_values("i").reset_index(drop=True)[
        ["filial", "produto", "subgrupo", "descricao", "preco"]]
    info["grupo"] = info["subgrupo"] // 100

    return dict(info=info, meses=meses, Y=Y, L=L, D=Y + L, fim_vendas=fim,
                parcial=parcial, parcial_perdida=parcial_perdida, mes_parcial=ult_completo + 1)


# ----------------------------------------------------------------- classes
def classificar(D: np.ndarray, t: int) -> np.ndarray:
    w = D[:, max(0, t - 12):t]
    nz = (w > 0).sum(1)
    with np.errstate(invalid="ignore", divide="ignore"):
        media = np.where(nz > 0, w.sum(1) / np.maximum(nz, 1), 0)
        var = np.where(nz > 1, np.where(w > 0, (w - media[:, None]) ** 2, 0).sum(1)
                       / np.maximum(nz - 1, 1), 0)
        cv2 = np.where(media > 0, var / media ** 2, 0)
    adi = np.where(nz > 0, w.shape[1] / np.maximum(nz, 1), 99)
    cls = np.full(len(D), "IRREGULAR", dtype=object)
    cls[(adi < 1.32) & (cv2 < 0.49)] = "SUAVE"
    cls[(adi < 1.32) & (cv2 >= 0.49)] = "ERRÁTICO"
    cls[(adi >= 1.32) & (cv2 < 0.49)] = "INTERMITENTE"
    cls[(nz >= 1) & (nz <= 2)] = "RARO"
    cls[nz == 0] = "INATIVO"
    return cls


# ----------------------------------------------------------------- métodos
def _sba_todas_origens(D, alpha):
    """F[:, t] = previsão SBA feita com D[:, :t]."""
    n, T = D.shape
    z = np.zeros(n); p = np.ones(n); q = np.zeros(n); ini = np.zeros(n, bool)
    F = np.zeros((n, T + 1), dtype=np.float32)
    for t in range(T):
        F[:, t] = np.where(ini, (1 - alpha / 2) * z / p, 0)
        d = D[:, t]
        q += 1
        nz = d > 0
        prim = nz & ~ini
        z[prim] = d[prim]; p[prim] = q[prim]
        up = nz & ini
        z[up] += alpha * (d[up] - z[up]); p[up] += alpha * (q[up] - p[up])
        ini |= nz
        q[nz] = 0
    F[:, T] = np.where(ini, (1 - alpha / 2) * z / p, 0)
    return F


def _mes(cal0, j):
    return int((cal0 + j) % 12)


def _indices_sazonais(D, sg, gr, cal0, t):
    a = max(0, t - 24)
    W = D[:, a:t]
    calw = np.array([_mes(cal0, j) for j in range(a, t)])

    def idx_de(rot):
        df = pd.DataFrame(W).groupby(rot).sum()
        S = df.to_numpy()
        med = S.mean(1)
        raw = np.ones((len(S), 12))
        for c in range(12):
            col = calw == c
            if col.any():
                with np.errstate(invalid="ignore", divide="ignore"):
                    raw[:, c] = np.where(med > 0, S[:, col].mean(1) / np.where(med > 0, med, 1), 1)
        idx = 1 + SAZ_SHRINK * (raw - 1)
        return pd.DataFrame(idx / idx.mean(1, keepdims=True), index=df.index), \
            pd.Series(S.sum(1), index=df.index)

    i_sg, v_sg = idx_de(sg)
    i_g, v_g = idx_de(gr)
    i_all, _ = idx_de(np.zeros(len(sg), int))
    out = np.tile(i_all.to_numpy()[0], (len(sg), 1))
    ok = pd.Series(gr).map(v_g).to_numpy() >= SAZ_MIN_VOLUME
    out[ok] = i_g.loc[gr[ok]].to_numpy()
    ok = pd.Series(sg).map(v_sg).to_numpy() >= SAZ_MIN_VOLUME
    out[ok] = i_sg.loc[sg[ok]].to_numpy()
    return out


def _prever_base(D, sg, gr, cal0, t, H, sba):
    n = len(D)
    out = {}
    rep = lambda x: np.repeat(x[:, None], H, axis=1)
    out["MM3"] = rep(D[:, max(0, t - 3):t].mean(1))
    out["MM6"] = rep(D[:, max(0, t - 6):t].mean(1))
    out["MM12"] = rep(D[:, max(0, t - 12):t].mean(1))
    sn = np.zeros((n, H))
    for h in range(H):
        j = t + h - 12
        if 0 <= j < t:
            sn[:, h] = D[:, j]
        elif j >= t:  # horizonte > 12: repete o ano anterior do ano anterior
            sn[:, h] = D[:, j - 12] if j - 12 >= 0 else 0
    out["SNAIVE"] = sn
    out["SBA10"] = rep(sba[0.10][:, t])
    out["SBA20"] = rep(sba[0.20][:, t])
    idx = _indices_sazonais(D, sg, gr, cal0, t)
    alvo = [_mes(cal0, t + h) for h in range(H)]
    out["SAZ12"] = D[:, max(0, t - 12):t].mean(1)[:, None] * idx[:, alvo]
    c6 = [_mes(cal0, j) for j in range(max(0, t - 6), t)]
    out["SAZ6"] = (D[:, max(0, t - 6):t].sum(1) / idx[:, c6].sum(1))[:, None] * idx[:, alvo]
    out["COMB"] = (out["SAZ12"] + out["SBA10"] + out["MM6"]) / 3
    out["BETA_V1"] = 0.5 * out["MM3"] + 0.5 * sn
    return out


class Previsor:
    """Guarda o que é caro de recalcular (SBA, agregação por produto)."""

    def __init__(self, D, info, meses):
        self.D = D
        self.info = info
        self.cal0 = meses[0].month - 1
        self.sg = info["subgrupo"].to_numpy()
        self.gr = info["grupo"].to_numpy()
        prods, self.pidx = np.unique(info["produto"].to_numpy(), return_inverse=True)
        self.NP = len(prods)
        self.DP = np.zeros((self.NP, D.shape[1]))
        np.add.at(self.DP, self.pidx, D)
        # subgrupo/grupo do produto = o do primeiro par
        primeiro = pd.Series(np.arange(len(self.pidx))).groupby(self.pidx).first().to_numpy()
        self.sgP, self.grP = self.sg[primeiro], self.gr[primeiro]
        self.sba = {a: _sba_todas_origens(D, a) for a in (0.10, 0.20)}
        self.sbaP = {a: _sba_todas_origens(self.DP, a) for a in (0.10, 0.20)}

    def participacao(self, t):
        D, pidx = self.D, self.pidx
        s12 = D[:, max(0, t - 12):t].sum(1); s24 = D[:, max(0, t - 24):t].sum(1)
        t12 = np.zeros(self.NP); np.add.at(t12, pidx, s12)
        t24 = np.zeros(self.NP); np.add.at(t24, pidx, s24)
        return np.where(t12[pidx] > 0, s12 / np.maximum(t12[pidx], 1e-9),
                        np.where(t24[pidx] > 0, s24 / np.maximum(t24[pidx], 1e-9), 0))

    def prever(self, t, H):
        P = _prever_base(self.D, self.sg, self.gr, self.cal0, t, H, self.sba)
        PP = _prever_base(self.DP, self.sgP, self.grP, self.cal0, t, H, self.sbaP)
        sh = self.participacao(t)[:, None]
        for k, v in PP.items():
            P["TD_" + k] = v[self.pidx] * sh
        return P


# ----------------------------------------------------------------- reconciliação
def reconciliar(F, alvo_series, pidx):
    """Escala as séries para que a soma por produto = soma por produto de alvo_series.
    F e alvo_series: n × k. Produto sem previsão nas filiais recebe a distribuição do alvo."""
    if F.ndim == 1:
        return reconciliar(F[:, None], alvo_series[:, None], pidx)[:, 0]
    NP = int(pidx.max()) + 1
    tot = np.zeros((NP, F.shape[1])); np.add.at(tot, pidx, F)
    alvo = np.zeros((NP, F.shape[1])); np.add.at(alvo, pidx, alvo_series)
    fator = np.where(tot > 0, alvo / np.where(tot > 0, tot, 1), 0)
    out = F * fator[pidx]
    sem = (tot <= 0) & (alvo > 0)
    return np.where(sem[pidx], alvo_series, out).astype(F.dtype)


# ----------------------------------------------------------------- limpeza de picos
def tetos(D, t):
    """Teto por série = OUTLIER_K × mediana dos meses com venda até t (mín. 1 unidade)."""
    X = np.where(D[:, :t] > 0, D[:, :t], np.nan)
    with np.errstate(all="ignore"):
        import warnings
        with warnings.catch_warnings():
            warnings.simplefilter("ignore")
            med = np.nanmedian(X, 1)
    return np.where(np.isnan(med), np.inf, np.maximum(OUTLIER_K * med, 1))


def capar(D, teto):
    return np.minimum(D, teto[:, None])


# ----------------------------------------------------------------- backtest
def backtest(S, prev: Previsor):
    """Guarda, por origem, a previsão (soma H_AVAL meses) de cada método e o realizado."""
    D, meses = S["D"], S["meses"]
    n, T = D.shape
    origens = list(range(T - H_AVAL - N_ORIGENS + 1, T - H_AVAL + 1))
    O = len(origens)
    Fs = {m: np.zeros((n, O), dtype=np.float32) for m in METODOS}
    h1 = {m: np.zeros((n, O), dtype=np.float32) for m in METODOS}
    A = np.zeros((n, O), dtype=np.float32)
    CLS = np.empty((n, O), dtype=object)
    for oi, t in enumerate(origens):
        P = prev.prever(t, H_AVAL)
        A[:, oi] = D[:, t:t + H_AVAL].sum(1)
        CLS[:, oi] = classificar(D, t)
        for m in METODOS:
            Fs[m][:, oi] = P[m].sum(1)
            h1[m][:, oi] = P[m][:, 0]
    return dict(origens=origens, Fs=Fs, h1=h1, A=A, CLS=CLS)


def tabela_erros(F, A, CLS, info, origens_txt, metodo, pidx=None):
    """Erros em R$ em 3 níveis. F, A: n × O (quantidade)."""
    preco = info["preco"].to_numpy()[:, None]
    n, O = F.shape
    base = pd.DataFrame({
        "origem": np.repeat(np.array(origens_txt)[None, :], n, 0).ravel(),
        "classe": CLS.ravel(),
        "giro": np.isin(CLS, CLASSES_GIRO).ravel(),
        "subgrupo": np.repeat(info["subgrupo"].to_numpy()[:, None], O, 1).ravel(),
        "filial": np.repeat(info["filial"].to_numpy()[:, None], O, 1).ravel(),
        "A": (A * preco).ravel(), "F": (F * preco).ravel(),
    })
    base["E"] = (base["F"] - base["A"]).abs()
    out = {}
    out["item_filial"] = base.groupby(["origem", "classe", "giro", "subgrupo", "filial"],
                                      as_index=False)[["A", "F", "E"]].sum()
    g = base.groupby(["origem", "subgrupo", "filial"], as_index=False)[["A", "F"]].sum()
    g["E"] = (g["F"] - g["A"]).abs()
    out["subgrupo_filial"] = g
    if pidx is not None:
        b = base.assign(p=np.repeat(pidx[:, None], O, 1).ravel())
        # "giro" do produto = alguma filial de giro
        g = b.groupby(["origem", "p"], as_index=False).agg(
            A=("A", "sum"), F=("F", "sum"), giro=("giro", "max"), subgrupo=("subgrupo", "first"))
        g["E"] = (g["F"] - g["A"]).abs()
        out["produto_rede"] = g.groupby(["origem", "giro", "subgrupo"], as_index=False)[["A", "F", "E"]].sum()
    for v in out.values():
        v["metodo"] = metodo
    return out


def escolher_metodos(err: pd.DataFrame) -> tuple[dict, pd.DataFrame]:
    g = err.groupby(["classe", "metodo"])[["A", "F", "E"]].sum()
    g["wape"] = g["E"] / g["A"].replace(0, np.nan)
    g["vies"] = g["F"] / g["A"].replace(0, np.nan) - 1
    g["score"] = g["wape"] + g["vies"].abs()
    escolha = {}
    for c in CLASSES:
        if c not in g.index.get_level_values(0):
            escolha[c] = "BETA_V1"
            continue
        s = g.loc[c].dropna(subset=["wape"])
        ok = s[s["vies"].abs() <= VIES_MAX]
        # menor erro entre os métodos com viés aceitável; se nenhum, o melhor erro+viés
        escolha[c] = ok["wape"].idxmin() if len(ok) else s["score"].idxmin()
    return escolha, g.reset_index()


# ----------------------------------------------------------------- execução completa
def rodar(vendas, perdidas, estoque=None, arq_equiv=None, log=print):
    t0 = time.time()
    eq = detectar_equivalencias(vendas, arq_equiv)
    equiv = _resolver(eq)
    log(f"  modelo: {len(eq)} códigos substituídos unidos ao código novo")
    S = montar_series(vendas, perdidas, estoque, equiv)
    D, info, meses = S["D"], S["info"], S["meses"]
    n, T = D.shape
    log(f"  modelo: {n:,} séries filial×produto, {T} meses ({meses[0]} a {meses[-1]})")
    # backtest: picos cortados com teto calculado só com o histórico anterior às origens
    origem0 = T - H_AVAL - N_ORIGENS + 1
    prev = Previsor(capar(D, tetos(D, origem0)), info, meses)
    B = backtest(S, prev)
    O = len(B["origens"])
    origens_txt = [str(meses[t]) for t in B["origens"]]
    preco = info["preco"].to_numpy()[:, None]

    # ranking dos métodos por classe (vetorizado)
    linhas = []
    for c in CLASSES:
        mask = B["CLS"] == c
        if not mask.any():
            continue
        Ac = (B["A"] * preco)[mask].sum()
        for m in METODOS:
            Fm = B["Fs"][m] * preco
            linhas.append((c, m, Ac, Fm[mask].sum(), np.abs(Fm - B["A"] * preco)[mask].sum()))
    rk = pd.DataFrame(linhas, columns=["classe", "metodo", "A", "F", "E"])
    escolha, ranking = escolher_metodos(rk)
    log("  métodos escolhidos: " + ", ".join(f"{c}={m}" for c, m in escolha.items()))

    # previsão do MODELO no backtest = método escolhido para a classe da série em cada origem
    Fmod = np.zeros_like(B["A"]); H1mod = np.zeros_like(B["A"])
    for c, m in escolha.items():
        mask = B["CLS"] == c
        Fmod[mask] = B["Fs"][m][mask]; H1mod[mask] = B["h1"][m][mask]
    if RECONCILIAR_COM:
        Fmod = reconciliar(Fmod, B["Fs"][RECONCILIAR_COM], prev.pidx)
        H1mod = reconciliar(H1mod, B["h1"][RECONCILIAR_COM], prev.pidx)
        log(f"  reconciliação: total de cada produto na rede ajustado pelo {RECONCILIAR_COM}")
    erros = {}
    for nome, Fx in (("MODELO", Fmod), ("BETA_V1", B["Fs"]["BETA_V1"]), ("MM3", B["Fs"]["MM3"])):
        for nivel, tab in tabela_erros(Fx, B["A"], B["CLS"], info, origens_txt, nome, prev.pidx).items():
            erros.setdefault(nivel, []).append(tab)
    erros = {k: pd.concat(v, ignore_index=True) for k, v in erros.items()}

    # previsão final (picos cortados com todo o histórico)
    prev_f = Previsor(capar(D, tetos(D, T)), info, meses)
    P = prev_f.prever(T, H_FINAL)
    cls = classificar(D, T)
    met = np.array([escolha[c] for c in cls], dtype=object)
    F = np.zeros((n, H_FINAL), dtype=np.float32)
    for m in METODOS:
        sel = met == m
        if sel.any():
            F[sel] = P[m][sel]
    # previsão retroativa (1 mês à frente) com o método atual do item, para o gráfico
    BT = np.zeros((n, O), dtype=np.float32)
    for m in METODOS:
        sel = met == m
        if sel.any():
            BT[sel] = B["h1"][m][sel]
    if RECONCILIAR_COM:
        F = reconciliar(F, P[RECONCILIAR_COM].astype(np.float32), prev_f.pidx)
        BT = reconciliar(BT, B["h1"][RECONCILIAR_COM], prev.pidx)

    # faixa de erro típica (h=1) por classe: erro absoluto médio / previsão média
    faixa = {}
    A1 = D[:, B["origens"]]
    for c in CLASSES:
        sel = cls == c
        faixa[c] = float(np.clip(np.abs(BT[sel] - A1[sel]).sum() / max(BT[sel].sum(), 1e-9), 0.1, 2.0)) \
            if sel.any() else 0.5

    # acurácia por item (backtest h=1 do método escolhido)
    with np.errstate(invalid="ignore", divide="ignore"):
        acur_item = 1 - np.abs(BT - A1).sum(1) / A1.sum(1)
    info = info.copy()
    info["classe"] = cls
    info["metodo"] = met
    info["acur_item"] = np.where(A1.sum(1) > 0, acur_item, np.nan)
    # códigos antigos unidos a cada código novo
    ant = eq.groupby(eq["codigo_novo"].map(lambda c: equiv.get(c, c)))["codigo_antigo"] \
        .apply(lambda s: ", ".join(sorted(s)))
    info["codigos_antigos"] = info["produto"].map(ant).fillna("")

    meses_prev = pd.period_range(meses[-1] + 1, periods=H_FINAL, freq="M")
    log(f"  modelo pronto em {time.time()-t0:.0f}s")
    return dict(info=info, meses=meses, meses_prev=meses_prev,
                D=D.astype(np.float32), L=S["L"].astype(np.float32), F=F, BT=BT,
                origens=origens_txt, escolha=escolha, ranking=ranking,
                faixa=faixa, erros=erros, equivalencias=eq, reconciliar=RECONCILIAR_COM,
                parcial=S["parcial"].astype(np.float32),
                parcial_perdida=S["parcial_perdida"].astype(np.float32),
                mes_parcial=S["mes_parcial"], fim_vendas=S["fim_vendas"],
                gerado_em=pd.Timestamp.now())


# ----------------------------------------------------------------- demanda diária p/ o motor
def demanda_diaria_horizonte(R, dia_ini: pd.Timestamp, dias: int) -> pd.Series:
    """Demanda média diária prevista entre dia_ini e dia_ini+dias, por (filial, produto)."""
    F = R["F"].astype(float)
    total = np.zeros(len(F))
    d = pd.Timestamp(dia_ini).normalize()
    fim = d + pd.Timedelta(days=dias)
    for h, mp in enumerate(R["meses_prev"]):
        ini_m, fim_m = mp.start_time, mp.end_time.normalize() + pd.Timedelta(days=1)
        sobre = (min(fim, fim_m) - max(d, ini_m)).days
        if sobre > 0:
            total += F[:, h] * sobre / mp.days_in_month
    ult = R["meses_prev"][-1].end_time.normalize() + pd.Timedelta(days=1)
    if fim > ult:
        total += F[:, -1] * (fim - max(d, ult)).days / 30
    idx = pd.MultiIndex.from_frame(R["info"][["filial", "produto"]])
    return pd.Series(total / dias, index=idx)
