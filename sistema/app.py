"""Aplicativo local (abre no navegador): painel, item a item, acurácia e importação de arquivos.

Rodar: INICIAR.bat  (ou: python sistema/app.py)
Só aceita conexões desta máquina (127.0.0.1).
"""
from __future__ import annotations

import json
import math
import pickle
import re
import socket
import sys
import threading
import traceback
import webbrowser
from datetime import datetime
from http.server import BaseHTTPRequestHandler, ThreadingHTTPServer
from pathlib import Path
from urllib.parse import parse_qs, quote, urlparse

import numpy as np
import pandas as pd

AQUI = Path(__file__).resolve().parent
sys.path.insert(0, str(AQUI))

import atualizar as A  # noqa: E402
import config as C  # noqa: E402
import modelo as M  # noqa: E402
import painel  # noqa: E402

WEB = AQUI / "web"
PORTA = 8501


# ----------------------------------------------------------------- estado em memória
class Estado:
    def __init__(self):
        self.lock = threading.Lock()
        self.R = None
        self.df = None
        self.meta = None
        self.painel_json = None
        self.idx = {}
        self.prod_busca = None
        self.subst = {}
        self.nomes_sg = {}
        self.erro_carga = None
        self.job = {"estado": "parado", "linhas": [], "inicio": None, "fim": None, "erro": None}

    def carregar(self):
        try:
            if not (A.ARQ_MODELO.exists() and A.ARQ_PREVISAO.exists()):
                self.erro_carga = "Ainda não há previsão calculada. Vá em 'Importar arquivos' e clique em Processar."
                return
            with open(A.ARQ_MODELO, "rb") as f:
                R = pickle.load(f)
            with open(A.ARQ_PREVISAO, "rb") as f:
                df, meta = pickle.load(f)
            estoque = pd.read_pickle(A.ARQ_ESTOQUE)
            perdidas = pd.read_pickle(A.ARQ_PERDIDAS)
            info = R["info"]
            idx = {(int(f), p): i for i, (f, p) in enumerate(zip(info["filial"], info["produto"]))}
            # índice de busca por produto
            T = R["D"].shape[1]
            dem12 = R["D"][:, max(0, T - 12):].sum(1)
            pb = pd.DataFrame({"produto": info["produto"], "descricao": info["descricao"],
                               "dem12": dem12, "filial": info["filial"]})
            pb = pb.groupby("produto").agg(descricao=("descricao", "first"), dem12=("dem12", "sum"),
                                           nfil=("filial", "nunique")).reset_index()
            pb["busca"] = (pb["produto"] + " " + pb["descricao"]).str.upper()
            subst = M._resolver(R["equivalencias"]) if len(R["equivalencias"]) else {}
            dfi = df.set_index(["filial", "produto"])
            pj = json.dumps(painel.dados_painel(df, perdidas, estoque, meta), ensure_ascii=False,
                            separators=(",", ":"), default=_json_default).encode("utf-8")
            with self.lock:
                self.R, self.df, self.dfi, self.meta = R, df, dfi, meta
                self.idx, self.prod_busca, self.subst = idx, pb, subst
                self.nomes_sg = painel.nomes_subgrupos(estoque)
                self.painel_json = pj
                self.acur_cache = None
                self.erro_carga = None
        except Exception as e:
            traceback.print_exc()
            self.erro_carga = f"Erro ao carregar a base: {e}"


E = Estado()


def _json_default(o):
    if isinstance(o, (np.integer,)):
        return int(o)
    if isinstance(o, (np.floating,)):
        v = float(o)
        return None if math.isnan(v) or math.isinf(v) else v
    if isinstance(o, (pd.Timestamp, datetime)):
        return o.strftime("%d/%m/%Y")
    if isinstance(o, pd.Period):
        return str(o)
    if isinstance(o, np.ndarray):
        return o.tolist()
    return str(o)


def _limpa(x):
    """Troca NaN/inf por None recursivamente (JSON válido)."""
    if isinstance(x, float):
        return None if (math.isnan(x) or math.isinf(x)) else x
    if isinstance(x, dict):
        return {k: _limpa(v) for k, v in x.items()}
    if isinstance(x, (list, tuple)):
        return [_limpa(v) for v in x]
    if isinstance(x, np.ndarray):
        return _limpa(x.tolist())
    if isinstance(x, (np.floating,)):
        return _limpa(float(x))
    if isinstance(x, (np.integer,)):
        return int(x)
    return x


def _nome_mes(per) -> str:
    nomes = ["jan", "fev", "mar", "abr", "mai", "jun", "jul", "ago", "set", "out", "nov", "dez"]
    p = pd.Period(per, freq="M")
    return f"{nomes[p.month-1]}/{str(p.year)[2:]}"


# ----------------------------------------------------------------- item
def api_item(q):
    f = int(q.get("f", ["0"])[0])
    p = q.get("p", [""])[0].strip().upper()
    R = E.R
    if R is None:
        return 503, {"erro": E.erro_carga}
    aviso = None
    if p in E.subst:
        aviso = f"O código {p} foi substituído por {E.subst[p]}; mostrando o código novo com o histórico unido."
        p = E.subst[p]
    info = R["info"]
    filiais_prod = sorted(int(x) for x in info.loc[info["produto"] == p, "filial"])
    if not filiais_prod:
        return 404, {"erro": f"Produto {p} não encontrado no histórico."}
    if (f, p) not in E.idx:
        f = max(filiais_prod, key=lambda ff: float(R["D"][E.idx[(ff, p)], -12:].sum()))
    i = E.idx[(f, p)]
    meses = R["meses"]
    N = min(24, len(meses))
    D = R["D"][i, -N:]; L = R["L"][i, -N:]
    origens = R["origens"]
    bt_map = dict(zip(origens, R["BT"][i].tolist()))
    hist = [{"mes": str(m), "rot": _nome_mes(m), "vendido": float(D[j] - L[j]), "perdido": float(L[j]),
             "prev_passado": bt_map.get(str(m))} for j, m in enumerate(meses[-N:])]
    cls = info.at[i, "classe"]
    faixa = R["faixa"].get(cls, 0.5)
    prev = [{"mes": str(m), "rot": _nome_mes(m), "prev": float(R["F"][i, h]),
             "min": max(0.0, float(R["F"][i, h]) * (1 - faixa)), "max": float(R["F"][i, h]) * (1 + faixa)}
            for h, m in enumerate(R["meses_prev"])]
    parcial = {"mes": str(R["mes_parcial"]), "rot": _nome_mes(R["mes_parcial"]),
               "vendido": float(R["parcial"][i]), "perdido": float(R["parcial_perdida"][i]),
               "ate": R["fim_vendas"].strftime("%d/%m/%Y")}

    def linha_motor(ff):
        try:
            r = E.dfi.loc[(ff, p)]
        except KeyError:
            return None
        return {"estoque": float(r["estoque"]), "status": r["status"],
                "cobertura": None if not np.isfinite(r["cobertura_dias"]) else float(r["cobertura_dias"]),
                "falta": r["data_prevista_falta"].strftime("%d/%m/%Y") if pd.notna(r["data_prevista_falta"]) else None,
                "sugestao": float(r["sugestao_compra"]), "demanda_dia": float(r["demanda_dia"]),
                "risco": float(r["valor_em_risco"]), "sobra": r.get("sobra_outras_filiais", ""),
                "perdido90": float(r["perdido_qtd"]), "ultima_venda":
                r["ultima_venda"].strftime("%d/%m/%Y") if pd.notna(r["ultima_venda"]) else None}

    outras = []
    for ff in filiais_prod:
        j = E.idx[(ff, p)]
        m = linha_motor(ff)
        outras.append({"filial": ff, "nome": f"{ff:02d} {C.NOMES_FILIAIS.get(ff, '')}",
                       "prev1": float(R["F"][j, 0]), "prev2": float(R["F"][j, :2].sum()),
                       "vendido12": float(R["D"][j, -12:].sum()),
                       "estoque": m["estoque"] if m else 0, "status": m["status"] if m else "—",
                       "cobertura": m["cobertura"] if m else None})
    sg = int(info.at[i, "subgrupo"])
    return 200, _limpa({
        "aviso": aviso, "filial": f, "filialNome": f"{f:02d} {C.NOMES_FILIAIS.get(f, '')}",
        "produto": p, "descricao": info.at[i, "descricao"], "subgrupo": sg,
        "subgrupoNome": E.nomes_sg.get(sg, ""), "classe": cls, "classeDesc": M.CLASSES_DESC.get(cls, ""),
        "metodo": info.at[i, "metodo"], "metodoDesc": M.METODOS.get(info.at[i, "metodo"], ""),
        "acurItem": float(info.at[i, "acur_item"]) if pd.notna(info.at[i, "acur_item"]) else None,
        "reconciliar": R.get("reconciliar"),
        "faixa": faixa, "preco": float(info.at[i, "preco"]),
        "codigosAntigos": info.at[i, "codigos_antigos"],
        "hist": hist, "prev": prev, "parcial": parcial, "motor": linha_motor(f), "outras": outras,
        "diaEstoque": E.meta["data_estoque_ts"].strftime("%Y-%m-%d"),
        "params": {"lt": C.LEAD_TIME_DIAS, "seg": C.SEGURANCA_DIAS, "ciclo": C.CICLO_COMPRA_DIAS},
    })


def api_buscar(q):
    termo = q.get("q", [""])[0].strip().upper()
    if E.prod_busca is None or len(termo) < 2:
        return 200, []
    pb = E.prod_busca
    if termo in E.subst:
        termo = E.subst[termo]
    palavras = termo.split()
    m = np.ones(len(pb), bool)
    for w in palavras:
        m &= pb["busca"].str.contains(re.escape(w), regex=True).to_numpy()
    r = pb[m].copy()
    r["pref"] = ~r["produto"].str.startswith(termo)
    r = r.sort_values(["pref", "dem12"], ascending=[True, False]).head(25)
    return 200, [{"produto": a, "descricao": b, "dem12": float(c), "nfil": int(d)}
                 for a, b, c, d in zip(r["produto"], r["descricao"], r["dem12"], r["nfil"])]


# ----------------------------------------------------------------- acurácia
def _agrega(df, cols):
    g = df.groupby(cols, dropna=False)[["A", "F", "E"]].sum().reset_index()
    g["acur"] = 1 - g["E"] / g["A"].replace(0, np.nan)
    g["vies"] = g["F"] / g["A"].replace(0, np.nan) - 1
    return g


def api_acuracia(_q):
    R = E.R
    if R is None:
        return 503, {"erro": E.erro_carga}
    if getattr(E, "acur_cache", None):
        return 200, E.acur_cache
    niveis = {}
    for nivel, df in R["erros"].items():
        escopos = {"todos": df}
        if "giro" in df:
            escopos = {"giro": df[df["giro"]], "todos": df}
        out = {}
        for esc, d in escopos.items():
            geral = _agrega(d, ["metodo"])
            orig = _agrega(d, ["metodo", "origem"])
            sg = _agrega(d[d["metodo"].isin(["MODELO", "BETA_V1"])], ["metodo", "subgrupo"])
            o = {
                "geral": {r.metodo: {"acur": r.acur, "vies": r.vies, "A": r.A} for r in geral.itertuples()},
                "origens": R["origens"],
                "porOrigem": {m: [None if g.empty else g.set_index("origem").reindex(R["origens"])["acur"].tolist()][0]
                              for m, g in orig.groupby("metodo")},
                "porSubgrupo": [{"sg": int(r.subgrupo), "nome": E.nomes_sg.get(int(r.subgrupo), f"SUBGRUPO {int(r.subgrupo)}"),
                                 "acur": r.acur, "vies": r.vies, "A": r.A,
                                 "acurBeta": sg[(sg.metodo == "BETA_V1") & (sg.subgrupo == r.subgrupo)]["acur"].sum()}
                                for r in sg[sg.metodo == "MODELO"].sort_values("A", ascending=False).itertuples()],
            }
            if "filial" in d:
                fl = _agrega(d[d["metodo"] == "MODELO"], ["filial"])
                o["porFilial"] = [{"filial": int(r.filial), "nome": f"{int(r.filial):02d} {C.NOMES_FILIAIS.get(int(r.filial), '')}",
                                   "acur": r.acur, "vies": r.vies, "A": r.A} for r in fl.itertuples()]
            out[esc] = o
        niveis[nivel] = out

    # perfis de demanda (situação atual)
    info = R["info"]
    T = R["D"].shape[1]
    val12 = R["D"][:, max(0, T - 12):].sum(1) * info["preco"].to_numpy()
    rk = R["ranking"]
    perfis = []
    for c in M.CLASSES:
        sel = info["classe"] == c
        m = R["escolha"].get(c)
        rr = rk[(rk["classe"] == c) & (rk["metodo"] == m)]
        perfis.append({"classe": c, "desc": M.CLASSES_DESC[c], "series": int(sel.sum()),
                       "valor12": float(val12[sel.to_numpy()].sum()), "metodo": m,
                       "metodoDesc": M.METODOS.get(m, ""),
                       "acur": float(1 - rr["wape"].iloc[0]) if len(rr) else None,
                       "vies": float(rr["vies"].iloc[0]) if len(rr) else None})
    metodos = sorted(rk["metodo"].unique(), key=lambda m: (m.startswith("TD_"), m))
    comp = {c: {r.metodo: {"acur": 1 - r.wape, "vies": r.vies} for r in rk[rk["classe"] == c].itertuples()}
            for c in M.CLASSES}

    # acompanhamento real (previsões guardadas x realizado)
    real = []
    meses_fech = {str(m): j for j, m in enumerate(R["meses"])}
    if A.HIST_PREVISOES.exists():
        for arq in sorted(A.HIST_PREVISOES.glob("previsao_feita_em_*.pkl")):
            snap = pd.read_pickle(arq)
            for col in [c for c in snap.columns if re.match(r"^\d{4}-\d{2}$", str(c))]:
                if col not in meses_fech:
                    real.append({"feitaEm": snap.attrs.get("feita_em", ""), "mes": _nome_mes(col),
                                 "acur": None, "vies": None, "situacao": "aguardando o mês fechar"})
                    continue
                j = meses_fech[col]
                ii = [E.idx.get((int(f), p)) for f, p in zip(snap["filial"], snap["produto"])]
                ok = np.array([x is not None for x in ii])
                Aq = np.zeros(len(snap)); Aq[ok] = R["D"][[x for x in ii if x is not None], j]
                pr = snap["preco"].to_numpy(); Fq = snap[col].to_numpy()
                giro = snap["classe"].isin(M.CLASSES_GIRO).to_numpy()
                Av, Ev, Fv = (Aq * pr)[giro].sum(), (np.abs(Fq - Aq) * pr)[giro].sum(), (Fq * pr)[giro].sum()
                real.append({"feitaEm": snap.attrs.get("feita_em", ""), "mes": _nome_mes(col),
                             "acur": 1 - Ev / Av if Av else None, "vies": Fv / Av - 1 if Av else None,
                             "situacao": "medido"})

    E.acur_cache = _limpa({
        "niveis": niveis, "perfis": perfis, "metodos": metodos, "metodosDesc": M.METODOS,
        "comparativo": comp, "escolha": R["escolha"], "real": real,
        "nEquivalencias": int(len(R["equivalencias"])), "geradoEm": R["gerado_em"].strftime("%d/%m/%Y %H:%M"),
        "params": {"origens": len(R["origens"]), "horizonte": M.H_AVAL, "viesMax": M.VIES_MAX,
                   "outlierK": M.OUTLIER_K, "reconciliar": R.get("reconciliar")},
    })
    return 200, E.acur_cache


# ----------------------------------------------------------------- base / importação
def api_base(_q):
    out = {"pendentes": [], "historico": [], "base": {}, "job": E.job, "erroCarga": E.erro_carga}
    A.ENTRADA.mkdir(exist_ok=True)
    for a in sorted(A.ENTRADA.iterdir()):
        if a.is_file() and not a.name.startswith("~$"):
            t = A.detectar_tipo(a)
            out["pendentes"].append({"nome": a.name, "tipo": t, "tipoNome": A.NOMES_TIPO.get(t, "Não reconhecido"),
                                     "mb": round(a.stat().st_size / 1e6, 1)})
    if A.ARQ_HIST_IMPORT.exists():
        h = pd.read_csv(A.ARQ_HIST_IMPORT, sep=";").head(20)
        out["historico"] = h.fillna("").to_dict("records")
    b = {}
    try:
        if E.meta:
            b["vendas"] = f"até {E.meta['vendas_ate']:%d/%m/%Y}"
            b["estoque"] = f"de {E.meta['data_estoque_ts']:%d/%m/%Y}"
            b["perdidas"] = f"{E.meta['perdidas_periodo'][0]:%d/%m/%Y} a {E.meta['perdidas_periodo'][1]:%d/%m/%Y}"
        if E.R is not None:
            b["modelo"] = E.R["gerado_em"].strftime("%d/%m/%Y %H:%M")
            b["series"] = int(len(E.R["info"]))
    except Exception:
        pass
    out["base"] = b
    xl = sorted(A.SAIDA.glob("Previsao_Falta_Pecas_BETA_*.xlsx"), key=lambda p: p.stat().st_mtime) if A.SAIDA.exists() else []
    out["excel"] = xl[-1].name if xl else None
    return 200, _limpa(out)


def _rodar_job(excel: bool):
    def lg(msg):
        print(msg, flush=True)
        E.job["linhas"].append(msg)
    try:
        pend = [a for a in sorted(A.ENTRADA.iterdir()) if a.is_file()]
        A.executar(pend, mover=True, excel=excel, log_fn=lg)
        lg("Recarregando o aplicativo com os dados novos...")
        E.carregar()
        E.job.update(estado="concluido", fim=datetime.now().strftime("%H:%M:%S"))
        lg("Concluído ✔")
    except SystemExit as e:
        E.job.update(estado="erro", erro=str(e), fim=datetime.now().strftime("%H:%M:%S"))
        lg(f"ERRO: {e}")
    except Exception as e:
        traceback.print_exc()
        E.job.update(estado="erro", erro=str(e), fim=datetime.now().strftime("%H:%M:%S"))
        lg(f"ERRO: {e}")


def api_processar(q):
    if E.job["estado"] == "rodando":
        return 409, {"erro": "Já existe um processamento em andamento."}
    excel = q.get("excel", ["1"])[0] == "1"
    E.job = {"estado": "rodando", "linhas": [], "inicio": datetime.now().strftime("%H:%M:%S"), "fim": None, "erro": None}
    threading.Thread(target=_rodar_job, args=(excel,), daemon=True).start()
    return 200, {"ok": True}


def _nome_seguro(nome: str) -> str:
    nome = Path(nome).name
    nome = re.sub(r'[<>:"/\\|?*\x00-\x1f]', "_", nome).strip() or "arquivo"
    return nome


# ----------------------------------------------------------------- HTTP
class Handler(BaseHTTPRequestHandler):
    server_version = "PrevisaoPecas/1.0"

    def log_message(self, fmt, *args):
        pass

    def _send(self, code, body: bytes, ctype="application/json; charset=utf-8", extra=None):
        self.send_response(code)
        self.send_header("Content-Type", ctype)
        self.send_header("Content-Length", str(len(body)))
        self.send_header("Cache-Control", "no-store")
        for k, v in (extra or {}).items():
            self.send_header(k, v)
        self.end_headers()
        self.wfile.write(body)

    def _json(self, code, obj):
        self._send(code, json.dumps(obj, ensure_ascii=False, default=_json_default).encode("utf-8"))

    def _arquivo(self, caminho: Path, ctype, baixar=False):
        if not caminho or not caminho.exists():
            return self._json(404, {"erro": "arquivo não encontrado"})
        extra = {"Content-Disposition": f"attachment; filename*=UTF-8''{quote(caminho.name)}"} if baixar else None
        self._send(200, caminho.read_bytes(), ctype, extra)

    def do_GET(self):
        u = urlparse(self.path)
        q = parse_qs(u.query)
        try:
            if u.path in ("/", "/index.html"):
                return self._arquivo(WEB / "app.html", "text/html; charset=utf-8")
            if u.path == "/api/painel":
                if E.painel_json is None:
                    return self._json(503, {"erro": E.erro_carga})
                return self._send(200, E.painel_json)
            rotas = {"/api/item": api_item, "/api/buscar": api_buscar,
                     "/api/acuracia": api_acuracia, "/api/base": api_base}
            if u.path in rotas:
                code, obj = rotas[u.path](q)
                return self._json(code, obj)
            if u.path == "/api/status":
                return self._json(200, E.job)
            if u.path == "/download/excel":
                xl = sorted(A.SAIDA.glob("Previsao_Falta_Pecas_BETA_*.xlsx"), key=lambda p: p.stat().st_mtime)
                return self._arquivo(xl[-1] if xl else None,
                                     "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", True)
            if u.path == "/download/codigos":
                return self._arquivo(A.SAIDA / "codigos_substituidos.csv", "text/csv; charset=utf-8", True)
            if u.path == "/download/painel":
                return self._arquivo(A.SAIDA / "Painel_Falta_Pecas.html", "text/html; charset=utf-8", True)
            return self._json(404, {"erro": "rota não encontrada"})
        except Exception as e:
            traceback.print_exc()
            return self._json(500, {"erro": str(e)})

    def do_POST(self):
        u = urlparse(self.path)
        q = parse_qs(u.query)
        try:
            if u.path == "/api/upload":
                if E.job["estado"] == "rodando":
                    return self._json(409, {"erro": "Aguarde o processamento atual terminar."})
                nome = _nome_seguro(q.get("nome", ["arquivo"])[0])
                tam = int(self.headers.get("Content-Length", "0"))
                A.ENTRADA.mkdir(exist_ok=True)
                destino = A.ENTRADA / nome
                tmp = destino.with_suffix(destino.suffix + ".parcial")
                restante = tam
                with open(tmp, "wb") as f:
                    while restante > 0:
                        bloco = self.rfile.read(min(1 << 20, restante))
                        if not bloco:
                            break
                        f.write(bloco)
                        restante -= len(bloco)
                tmp.replace(destino)
                # preserva a data do arquivo original (o 1211 não traz a data dentro)
                mod = q.get("mod", [""])[0]
                if mod.isdigit():
                    import os
                    ts = int(mod) / 1000
                    os.utime(destino, (ts, ts))
                tipo = A.detectar_tipo(destino)
                return self._json(200, {"nome": nome, "tipo": tipo, "tipoNome": A.NOMES_TIPO.get(tipo, "Não reconhecido"),
                                        "mb": round(tam / 1e6, 1)})
            if u.path == "/api/remover":
                nome = _nome_seguro(q.get("nome", [""])[0])
                alvo = A.ENTRADA / nome
                if alvo.exists() and alvo.parent == A.ENTRADA:
                    alvo.unlink()
                return self._json(200, {"ok": True})
            if u.path == "/api/processar":
                code, obj = api_processar(q)
                return self._json(code, obj)
            return self._json(404, {"erro": "rota não encontrada"})
        except Exception as e:
            traceback.print_exc()
            return self._json(500, {"erro": str(e)})


def porta_livre(p):
    for porta in range(p, p + 20):
        with socket.socket(socket.AF_INET, socket.SOCK_STREAM) as s:
            if s.connect_ex(("127.0.0.1", porta)) != 0:
                return porta
    raise SystemExit("Nenhuma porta livre")


def main():
    abrir = "--sem-navegador" not in sys.argv
    print("Carregando base...", flush=True)
    E.carregar()
    if E.erro_carga:
        print(E.erro_carga)
    porta = porta_livre(PORTA)
    srv = ThreadingHTTPServer(("127.0.0.1", porta), Handler)
    url = f"http://localhost:{porta}/"
    print(f"\n  Aplicativo rodando em {url}\n  (deixe esta janela aberta; feche-a para encerrar)\n", flush=True)
    if abrir:
        threading.Timer(1.0, lambda: webbrowser.open(url)).start()
    try:
        srv.serve_forever()
    except KeyboardInterrupt:
        pass


if __name__ == "__main__":
    main()
