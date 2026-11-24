"""Consulta de Peças — últimas vendas, clientes e previsão por filial.

Aplicativo de janela (sem navegador). Lê os relatórios que estiverem na MESMA PASTA
do executável (ou na raiz do projeto, quando rodado como script):
  - 1174 (vendas)            -> .xlsx com "1174" no nome (ou reconhecido pelo cabeçalho)
  - 1211 (estoque)           -> .csv  com "1211" no nome
  - Vendas Perdidas          -> .xlsx com "perdid" no nome
e, se existir, a previsão do modelo em base\\modelo.pkl + base\\previsao.pkl
(gerada pelo ATUALIZAR.bat / aplicativo de previsão).

Na 1ª leitura de cada arquivo é criado um cache em base\\cache_consulta; nas próximas
aberturas o carregamento é quase instantâneo. Se o arquivo mudar, é relido sozinho.
"""
from __future__ import annotations

import math
import pickle
import queue
import re
import sys
import threading
import traceback
from datetime import date, datetime
from pathlib import Path

import tkinter as tk
from tkinter import ttk, messagebox

import numpy as np
import pandas as pd

AQUI = Path(__file__).resolve().parent
sys.path.insert(0, str(AQUI))
from leitores import ler_1174, ler_1211, ler_vendas_perdidas  # noqa: E402
import rede_estoque as rede  # noqa: E402

try:
    from config import NOMES_FILIAIS
except Exception:  # noqa: BLE001
    NOMES_FILIAIS = {1: "SM", 2: "SG", 3: "BG", 4: "AL", 5: "JC", 6: "PL", 7: "UR",
                     8: "IT", 9: "SB", 10: "LJ", 11: "CS", 12: "SV", 13: "ES"}

if getattr(sys, "frozen", False):
    PASTA = Path(sys.executable).resolve().parent
else:
    PASTA = AQUI.parent

VERSAO_CACHE = 3  # v2: 1174 com cliente/vendedor/cidade · v3: + custo médio da venda
EXTRAS_1174 = {"Conta": "cliente_cod", "Nome": "cliente", "Vendedor": "vendedor",
               "Cidade": "cidade", "Numero NF": "nf"}
MESES_PT = ["jan", "fev", "mar", "abr", "mai", "jun", "jul", "ago", "set", "out", "nov", "dez"]
MESES_HIST = 24  # meses de histórico no gráfico

# ------------------------------------------------------------------ cores / estilo
# Paleta neutra: fundo cinza muito claro, superfícies brancas com borda fina e UMA cor de destaque.
COR_FUNDO = "#F5F6F8"
COR_CARD = "#FFFFFF"
COR_BORDA = "#E4E7EC"
COR_TEXTO = "#101828"
COR_SUAVE = "#667085"
COR_FRACO = "#98A2B3"
COR_ACENTO = "#2E5BDB"
COR_ACENTO_HOVER = "#2348B8"
COR_ACENTO_CLARO = "#EEF3FF"
COR_HOVER = "#F2F4F7"
COR_VERDE = "#067647"
COR_VERMELHO = "#C4320A"
COR_LARANJA = "#B54708"
COR_LINHA_PAR = "#FAFBFC"
COR_SEM_ESTOQUE = "#FEF3F2"
COR_PARADO = "#A0A8B5"
COR_GRADE = "#EEF0F3"
COR_TOPO = COR_CARD
G_VENDIDO = "#3D6BE5"
G_PERDIDO = "#F5B544"
G_PREV = "#D6E2FF"
G_FAIXA = "#1F3C99"
G_ESTOQUE = "#2BA66B"
FONTE = "Segoe UI"
FONTE_FORTE = "Segoe UI Semibold"


# ------------------------------------------------------------------ formatação
def fmt_num(x, casas=2):
    if x is None or (isinstance(x, float) and pd.isna(x)):
        return ""
    s = f"{x:,.{casas}f}"
    return s.replace(",", "X").replace(".", ",").replace("X", ".")


def fmt_qtd(x):
    if x is None or pd.isna(x):
        return ""
    if abs(x - round(x)) < 1e-9:
        return fmt_num(round(x), 0)
    return fmt_num(x, 2)


def fmt_prev(x):
    """Quantidade prevista (fracionária): 1 casa abaixo de 10, inteiro acima."""
    if x is None or pd.isna(x):
        return ""
    if abs(x) < 0.05:
        return "0"
    return fmt_num(x, 1) if abs(x) < 10 else fmt_num(x, 0)


def fmt_data(d):
    if d is None or pd.isna(d):
        return ""
    return pd.Timestamp(d).strftime("%d/%m/%Y")


def fmt_moeda(x):
    if x is None or pd.isna(x):
        return ""
    return "R$ " + fmt_num(x, 2)


def fmt_compacto(x):
    """R$ 4,1 mi · R$ 850 mil · R$ 320."""
    if x is None or pd.isna(x):
        return ""
    if abs(x) >= 1e6:
        return f"R$ {fmt_num(x / 1e6, 1)} mi"
    if abs(x) >= 1e4:
        return f"R$ {fmt_num(x / 1e3, 0)} mil"
    return "R$ " + fmt_num(x, 0)


def nome_mes(per) -> str:
    per = pd.Period(per, "M")
    return f"{MESES_PT[per.month - 1]}/{str(per.year)[2:]}"


def nome_filial(n):
    return f"{int(n):02d} - {NOMES_FILIAIS.get(int(n), '')}".rstrip(" -")


def sigla_filial(n):
    return NOMES_FILIAIS.get(int(n), str(n))


def limpar_nome(s) -> str:
    """'276 GEFERSON DA SILVA' -> 'GEFERSON DA SILVA'."""
    s = "" if s is None or (isinstance(s, float) and pd.isna(s)) else str(s).strip()
    if s in ("nan", "None", "0"):
        return ""
    return re.sub(r"^\d+\s+", "", s)


def escala(vmax, n=4):
    """Escala 'redonda' para o eixo Y."""
    if vmax <= 0:
        return 1.0, [0, 1]
    bruto = vmax / n
    mag = 10 ** math.floor(math.log10(bruto))
    passo = next(m * mag for m in (1, 2, 2.5, 5, 10) if m * mag >= bruto)
    if passo < 1 and vmax >= 1:
        passo = 1
    topo = math.ceil(vmax / passo) * passo
    return topo, [i * passo for i in range(int(round(topo / passo)) + 1)]


# ------------------------------------------------------------------ dados
def _detectar(arq: Path) -> str | None:
    nome = arq.name.lower()
    if nome.startswith("~$"):
        return None
    if "perdid" in nome and arq.suffix.lower() in (".xlsx", ".xlsm"):
        return "perdidas"
    if "1211" in nome and arq.suffix.lower() in (".csv", ".txt"):
        return "1211"
    if "1174" in nome and arq.suffix.lower() in (".xlsx", ".xlsm"):
        return "1174"
    # sem pista no nome: olha o cabeçalho
    try:
        if arq.suffix.lower() in (".csv", ".txt"):
            txt = arq.open("rb").read(4000).decode("utf-8-sig", errors="ignore")
            if "Produto" in txt and ("Custo unit" in txt or "Valor total" in txt) and ";" in txt:
                return "1211"
        elif arq.suffix.lower() in (".xlsx", ".xlsm"):
            import openpyxl
            wb = openpyxl.load_workbook(arq, read_only=True, data_only=True)
            ws = wb.worksheets[0]
            linhas = []
            for i, r in enumerate(ws.iter_rows(values_only=True)):
                linhas.append(" ".join(str(x) for x in r if x is not None))
                if i >= 8:
                    break
            wb.close()
            plano = " | ".join(linhas)
            if "Venda Perdida" in plano:
                return "perdidas"
            if "Tipo de Lançamento" in plano and "Produto" in plano:
                return "1174"
    except Exception:  # noqa: BLE001
        pass
    return None


def _assinatura(*arqs: Path):
    return (VERSAO_CACHE,) + tuple((a.name, a.stat().st_size, int(a.stat().st_mtime)) for a in arqs)


def _cache(nome: str, sig, gerar, avisar, rotulo: str):
    """Lê do cache se a assinatura bater; senão gera e grava."""
    pasta_cache = PASTA / "base" / "cache_consulta"
    cache = pasta_cache / (re.sub(r"[^\w.-]+", "_", nome) + ".pkl")
    if cache.exists():
        try:
            with open(cache, "rb") as f:
                obj = pickle.load(f)
            if obj.get("sig") == sig:
                avisar(f"{rotulo}: lido do cache")
                return obj["dado"]
        except Exception:  # noqa: BLE001
            pass
    dado = gerar()
    try:
        pasta_cache.mkdir(parents=True, exist_ok=True)
        with open(cache, "wb") as f:
            pickle.dump({"sig": sig, "dado": dado}, f, protocol=pickle.HIGHEST_PROTOCOL)
    except Exception:  # noqa: BLE001
        pass
    return dado


def _ler_relatorio(arq: Path, tipo: str, avisar):
    def gerar():
        mb = arq.stat().st_size / 1e6
        extra = " — arquivo grande, pode levar alguns minutos (só desta vez)" if mb > 20 else ""
        avisar(f"Lendo {arq.name} ({mb:.0f} MB){extra}...")
        if tipo == "1174":
            df = ler_1174(arq, extras=EXTRAS_1174)
            cols = ["filial", "produto", "descricao", "marca", "subgrupo", "data", "qtd", "valor", "custo",
                    "tipo_lanc", "operacao", "local_venda", "cod_nf"] + list(EXTRAS_1174.values())
            df = df[cols].copy()
            for c in ("cliente", "vendedor", "cidade"):
                df[c] = df[c].map(limpar_nome if c == "vendedor" else
                                  (lambda s: "" if s is None else str(s).strip())).astype("category")
            for c in ("cliente_cod", "nf"):
                df[c] = pd.to_numeric(df[c], errors="coerce").fillna(0).astype("int64")
            return df
        if tipo == "1211":
            return ler_1211(arq)
        return ler_vendas_perdidas(arq)
    return _cache(arq.name, _assinatura(arq), gerar, avisar, arq.name)


def _ler_previsao(avisar):
    """Previsão do modelo do projeto (base\\modelo.pkl + base\\previsao.pkl), em forma compacta."""
    arq_m = PASTA / "base" / "modelo.pkl"
    arq_p = PASTA / "base" / "previsao.pkl"
    if not arq_m.exists():
        return None
    arqs = [arq_m] + ([arq_p] if arq_p.exists() else [])

    def gerar():
        avisar("Lendo a previsão do modelo...")
        m = pd.read_pickle(arq_m)
        df = m["info"][["filial", "produto", "classe", "codigos_antigos"]].copy()
        for h in range(m["F"].shape[1]):
            df[f"f{h}"] = m["F"][:, h].astype(float)
        df["faixa"] = df["classe"].map(m.get("faixa", {})).fillna(0.5).astype(float)
        if arq_p.exists():
            mot = pd.read_pickle(arq_p)[0]
            cols = [c for c in ["filial", "produto", "status", "data_prevista_falta", "sugestao_compra",
                                "cobertura_dias", "demanda_mes", "recorrencia"] if c in mot]
            df = df.merge(mot[cols], on=["filial", "produto"], how="left")
        meta = {"meses": [str(p) for p in m["meses_prev"]],
                "gerado": m.get("gerado_em"), "fim_vendas": m.get("fim_vendas")}
        return {"df": df, "meta": meta}
    try:
        return _cache("previsao_modelo", _assinatura(*arqs), gerar, avisar, "Previsão do modelo")
    except Exception as e:  # noqa: BLE001  (modelo de outra versão, arquivo corrompido...)
        avisar(f"Previsão do modelo indisponível: {e}")
        return None


def _mesclar_periodo(base, novo):
    if base is None or base.empty:
        return novo
    ini, fim = novo["data"].min(), novo["data"].max()
    manter = base[(base["data"] < ini) | (base["data"] > fim)]
    return pd.concat([manter.astype(object), novo.astype(object)], ignore_index=True).infer_objects()


class Base:
    def __init__(self):
        self.vendas = pd.DataFrame()
        self.perdidas = pd.DataFrame()
        self.estoque = pd.DataFrame()
        self.prev = None
        self.prev_meta = None
        self.arquivos = {"1174": [], "1211": [], "perdidas": []}
        self.produtos: list[str] = []
        self.descricoes: dict[str, str] = {}
        self.antigo_para_novo: dict[str, str] = {}
        self._idx_v = self._idx_p = self._idx_e = self._idx_f = {}

    def carregar(self, avisar):
        achados = {"1174": [], "1211": [], "perdidas": []}
        for arq in sorted(PASTA.iterdir(), key=lambda p: p.stat().st_mtime):
            if not arq.is_file() or arq.suffix.lower() not in (".xlsx", ".xlsm", ".csv", ".txt"):
                continue
            t = _detectar(arq)
            if t:
                achados[t].append(arq)
        self.arquivos = achados

        vendas = None
        for arq in achados["1174"]:
            vendas = _mesclar_periodo(vendas, _ler_relatorio(arq, "1174", avisar))
        perd = None
        for arq in achados["perdidas"]:
            perd = _mesclar_periodo(perd, _ler_relatorio(arq, "perdidas", avisar))
        est = None
        if achados["1211"]:
            est = _ler_relatorio(achados["1211"][-1], "1211", avisar)  # o mais recente
        prev = _ler_previsao(avisar)

        avisar("Organizando os dados...")
        self.vendas = vendas if vendas is not None else pd.DataFrame(
            columns=["filial", "produto", "descricao", "marca", "subgrupo", "data", "qtd", "valor", "custo",
                     "tipo_lanc", "operacao", "local_venda", "cod_nf"] + list(EXTRAS_1174.values()))
        self.perdidas = perd.drop_duplicates("cod_vp") if perd is not None else pd.DataFrame(
            columns=["cod_vp", "filial", "data", "vendedor", "cliente", "produto", "descricao",
                     "qtd", "preco", "valor", "motivo", "origem", "subgrupo", "obs"])
        self.estoque = est if est is not None else pd.DataFrame(
            columns=["filial", "subgrupo", "subgrupo_desc", "produto", "descricao",
                     "estoque", "custo_unit", "valor_estoque"])
        for df in (self.vendas, self.perdidas, self.estoque):
            if not df.empty:
                df["produto"] = df["produto"].astype(str)
                df["filial"] = df["filial"].astype(int)
        if not self.vendas.empty:
            self.vendas["data"] = pd.to_datetime(self.vendas["data"])
            self.vendas["qtd"] = self.vendas["qtd"].astype(float)
            self.vendas["valor"] = self.vendas["valor"].astype(float)
            self.vendas["custo"] = self.vendas["custo"].astype(float)  # custo médio da LINHA (total)
            self.vendas = self.vendas.sort_values("data", kind="stable").reset_index(drop=True)
        if not self.perdidas.empty:
            self.perdidas["data"] = pd.to_datetime(self.perdidas["data"])

        if prev is not None:
            self.prev, self.prev_meta = prev["df"], prev["meta"]
            self._idx_f = self.prev.groupby("produto", sort=False).indices
            for novo, antigos in zip(self.prev["produto"], self.prev["codigos_antigos"].fillna("")):
                for a in re.split(r"[,;\s]+", str(antigos)):
                    if a and a != novo:
                        self.antigo_para_novo[a.upper()] = novo

        self._idx_v = self.vendas.groupby("produto", sort=False).indices if not self.vendas.empty else {}
        self._idx_p = self.perdidas.groupby("produto", sort=False).indices if not self.perdidas.empty else {}
        self._idx_e = self.estoque.groupby("produto", sort=False).indices if not self.estoque.empty else {}

        desc = {}
        for df in (self.perdidas, self.vendas, self.estoque):  # estoque por último = prioridade
            if not df.empty:
                d = df.drop_duplicates("produto", keep="last")
                desc.update(dict(zip(d["produto"], d["descricao"].astype(str))))
        self.descricoes = desc
        self.produtos = sorted(desc)

    # ----- consultas
    def linhas(self, cod):
        v = self.vendas.iloc[self._idx_v[cod]] if cod in self._idx_v else self.vendas.iloc[0:0]
        p = self.perdidas.iloc[self._idx_p[cod]] if cod in self._idx_p else self.perdidas.iloc[0:0]
        e = self.estoque.iloc[self._idx_e[cod]] if cod in self._idx_e else self.estoque.iloc[0:0]
        return v, p, e

    def previsao(self, cod):
        """(linhas da previsão do modelo, código usado no modelo) — segue códigos substituídos."""
        if self.prev is None:
            return None, cod
        alvo = cod if cod in self._idx_f else self.antigo_para_novo.get(cod)
        if not alvo or alvo not in self._idx_f:
            return None, cod
        return self.prev.iloc[self._idx_f[alvo]], alvo

    def relacionados(self, cod, limite=10):
        if len(cod) < 5:
            return []
        out = [p for p in self.produtos if p != cod and len(p) >= 5 and (cod in p or p in cod)]
        novo = self.antigo_para_novo.get(cod)
        if novo and novo not in out and novo != cod:
            out.insert(0, novo)
        out.sort(key=lambda p: (p != novo, abs(len(p) - len(cod)), p))
        return out[:limite]

    def sugestoes(self, texto, limite=15):
        t = texto.upper().strip()
        if not t:
            return []
        por_cod = [p for p in self.produtos if t in p]
        palavras = [w for w in re.split(r"\s+", t) if w]
        por_desc = []
        if len(t) >= 3:
            por_desc = [p for p, d in self.descricoes.items()
                        if all(w in d.upper() for w in palavras) and p not in por_cod]
        return (por_cod + sorted(por_desc))[:limite]

    def resumo_periodo(self):
        def faixa(df):
            if df is None or df.empty:
                return "sem arquivo"
            return f"{fmt_data(df['data'].min())} a {fmt_data(df['data'].max())}"
        est = "sem arquivo"
        if self.arquivos["1211"]:
            a = self.arquivos["1211"][-1]
            est = f"{a.name} ({datetime.fromtimestamp(a.stat().st_mtime):%d/%m/%Y})"
        prev = "sem previsão do modelo"
        if self.prev_meta and self.prev_meta.get("gerado") is not None:
            prev = f"modelo de {fmt_data(self.prev_meta['gerado'])}"
        return {"1174": faixa(self.vendas), "perdidas": faixa(self.perdidas), "1211": est, "prev": prev}


# ------------------------------------------------------------------ componentes visuais
class BotaoPlano(tk.Label):
    """Botão plano com hover.

    estilo: "primario" (destaque cheio), "secundario" (branco com borda fina) ou
    "fantasma" (só texto, ganha fundo no hover). cor/cor_hover/fg sobrescrevem o estilo.
    """

    ESTILOS = {
        "primario": (COR_ACENTO, COR_ACENTO_HOVER, "white", None),
        "secundario": (COR_CARD, COR_HOVER, COR_TEXTO, COR_BORDA),
        "fantasma": (None, COR_HOVER, COR_SUAVE, None),
    }

    def __init__(self, master, texto, comando, estilo="primario", cor=None, cor_hover=None,
                 fg=None, fonte=(FONTE_FORTE, 10), padx=14, pady=6):
        c0, h0, f0, borda = self.ESTILOS[estilo]
        cor = cor or c0 or master.cget("bg")
        super().__init__(master, text=texto, bg=cor, fg=fg or f0, font=fonte, padx=padx, pady=pady,
                         cursor="hand2", highlightthickness=1 if borda else 0,
                         highlightbackground=borda or cor, highlightcolor=borda or cor)
        self._cor, self._hover, self._cmd = cor, cor_hover or h0, comando
        self.bind("<Enter>", lambda e: self.config(bg=self._hover))
        self.bind("<Leave>", lambda e: self.config(bg=self._cor))
        self.bind("<Button-1>", lambda e: self._cmd())


def cartao(master, **kw):
    """Superfície branca com borda fina de 1 px."""
    return tk.Frame(master, bg=COR_CARD, highlightthickness=1, highlightbackground=COR_BORDA,
                    highlightcolor=COR_BORDA, **kw)


class Grafico(tk.Canvas):
    """Canvas que se redesenha sozinho ao mudar de tamanho, com dicas (tooltip) nas barras."""

    def __init__(self, master, **kw):
        super().__init__(master, bg=COR_CARD, highlightthickness=0, **kw)
        self._fn = None
        self.bind("<Configure>", lambda e: self.redesenhar())

    def definir(self, fn):
        self._fn = fn
        self.redesenhar()

    def redesenhar(self):
        self.delete("all")
        w, h = self.winfo_width(), self.winfo_height()
        if self._fn and w > 50 and h > 50:
            self._fn(self, w, h)

    def dica(self, tag, texto):
        def mostrar(e):
            self.delete("tip")
            t = self.create_text(0, 0, text=texto, anchor="nw", font=(FONTE, 9), fill="white", tags="tip")
            x0, y0, x1, y1 = self.bbox(t)
            w, h = x1 - x0 + 16, y1 - y0 + 10
            x = min(max(e.x + 12, 4), self.winfo_width() - w - 4)
            y = e.y - h - 8 if e.y - h - 8 > 4 else e.y + 16
            r = self.create_rectangle(x, y, x + w, y + h, fill="#1D2939", outline="", tags="tip")
            self.coords(t, x + 8, y + 5)
            self.tag_raise(r)
            self.tag_raise(t)
        self.tag_bind(tag, "<Enter>", mostrar)
        self.tag_bind(tag, "<Motion>", mostrar)
        self.tag_bind(tag, "<Leave>", lambda e: self.delete("tip"))

    def legenda(self, itens, y=16):
        x = self.winfo_width() - 20
        for cor, txt, borda in reversed(itens):
            t = self.create_text(x, y, text=txt, anchor="e", font=(FONTE, 9), fill=COR_SUAVE)
            x = self.bbox(t)[0] - 6
            self.create_oval(x - 8, y - 4, x, y + 4, fill=cor, outline=borda or cor)
            x -= 22

    def titulo(self, texto, sub=""):
        self.create_text(20, 22, text=texto, anchor="w", font=(FONTE_FORTE, 11), fill=COR_TEXTO)
        if sub:
            self.create_text(20, 42, text=sub, anchor="w", font=(FONTE, 9), fill=COR_SUAVE)

    def vazio(self, texto):
        self.create_text(self.winfo_width() / 2, self.winfo_height() / 2, text=texto,
                         font=(FONTE, 11), fill=COR_SUAVE)


class _Dica:
    """Tooltip simples para widgets."""

    def __init__(self, w, texto):
        self.w, self.texto, self.tip = w, texto, None
        if texto:
            w.bind("<Enter>", self._mostrar, add="+")
            w.bind("<Leave>", self._esconder, add="+")

    def _mostrar(self, _e):
        x, y = self.w.winfo_rootx() + 10, self.w.winfo_rooty() + 24
        self.tip = tk.Toplevel(self.w)
        self.tip.wm_overrideredirect(True)
        self.tip.wm_geometry(f"+{x}+{y}")
        tk.Label(self.tip, text=self.texto, bg="#1D2939", fg="white", font=(FONTE, 9),
                 padx=8, pady=4).pack()

    def _esconder(self, _e):
        if self.tip:
            self.tip.destroy()
            self.tip = None


# ------------------------------------------------------------------ aplicativo
class App(tk.Tk):
    ABA_VISAO, ABA_RESUMO, ABA_CLIENTES, ABA_PREV, ABA_VENDAS, ABA_PERDIDAS = range(6)
    PAGINAS = [("consulta", "Consulta de peça"), ("transf", "Transferências"), ("saude", "Saúde do estoque")]
    TODOS_SUB = "Todos os subgrupos"
    MAX_LINHAS = 3000  # linhas mostradas nas tabelas da rede (o Excel exporta tudo)

    def __init__(self):
        super().__init__()
        self.title("Consulta de Peças — Últimas vendas por filial")
        self.configure(bg=COR_FUNDO)
        self.geometry("1360x860")
        self.minsize(1100, 700)
        try:
            self.state("zoomed")
        except tk.TclError:
            pass
        self.base = Base()
        self.fila = queue.Queue()
        self.cod_atual = None
        self.filial_sel = None  # None = todas
        self.lista_codigos: list[str] = []
        self.atual = None  # dados calculados da peça consultada
        self.modo_prev = "periodo"  # "periodo" (média do período do 1174) ou "modelo"
        self.meses_disp: list = []  # meses disponíveis no 1174 (pd.Period)
        self._estilos()
        self._montar()
        self.after(100, self.recarregar)

    # ---------------- estilos
    def _estilos(self):
        st = ttk.Style(self)
        st.theme_use("clam")
        # tabelas: sem moldura, cabeçalho branco com texto suave, linhas altas
        st.layout("Treeview", [("Treeview.treearea", {"sticky": "nswe"})])
        st.configure("Treeview", font=(FONTE, 10), rowheight=30, background=COR_CARD,
                     fieldbackground=COR_CARD, foreground=COR_TEXTO, borderwidth=0, relief="flat")
        st.configure("Treeview.Heading", font=(FONTE_FORTE, 9), background=COR_CARD,
                     foreground=COR_SUAVE, relief="flat", padding=(8, 8), borderwidth=1,
                     bordercolor=COR_BORDA, lightcolor=COR_CARD, darkcolor=COR_BORDA)
        st.map("Treeview.Heading", background=[("active", COR_HOVER)])
        st.map("Treeview", background=[("selected", COR_ACENTO_CLARO)],
               foreground=[("selected", COR_TEXTO)])
        # abas nativas escondidas (usamos uma barra própria, sublinhada)
        st.configure("Limpo.TNotebook", background=COR_CARD, borderwidth=0, tabmargins=0,
                     bordercolor=COR_CARD, lightcolor=COR_CARD, darkcolor=COR_CARD)
        st.layout("Limpo.TNotebook.Tab", [])
        st.layout("Limpo.TNotebook", [("Notebook.client", {"sticky": "nswe"})])
        # barra de rolagem fina, sem setas
        st.layout("Fino.Vertical.TScrollbar", [("Vertical.Scrollbar.trough", {"sticky": "ns", "children": [
            ("Vertical.Scrollbar.thumb", {"expand": "1", "sticky": "nswe"})]})])
        st.configure("Fino.Vertical.TScrollbar", troughcolor=COR_CARD, background="#D0D5DD",
                     bordercolor=COR_CARD, lightcolor="#D0D5DD", darkcolor="#D0D5DD", gripcount=0,
                     width=8, arrowsize=8)
        st.map("Fino.Vertical.TScrollbar", background=[("active", COR_FRACO)])
        # combobox branca de borda fina
        st.configure("TCombobox", fieldbackground=COR_CARD, background=COR_CARD, foreground=COR_TEXTO,
                     bordercolor=COR_BORDA, lightcolor=COR_CARD, darkcolor=COR_CARD,
                     arrowcolor=COR_SUAVE, padding=(6, 3), arrowsize=12)
        st.map("TCombobox", fieldbackground=[("readonly", COR_CARD)], background=[("readonly", COR_CARD)],
               selectbackground=[("readonly", COR_CARD)], selectforeground=[("readonly", COR_TEXTO)],
               bordercolor=[("focus", COR_ACENTO), ("hover", "#C9CED6")])
        self.option_add("*TCombobox*Listbox.font", (FONTE, 10))
        self.option_add("*TCombobox*Listbox.selectBackground", COR_ACENTO_CLARO)
        self.option_add("*TCombobox*Listbox.selectForeground", COR_TEXTO)
        self.option_add("*TCombobox*Listbox.borderWidth", 0)
        st.configure("Horizontal.TProgressbar", troughcolor=COR_GRADE, background=COR_ACENTO,
                     bordercolor=COR_GRADE, lightcolor=COR_ACENTO, darkcolor=COR_ACENTO, thickness=4)

    # ---------------- layout
    def _montar(self):
        # cabeçalho branco com linha inferior
        topo = tk.Frame(self, bg=COR_CARD)
        topo.pack(fill="x")
        tk.Frame(self, bg=COR_BORDA, height=1).pack(fill="x")
        tk.Label(topo, text="Consulta de Peças", bg=COR_CARD, fg=COR_TEXTO,
                 font=(FONTE_FORTE, 14)).pack(side="left", padx=(24, 22), pady=14)
        nav = tk.Frame(topo, bg=COR_CARD, highlightthickness=1, highlightbackground=COR_BORDA)
        nav.pack(side="left", pady=14)
        self.btn_pag = {}
        for pag, txt in self.PAGINAS:
            b = tk.Label(nav, text=txt, font=(FONTE_FORTE, 10), padx=14, pady=5, cursor="hand2",
                         bg=COR_CARD, fg=COR_SUAVE)
            b.pack(side="left")
            b.bind("<Button-1>", lambda e, p=pag: self._ir_pagina(p))
            self.btn_pag[pag] = b
        self.cbs_periodo: list = []
        self.cbs_subgrupo: list = []
        self.var_subgrupo = tk.StringVar(value=self.TODOS_SUB)
        self._sub_aplicado = self.TODOS_SUB
        self.var_recorr = tk.StringVar(value="2 meses ou mais")
        BotaoPlano(topo, "Recarregar arquivos", self.recarregar, estilo="secundario",
                   fonte=(FONTE_FORTE, 9), padx=12, pady=5).pack(side="right", padx=(8, 24))
        self.lbl_bases = tk.Label(topo, text="", bg=COR_CARD, fg=COR_FRACO, font=(FONTE, 8),
                                  justify="right")
        self.lbl_bases.pack(side="right", padx=4)

        externo = tk.Frame(self, bg=COR_FUNDO)
        externo.pack(fill="both", expand=True, padx=24, pady=(16, 12))
        self.paginas = {p: tk.Frame(externo, bg=COR_FUNDO) for p, _ in self.PAGINAS}
        corpo = self.paginas["consulta"]

        # busca
        busca = cartao(corpo)
        busca.pack(fill="x")
        linha = tk.Frame(busca, bg=COR_CARD)
        linha.pack(fill="x", padx=16, pady=12)
        self.var_busca = tk.StringVar()
        self.caixa = tk.Frame(linha, bg=COR_CARD, highlightthickness=1, highlightbackground="#D0D5DD",
                              highlightcolor=COR_ACENTO)
        self.caixa.pack(side="left")
        tk.Label(self.caixa, text="Código", bg=COR_CARD, fg=COR_FRACO,
                 font=(FONTE, 9)).pack(side="left", padx=(12, 0))
        self.ent = tk.Entry(self.caixa, textvariable=self.var_busca, font=(FONTE, 15), width=26,
                            relief="flat", bg=COR_CARD, fg=COR_TEXTO, insertwidth=1, bd=0,
                            highlightthickness=0, insertbackground=COR_TEXTO)
        self.ent.pack(side="left", ipady=6, padx=(8, 12))
        self.ent.bind("<FocusIn>", lambda e: self.caixa.config(highlightbackground=COR_ACENTO))
        self.ent.bind("<FocusOut>", lambda e: self.caixa.config(highlightbackground="#D0D5DD"))
        self.ent.bind("<Return>", lambda e: self.consultar())
        self.ent.bind("<KP_Enter>", lambda e: self.consultar())
        self.ent.bind("<<Paste>>", self._ao_colar)
        BotaoPlano(linha, "Consultar", self.consultar, pady=8).pack(side="left", padx=(10, 6))
        BotaoPlano(linha, "Colar e consultar", self.colar_consultar, estilo="secundario",
                   pady=8).pack(side="left", padx=6)
        BotaoPlano(linha, "Limpar", self.limpar, estilo="fantasma", pady=8).pack(side="left", padx=2)
        tk.Label(linha, text="Ctrl+V cola e já consulta  ·  aceita vários códigos, parte do código "
                             "ou da descrição", bg=COR_CARD, fg=COR_FRACO,
                 font=(FONTE, 9)).pack(side="left", padx=14)
        self.fr_chips = tk.Frame(busca, bg=COR_CARD)
        self.fr_chips.pack(fill="x", padx=16)

        # cabeçalho da peça (direto no fundo, sem caixa)
        topo_card = tk.Frame(corpo, bg=COR_FUNDO)
        topo_card.pack(fill="x", pady=(18, 10))
        self.lbl_cod = tk.Label(topo_card, text="Nenhuma peça consultada", bg=COR_FUNDO, fg=COR_TEXTO,
                                font=(FONTE_FORTE, 18))
        self.lbl_cod.pack(side="left")
        self.lbl_desc = tk.Label(topo_card, text="", bg=COR_FUNDO, fg=COR_SUAVE, font=(FONTE, 12))
        self.lbl_desc.pack(side="left", padx=14, pady=(4, 0))
        self.fr_rel = tk.Frame(topo_card, bg=COR_FUNDO)
        self.fr_rel.pack(side="right")

        # indicadores
        self.fr_kpis = tk.Frame(corpo, bg=COR_FUNDO)
        self.fr_kpis.pack(fill="x")
        self.kpis = {}
        itens = [("estoque", "Estoque na rede"), ("ult", "Última venda"), ("v12", "Vendido 12 meses"),
                 ("custo", "Custo médio vendido"), ("cli", "Clientes 12 meses"),
                 ("perd", "Vendas perdidas"), ("prev", "Previsão")]
        for i, (chave, titulo) in enumerate(itens):
            f = cartao(self.fr_kpis, padx=14, pady=10)
            f.grid(row=0, column=i, sticky="nsew", padx=(0 if i == 0 else 5, 0 if i == len(itens) - 1 else 5))
            self.fr_kpis.columnconfigure(i, weight=1, uniform="kpi")
            t = tk.Label(f, text=titulo, bg=COR_CARD, fg=COR_SUAVE, font=(FONTE, 9))
            t.pack(anchor="w")
            v = tk.Label(f, text="—", bg=COR_CARD, fg=COR_TEXTO, font=(FONTE_FORTE, 17))
            v.pack(anchor="w", pady=(2, 0))
            s = tk.Label(f, text="", bg=COR_CARD, fg=COR_FRACO, font=(FONTE, 8), anchor="w", justify="left")
            s.pack(anchor="w", fill="x")
            s.bind("<Configure>", lambda e, lb=s: lb.config(wraplength=max(80, e.width)))
            self.kpis[chave] = (v, s, t)

        # barra de controles: filtro de filial · previsão · copiar
        cab = tk.Frame(corpo, bg=COR_FUNDO)
        cab.pack(fill="x", pady=(16, 8))
        self.lbl_filtro = tk.Label(cab, text="Filial: todas", bg=COR_FUNDO, fg=COR_TEXTO,
                                   font=(FONTE_FORTE, 10))
        self.lbl_filtro.pack(side="left")
        BotaoPlano(cab, "Ver todas", self._todas_filiais, estilo="fantasma", cor=COR_FUNDO,
                   fg=COR_ACENTO, fonte=(FONTE_FORTE, 9), padx=8, pady=3).pack(side="left", padx=(4, 18))

        tk.Label(cab, text="Previsão", bg=COR_FUNDO, fg=COR_SUAVE, font=(FONTE, 9)).pack(side="left", padx=(0, 8))
        seg = tk.Frame(cab, bg=COR_CARD, highlightthickness=1, highlightbackground=COR_BORDA)
        seg.pack(side="left")
        self.btn_modo = {}
        for modo, txt in [("periodo", "Período do 1174"), ("modelo", "Modelo do projeto")]:
            b = tk.Label(seg, text=txt, font=(FONTE_FORTE, 9), padx=12, pady=4, cursor="hand2")
            b.pack(side="left")
            b.bind("<Button-1>", lambda e, m=modo: self._definir_modo(m))
            self.btn_modo[modo] = b
        self.fr_periodo = tk.Frame(cab, bg=COR_FUNDO)
        self.fr_periodo.pack(side="left", padx=(16, 0))
        tk.Label(self.fr_periodo, text="Base", bg=COR_FUNDO, fg=COR_SUAVE, font=(FONTE, 9)).pack(side="left")
        self.var_de, self.var_ate, self.var_atender = tk.StringVar(), tk.StringVar(), tk.StringVar(value="3 meses")
        self.cb_de = ttk.Combobox(self.fr_periodo, textvariable=self.var_de, width=7, state="readonly",
                                  font=(FONTE, 10))
        self.cb_de.pack(side="left", padx=(6, 0))
        tk.Label(self.fr_periodo, text="até", bg=COR_FUNDO, fg=COR_SUAVE, font=(FONTE, 9)).pack(side="left", padx=6)
        self.cb_ate = ttk.Combobox(self.fr_periodo, textvariable=self.var_ate, width=7, state="readonly",
                                   font=(FONTE, 10))
        self.cb_ate.pack(side="left")
        for cb in (self.cb_de, self.cb_ate):
            cb.bind("<<ComboboxSelected>>", self._ao_mudar_periodo)
            self.cbs_periodo.append(cb)
        tk.Label(cab, text="Atender", bg=COR_FUNDO, fg=COR_SUAVE, font=(FONTE, 9)).pack(side="left", padx=(16, 0))
        self.cb_atender = ttk.Combobox(cab, textvariable=self.var_atender, width=9, state="readonly",
                                       font=(FONTE, 10),
                                       values=[f"{n} {'mês' if n == 1 else 'meses'}" for n in range(1, 13)])
        self.cb_atender.pack(side="left", padx=(6, 0))
        self.cb_atender.bind("<<ComboboxSelected>>", lambda e: self._recalcular())
        self._pintar_modo()
        BotaoPlano(cab, "Copiar tabela", self.copiar_aba, estilo="secundario", fonte=(FONTE_FORTE, 9),
                   padx=12, pady=4).pack(side="right")

        # área das abas: cartão branco com barra de abas sublinhada
        area = cartao(corpo)
        area.pack(fill="both", expand=True)
        barra = tk.Frame(area, bg=COR_CARD)
        barra.pack(fill="x", padx=12)
        tk.Frame(area, bg=COR_BORDA, height=1).pack(fill="x")
        self.nb = ttk.Notebook(area, style="Limpo.TNotebook")
        self.nb.pack(fill="both", expand=True)
        nomes = ["Visão geral", "Resumo por filial", "Clientes por filial", "Previsão",
                 "Últimas vendas", "Vendas perdidas"]
        abas = [tk.Frame(self.nb, bg=COR_CARD) for _ in nomes]
        self.abas_ui = []
        for i, (a, t) in enumerate(zip(abas, nomes)):
            self.nb.add(a, text=t)
            col = tk.Frame(barra, bg=COR_CARD, cursor="hand2")
            col.pack(side="left", padx=(0, 4))
            lb = tk.Label(col, text=t, bg=COR_CARD, fg=COR_SUAVE, font=(FONTE_FORTE, 10), padx=10, pady=10,
                          cursor="hand2")
            lb.pack()
            sub = tk.Frame(col, bg=COR_CARD, height=2)
            sub.pack(fill="x")
            for w in (col, lb):
                w.bind("<Button-1>", lambda e, k=i: self.nb.select(k))
            lb.bind("<Enter>", lambda e, k=i: self._hover_aba(k, True))
            lb.bind("<Leave>", lambda e, k=i: self._hover_aba(k, False))
            self.abas_ui.append((lb, sub))
        self.nb.bind("<<NotebookTabChanged>>", lambda e: self._pintar_abas())
        aba_vis, aba_res, aba_cli, aba_prev, aba_v, aba_p = abas

        # visão geral: gráfico mensal (esq.) + filiais (dir.), separados por linha fina
        aba_vis.columnconfigure(0, weight=62)
        aba_vis.columnconfigure(2, weight=38)
        aba_vis.rowconfigure(0, weight=1)
        self.g_mensal = Grafico(aba_vis)
        self.g_mensal.grid(row=0, column=0, sticky="nsew", padx=(4, 0), pady=6)
        tk.Frame(aba_vis, bg=COR_BORDA, width=1).grid(row=0, column=1, sticky="ns", pady=16)
        self.g_filiais = Grafico(aba_vis)
        self.g_filiais.grid(row=0, column=2, sticky="nsew", padx=(0, 4), pady=6)

        # resumo
        fr = tk.Frame(aba_res, bg=COR_CARD)
        fr.pack(fill="x", padx=12, pady=(10, 0))
        self._legenda(fr, [(COR_SEM_ESTOQUE, "vende mas está sem estoque"), (COR_LINHA_PAR, "sem movimento")])
        self.tv_res = self._tabela(aba_res, [
            ("filial", "Filial", 80, "w"), ("estoque", "Estoque", 70, "e"),
            ("ult_data", "Última venda", 100, "center"), ("ult_dias", "Há (dias)", 70, "e"),
            ("ult_qtd", "Qtd", 50, "e"), ("ult_cli", "Último cliente", 230, "w"),
            ("ult_unit", "Preço unit. últ.", 105, "e"), ("custo", "Custo médio 12m", 110, "e"),
            ("v30", "Vend. 30d", 72, "e"), ("v90", "Vend. 90d", 72, "e"),
            ("v12", "Vend. 12m", 72, "e"), ("vtot", "Vend. total", 78, "e"),
            ("ncli", "Clientes 12m", 88, "e"),
            ("prev", "Prev. próx. mês", 100, "e"),
            ("perd_data", "Últ. venda perdida", 115, "center"),
            ("perd_qtd", "Qtd perdida", 82, "e")], altura=14)
        self.tv_res.bind("<<TreeviewSelect>>", self._ao_selecionar_filial)
        self.tv_res.bind("<Double-1>", lambda e: self.nb.select(self.ABA_CLIENTES))

        # clientes: árvore filial -> clientes (esq.) + gráfico maiores compradores (dir.)
        aba_cli.columnconfigure(0, weight=64)
        aba_cli.columnconfigure(2, weight=36)
        aba_cli.rowconfigure(0, weight=1)
        fr_cli = tk.Frame(aba_cli, bg=COR_CARD)
        fr_cli.grid(row=0, column=0, sticky="nsew")
        tk.Frame(aba_cli, bg=COR_BORDA, width=1).grid(row=0, column=1, sticky="ns", pady=16)
        bar = tk.Frame(fr_cli, bg=COR_CARD)
        bar.pack(fill="x", padx=12, pady=(10, 0))
        tk.Label(bar, text="Duplo clique numa filial para abrir os clientes · mais recentes primeiro",
                 bg=COR_CARD, fg=COR_FRACO, font=(FONTE, 9)).pack(side="left")
        BotaoPlano(bar, "Recolher", lambda: self._expandir_cli(False), estilo="fantasma",
                   fonte=(FONTE_FORTE, 9), padx=8, pady=2).pack(side="right", padx=2)
        BotaoPlano(bar, "Expandir", lambda: self._expandir_cli(True), estilo="fantasma",
                   fonte=(FONTE_FORTE, 9), padx=8, pady=2).pack(side="right", padx=2)
        self.tv_cli = self._tabela(fr_cli, [
            ("conta", "Conta", 70, "e"), ("compras", "Compras", 80, "e"), ("qtd", "Qtd", 55, "e"),
            ("valor", "Valor", 105, "e"), ("ult", "Última compra", 115, "center"),
            ("dias", "Há (dias)", 80, "e"), ("preco", "Últ. preço", 95, "e"),
            ("vendedor", "Vendedor", 180, "w"), ("cidade", "Cidade", 125, "w")],
            altura=14, arvore=("Filial / Cliente", 300))
        self.tv_cli.tag_configure("grupo", font=(FONTE_FORTE, 10), background="#F7F8FA")
        self.g_cli = Grafico(aba_cli)
        self.g_cli.grid(row=0, column=2, sticky="nsew", padx=(0, 4), pady=6)

        # previsão: nota + tabela
        self.lbl_prev = tk.Label(aba_prev, text="", bg=COR_CARD, fg=COR_SUAVE, font=(FONTE, 9),
                                 justify="left", anchor="w")
        self.lbl_prev.pack(fill="x", padx=14, pady=(10, 0))
        self.lbl_prev.bind("<Configure>", lambda e: self.lbl_prev.config(wraplength=max(200, e.width - 10)))
        self.tv_prev = self._tabela(aba_prev, [("filial", "Filial", 80, "w")], altura=14)
        for tag, cor in [("st_rup", "#FEF3F2"), ("st_crit", "#FFF6ED"), ("st_aten", "#FFFAEB")]:
            self.tv_prev.tag_configure(tag, background=cor)

        self.tv_vend = self._tabela(aba_v, [
            ("data", "Data", 90, "center"), ("filial", "Filial", 70, "w"),
            ("cliente", "Cliente", 250, "w"), ("qtd", "Qtd", 50, "e"),
            ("unit", "Preço unit.", 95, "e"), ("custo", "Custo unit.", 95, "e"),
            ("valor", "Valor líquido", 105, "e"),
            ("tipo", "Tipo", 125, "w"), ("vendedor", "Vendedor", 175, "w"),
            ("cidade", "Cidade", 115, "w"), ("operacao", "Operação", 240, "w"),
            ("nf", "NF", 70, "e")], altura=14)
        self.tv_perd = self._tabela(aba_p, [
            ("data", "Data", 95, "center"), ("filial", "Filial", 80, "w"), ("qtd", "Qtd", 60, "e"),
            ("preco", "Preço unit.", 100, "e"), ("valor", "Total", 100, "e"),
            ("cliente", "Cliente", 260, "w"), ("vendedor", "Vendedor", 180, "w"),
            ("origem", "Origem", 90, "w"), ("motivo", "Motivo", 230, "w")], altura=14)
        self._pintar_abas()

        rodape = tk.Frame(self, bg=COR_CARD)
        rodape.pack(fill="x", side="bottom")
        tk.Frame(rodape, bg=COR_BORDA, height=1).pack(fill="x")
        self.lbl_status = tk.Label(rodape, text="", bg=COR_CARD, fg=COR_FRACO, font=(FONTE, 9),
                                   anchor="w", padx=24, pady=5)
        self.lbl_status.pack(fill="x")

        # tela de carregamento
        self.overlay = tk.Frame(self, bg=COR_FUNDO)
        miolo = tk.Frame(self.overlay, bg=COR_FUNDO)
        miolo.place(relx=0.5, rely=0.4, anchor="center")
        tk.Label(miolo, text="Carregando as bases", bg=COR_FUNDO, fg=COR_TEXTO,
                 font=(FONTE_FORTE, 15)).pack(pady=(0, 6))
        self.lbl_carga = tk.Label(miolo, text="", bg=COR_FUNDO, fg=COR_SUAVE,
                                  font=(FONTE, 10), wraplength=700)
        self.lbl_carga.pack()
        self.pb = ttk.Progressbar(miolo, mode="indeterminate", length=360)
        self.pb.pack(pady=18)
        self._graficos_vazios("Cole um código de peça para ver os gráficos.")
        self._montar_transf(self.paginas["transf"])
        self._montar_saude(self.paginas["saude"])
        self.pagina = None
        self._ir_pagina("consulta")

    # ---------------- barra de abas própria
    def _pintar_abas(self):
        try:
            atual = self.nb.index(self.nb.select())
        except tk.TclError:
            atual = 0
        for i, (lb, sub) in enumerate(self.abas_ui):
            ativo = i == atual
            lb.config(fg=COR_ACENTO if ativo else COR_SUAVE)
            sub.config(bg=COR_ACENTO if ativo else COR_CARD)

    def _hover_aba(self, i, dentro):
        lb, _ = self.abas_ui[i]
        if i != self.nb.index(self.nb.select()):
            lb.config(fg=COR_TEXTO if dentro else COR_SUAVE)

    def _aba_txt(self, i, texto):
        self.abas_ui[i][0].config(text=texto)
        self.nb.tab(i, text=texto)

    def _legenda(self, pai, itens):
        fundo = pai.cget("bg")
        for cor, txt in itens:
            f = tk.Frame(pai, bg=fundo)
            f.pack(side="right", padx=8)
            tk.Label(f, text="  ", bg=cor, highlightthickness=1, font=(FONTE, 7),
                     highlightbackground=COR_BORDA).pack(side="left")
            tk.Label(f, text=" " + txt, bg=fundo, fg=COR_SUAVE, font=(FONTE, 9)).pack(side="left")

    def _tabela(self, pai, cols, altura=10, arvore=None):
        fr = tk.Frame(pai, bg=COR_CARD)
        fr.pack(fill="both", expand=True, padx=8, pady=(6, 6))
        tv = ttk.Treeview(fr, columns=[c[0] for c in cols],
                          show="tree headings" if arvore else "headings", height=altura)
        if arvore:
            tv.heading("#0", text=arvore[0], anchor="w")
            tv.column("#0", width=arvore[1], stretch=True)
        for cid, titulo, larg, anc in cols:
            tv.heading(cid, text=titulo, anchor=anc, command=lambda c=cid, t=tv: self._ordenar(t, c))
            tv.column(cid, width=larg, anchor=anc, stretch=True)
        sb = ttk.Scrollbar(fr, orient="vertical", command=tv.yview, style="Fino.Vertical.TScrollbar")
        tv.configure(yscrollcommand=sb.set)
        tv.pack(side="left", fill="both", expand=True)
        sb.pack(side="right", fill="y")
        tv.tag_configure("par", background=COR_LINHA_PAR)
        tv.tag_configure("impar", background=COR_CARD)
        tv.tag_configure("semestoque", background=COR_SEM_ESTOQUE)
        tv.tag_configure("parado", foreground=COR_PARADO, background=COR_LINHA_PAR)
        tv.tag_configure("negativo", foreground=COR_VERMELHO)
        tv.tag_configure("total", font=(FONTE_FORTE, 10), background="#F2F4F7")
        return tv

    def _ordenar(self, tv, col, rev=None):
        if rev is None:
            rev = getattr(tv, "_rev_" + col, False)

        def chave(k):
            s = tv.set(k, col)
            m = re.fullmatch(r"(\d{2})/(\d{2})/(\d{4})", s)
            if m:
                return (0, f"{m.group(3)}{m.group(2)}{m.group(1)}")
            t = s.replace("R$", "").replace(".", "").replace(",", ".").strip()
            try:
                return (0, float(t))
            except ValueError:
                return (1 if not s else 0, s)
        pais = [""] + [k for k in tv.get_children("") if tv.get_children(k)]
        for pai in pais:
            itens = [k for k in tv.get_children(pai) if "total" not in tv.item(k, "tags")]
            totais = [k for k in tv.get_children(pai) if "total" in tv.item(k, "tags")]
            if pai == "" and len(pais) > 1:
                continue  # na árvore de clientes, ordena só dentro de cada filial
            itens.sort(key=chave, reverse=rev)
            for i, k in enumerate(itens + totais):
                tv.move(k, pai, i)
        setattr(tv, "_rev_" + col, not rev)

    # ---------------- carregamento
    def recarregar(self):
        self.overlay.place(x=0, y=0, relwidth=1, relheight=1)
        self.overlay.lift()
        self.pb.start(12)
        self.lbl_carga.config(text=f"Procurando relatórios em:\n{PASTA}")

        def trabalho():
            try:
                nova = Base()
                nova.carregar(lambda m: self.fila.put(("msg", m)))
                self.fila.put(("ok", nova))
            except Exception as e:  # noqa: BLE001
                self.fila.put(("erro", f"{e}\n\n{traceback.format_exc()}"))
        threading.Thread(target=trabalho, daemon=True).start()
        self.after(150, self._checar_fila)

    def _checar_fila(self):
        try:
            while True:
                tipo, dado = self.fila.get_nowait()
                if tipo == "msg":
                    self.lbl_carga.config(text=dado)
                elif tipo == "ok":
                    self.base = dado
                    self._fim_carga()
                    return
                elif tipo == "erro":
                    self.pb.stop()
                    self.overlay.place_forget()
                    messagebox.showerror("Erro ao carregar", dado)
                    return
        except queue.Empty:
            pass
        self.after(150, self._checar_fila)

    def _fim_carga(self):
        self.pb.stop()
        self.overlay.place_forget()
        r = self.base.resumo_periodo()
        self.lbl_bases.config(text=f"Vendas 1174: {r['1174']}   ·   Perdidas: {r['perdidas']}\n"
                                   f"Estoque: {r['1211']}   ·   Previsão: {r['prev']}")
        faltando = [n for k, n in [("1174", "1174 (vendas)"), ("1211", "1211 (estoque)"),
                                   ("perdidas", "Vendas Perdidas")] if not self.base.arquivos[k]]
        n_arq = sum(len(v) for v in self.base.arquivos.values())
        self.status(f"{n_arq} arquivo(s) lido(s) de {PASTA}  ·  {len(self.base.produtos):,} códigos "
                    "na base".replace(",", "."))
        if faltando:
            messagebox.showwarning("Arquivos não encontrados",
                                   "Não achei na pasta:\n  • " + "\n  • ".join(faltando) +
                                   f"\n\nPasta lida:\n{PASTA}\n\nColoque os relatórios nessa pasta e "
                                   "clique em 'Recarregar arquivos'.")
        self._preparar_periodos()
        self.subgrupos_rot = [self.TODOS_SUB] + [f"{c} · {d}" if d else c
                                                 for c, d in rede.lista_subgrupos(self.base)]
        for cb in self.cbs_subgrupo:
            cb["values"] = self.subgrupos_rot
        if self.var_subgrupo.get() not in self.subgrupos_rot:
            self.var_subgrupo.set(self.TODOS_SUB)
        if self.cod_atual:
            self.mostrar(self.cod_atual)
        if self.pagina == "consulta":
            self.ent.focus_set()
        else:
            self._atualizar_pagina()

    # ---------------- seletor de período
    def _preparar_periodos(self):
        v = self.base.vendas
        if v.empty:
            self.meses_disp = []
        else:
            self.meses_disp = list(pd.period_range(v["data"].min(), v["data"].max(), freq="M"))
        rot = [nome_mes(m) for m in self.meses_disp]
        for cb in self.cbs_periodo:
            cb["values"] = rot
        if rot and (self.var_de.get() not in rot or self.var_ate.get() not in rot):
            self.var_ate.set(rot[-1])
            self.var_de.set(rot[max(0, len(rot) - 12)])  # padrão: últimos 12 meses
        if self.base.prev is None and self.modo_prev == "modelo":
            self.modo_prev = "periodo"
        self._pintar_modo()

    def _periodo_base(self):
        rot = [nome_mes(m) for m in self.meses_disp]
        if not rot:
            return None, None
        i0 = rot.index(self.var_de.get()) if self.var_de.get() in rot else 0
        i1 = rot.index(self.var_ate.get()) if self.var_ate.get() in rot else len(rot) - 1
        if i0 > i1:
            i0, i1 = i1, i0
        return self.meses_disp[i0], self.meses_disp[i1]

    def _meses_atender(self) -> int:
        m = re.match(r"\d+", self.var_atender.get())
        return int(m.group()) if m else 3

    def _pintar_modo(self):
        for modo, b in self.btn_modo.items():
            ativo = modo == self.modo_prev
            indisponivel = modo == "modelo" and self.base.prev is None
            b.config(bg=COR_ACENTO_CLARO if ativo else COR_CARD,
                     fg=COR_ACENTO if ativo else (COR_FRACO if indisponivel else COR_SUAVE),
                     cursor="arrow" if indisponivel else "hand2")
        estado = "readonly"
        for cb in (self.cb_de, self.cb_ate):
            cb.config(state=estado)

    def _definir_modo(self, modo):
        if modo == "modelo" and self.base.prev is None:
            messagebox.showinfo("Modelo indisponível", "Não achei a previsão do modelo (base\\modelo.pkl).\n"
                                "Rode o ATUALIZAR.bat para gerá-la.")
            return
        self.modo_prev = modo
        self._pintar_modo()
        self._recalcular()

    def _ao_mudar_periodo(self, _e=None):
        de, ate = self._periodo_base()
        if de is not None:  # mantém 'de' <= 'até'
            self.var_de.set(nome_mes(de))
            self.var_ate.set(nome_mes(ate))
        self.modo_prev = "periodo"  # escolher um período passa a previsão para o período do 1174
        self._pintar_modo()
        self._recalcular()

    def _recalcular(self):
        if self.cod_atual and self.atual:
            self.mostrar(self.cod_atual, manter=True)
        if self.pagina in ("transf", "saude"):
            self._atualizar_pagina()

    def status(self, txt):
        self.lbl_status.config(text=txt)

    # ---------------- busca
    @staticmethod
    def _separar(texto):
        partes = [p.strip().upper() for p in re.split(r"[\r\n\t;,]+", texto or "")]
        vistos, out = set(), []
        for p in partes:
            p = p.strip(" '\"")
            if p and p not in vistos:
                vistos.add(p)
                out.append(p)
        return out

    def _ao_colar(self, _e):
        try:
            txt = self.clipboard_get()
        except tk.TclError:
            return None
        self.var_busca.set(txt.strip())
        self.after(10, self.consultar)
        return "break"

    def colar_consultar(self):
        try:
            txt = self.clipboard_get()
        except tk.TclError:
            messagebox.showinfo("Área de transferência vazia", "Copie um código primeiro.")
            return
        self.var_busca.set(txt.strip())
        self.consultar()

    def limpar(self):
        self.var_busca.set("")
        self.lista_codigos = []
        self._chips()
        self.ent.focus_set()

    def consultar(self):
        codigos = self._separar(self.var_busca.get())
        if not codigos:
            return
        if len(codigos) > 1:
            self.lista_codigos = codigos
            self.var_busca.set(codigos[0])
        elif codigos[0] not in self.lista_codigos:
            self.lista_codigos = []
        self._chips(codigos[0])
        self.mostrar(codigos[0])

    def _chips(self, ativo=None):
        for w in self.fr_chips.winfo_children():
            w.destroy()
        if len(self.lista_codigos) < 2:
            return
        tk.Label(self.fr_chips, text=f"{len(self.lista_codigos)} códigos colados", bg=COR_CARD,
                 fg=COR_FRACO, font=(FONTE, 9)).pack(side="left", padx=(0, 8), pady=(0, 12))
        for c in self.lista_codigos[:30]:
            sel = c == ativo
            BotaoPlano(self.fr_chips, c, lambda c=c: (self.var_busca.set(c), self._chips(c), self.mostrar(c)),
                       estilo="primario" if sel else "secundario", fonte=(FONTE_FORTE, 9),
                       padx=8, pady=2).pack(side="left", padx=2, pady=(0, 12))
        if len(self.lista_codigos) > 30:
            tk.Label(self.fr_chips, text=f"+{len(self.lista_codigos) - 30}", bg=COR_CARD,
                     fg=COR_SUAVE).pack(side="left", pady=(0, 12))

    # ---------------- cálculo
    def _links(self, titulo, codigos, cor=COR_ACENTO):
        for w in self.fr_rel.winfo_children():
            w.destroy()
        if not codigos:
            return
        tk.Label(self.fr_rel, text=titulo, bg=COR_FUNDO, fg=COR_FRACO,
                 font=(FONTE, 9)).pack(side="left")
        for c in codigos:
            lb = tk.Label(self.fr_rel, text=c, bg=COR_FUNDO, fg=cor, cursor="hand2",
                          font=(FONTE_FORTE, 9))
            lb.pack(side="left", padx=5)
            lb.bind("<Button-1>", lambda e, c=c: (self.var_busca.set(c), self.mostrar(c)))
            _Dica(lb, self.base.descricoes.get(c, ""))

    def _kpi(self, chave, valor, sub="", cor=COR_TEXTO, titulo=None):
        v, s, t = self.kpis[chave]
        v.config(text=valor, fg=cor)
        s.config(text=sub)
        if titulo:
            t.config(text=titulo)

    def _calcular_previsao(self, cod, liquida, perd, est_fil, ref, filiais):
        """Previsão por filial para os N meses a atender.

        - modo "periodo": média mensal (venda líquida + venda perdida) no período-base do 1174 × N.
        - modo "modelo": previsão do modelo do projeto (até 4 meses); sem o item no modelo,
          estimativa simples (50% média últimos 3 meses + 50% mesmo mês do ano anterior).
        Cada filial recebe: f (lista mensal), meia (meia-largura da faixa por mês), meia_tot (faixa do total).
        """
        n_at = self._meses_atender()
        per_ref = pd.Period(ref, "M")
        hoje = pd.Timestamp(date.today())
        out = {}
        de, ate = self._periodo_base()
        cod_modelo = cod
        if self.modo_prev == "periodo" and de is not None:
            meses = [per_ref + k for k in range(1, n_at + 1)]
            meses_base = list(pd.period_range(de, ate, freq="M"))
            chave = [liquida["filial"], liquida["data"].dt.to_period("M")]
            vm = liquida.groupby(chave)["qtd"].sum() if not liquida.empty else pd.Series(dtype=float)
            pm = perd.groupby([perd["filial"], perd["data"].dt.to_period("M")])["qtd"].sum() \
                if not perd.empty else pd.Series(dtype=float)
            for fl in filiais:
                vend = np.array([float(vm.get((fl, m), 0.0)) for m in meses_base])
                pdd = np.array([float(pm.get((fl, m), 0.0)) for m in meses_base])
                serie = np.maximum(vend + pdd, 0.0)
                media, dp = float(serie.mean()), float(serie.std())
                est = est_fil.get(fl, 0.0)
                total = media * n_at
                cob = est / (media / 30) if media > 0 else None
                comprar = max(0, math.ceil(total - est - 1e-9)) if media > 0 else 0
                if media <= 0:
                    st = "Sem giro"
                elif est <= 0:
                    st = "Ruptura"
                elif comprar > 0:
                    st = "Comprar"
                else:
                    st = "OK"
                out[fl] = dict(f=[media] * n_at, meia=[dp] * n_at, meia_tot=dp * math.sqrt(n_at),
                               classe="", status=st, sug=comprar, cob=cob,
                               falta=(hoje + pd.Timedelta(days=int(cob))) if cob is not None and cob < 3650 else None,
                               media=media, base_vend=float(vend.sum()), base_perd=float(pdd.sum()))
            return dict(meses=meses, filiais=out, fonte="periodo", cod_modelo=cod, base=(de, ate),
                        n_base=len(meses_base))

        pv, cod_modelo = self.base.previsao(cod)
        if pv is not None and not pv.empty:
            meses = [pd.Period(m, "M") for m in self.base.prev_meta["meses"]][:n_at]
            for r in pv.itertuples(index=False):
                f = [float(getattr(r, f"f{h}")) for h in range(len(meses))]
                fx = float(r.faixa)
                out[int(r.filial)] = dict(
                    f=f, meia=[q * fx for q in f], meia_tot=sum(f) * fx, classe=str(r.classe),
                    status=str(getattr(r, "status", "") or ""),
                    falta=getattr(r, "data_prevista_falta", None),
                    sug=getattr(r, "sugestao_compra", None),
                    cob=getattr(r, "cobertura_dias", None))
            fonte = "modelo"
        else:
            meses = [per_ref + k for k in range(1, n_at + 1)]
            serie = liquida.groupby([liquida["filial"], liquida["data"].dt.to_period("M")])["qtd"].sum() \
                if not liquida.empty else pd.Series(dtype=float)
            for fl in filiais:
                base3 = sum(serie.get((fl, per_ref - k), 0.0) for k in range(3)) / 3
                f = [max(0.0, 0.5 * base3 + 0.5 * serie.get((fl, m - 12), 0.0)) for m in meses]
                out[fl] = dict(f=f, meia=[q * 0.5 for q in f], meia_tot=sum(f) * 0.5,
                               classe="estimativa simples", status="", falta=None, sug=None, cob=None)
            fonte = "simples"
        return dict(meses=meses, filiais=out, fonte=fonte, cod_modelo=cod_modelo)

    def mostrar(self, cod, manter=False):
        aba_atual = self.nb.index(self.nb.select())
        if not manter:
            self.filial_sel = None
        self.cod_atual = cod
        v, p, e = self.base.linhas(cod)
        for tv in (self.tv_res, self.tv_vend, self.tv_perd, self.tv_cli, self.tv_prev):
            tv.delete(*tv.get_children())

        if v.empty and p.empty and e.empty:
            self.atual = None
            sug = self.base.sugestoes(cod)
            self.lbl_cod.config(text=cod, fg=COR_VERMELHO)
            self.lbl_desc.config(text="código não encontrado em nenhuma base")
            for k in self.kpis:
                self._kpi(k, "—")
            self._links("Você quis dizer:" if sug else "Nenhum código ou descrição parecida.", sug)
            self.status(f"'{cod}' não encontrado.")
            self.lbl_filtro.config(text="")
            self._graficos_vazios("Código não encontrado.")
            self.nb.select(self.ABA_VISAO)
            return

        hoje = pd.Timestamp(date.today())
        ref = self.base.vendas["data"].max() if not self.base.vendas.empty else hoje
        d30, d90, d365 = (ref - pd.Timedelta(days=n) for n in (30, 90, 365))
        desc = self.base.descricoes.get(cod, "")
        marca = str(v["marca"].iloc[-1]) if not v.empty else ""
        self.lbl_cod.config(text=cod, fg=COR_TEXTO)
        self.lbl_desc.config(text=desc + (f"   ·   marca {marca}" if marca and marca != "nan" else ""))

        tipo = v["tipo_lanc"].astype(str)
        liquida = v[tipo.isin(["Vendas", "Devolução de vendas"])]
        vendas_pos = v[(tipo == "Vendas") & (v["qtd"] > 0)]
        filiais = sorted(set(NOMES_FILIAIS) | set(v["filial"]) | set(p["filial"]) | set(e["filial"]))
        est_fil = e.groupby("filial")["estoque"].sum().to_dict() if not e.empty else {}
        prev = self._calcular_previsao(cod, liquida, p, est_fil, ref, filiais)

        def custo_medio(vv):
            """(custo médio unitário, preço médio unitário, origem): últimos 12 meses; senão todo o 1174."""
            for df, origem in ((vv[vv["data"] > d365], "12 meses"), (vv, "desde " + fmt_data(vv["data"].min())
                                                                   if not vv.empty else "")):
                q = float(df["qtd"].sum())
                c = float(df["custo"].sum())
                if q > 0 and c > 0:
                    return c / q, float(df["valor"].sum()) / q, origem
            return None, None, ""
        cm_rede, pm_rede, cm_origem = custo_medio(vendas_pos)
        if cm_rede is None and not e.empty and (e["custo_unit"] > 0).any():  # sem venda: custo do estoque
            cm_rede, cm_origem = float(e.loc[e["custo_unit"] > 0, "custo_unit"].mean()), "custo do estoque (1211)"
        self.atual = dict(v=v, p=p, e=e, liquida=liquida, vendas_pos=vendas_pos, ref=ref, hoje=hoje,
                          filiais=filiais, prev=prev, custo=cm_rede)

        # ---- resumo por filial
        tot = dict(estoque=0.0, v30=0.0, v90=0.0, v12=0.0, vtot=0.0, perd=0.0, prev=0.0)
        resumo = {}
        for f in filiais:
            vf = vendas_pos[vendas_pos["filial"] == f]
            lf = liquida[liquida["filial"] == f]
            pf = p[p["filial"] == f]
            est = float(e.loc[e["filial"] == f, "estoque"].sum())
            ult = vf.iloc[-1] if not vf.empty else None
            r = dict(est=est, ult=ult, custo=custo_medio(vf)[0],
                     v30=lf.loc[lf["data"] > d30, "qtd"].sum(), v90=lf.loc[lf["data"] > d90, "qtd"].sum(),
                     v12=lf.loc[lf["data"] > d365, "qtd"].sum(), vtot=lf["qtd"].sum(),
                     n12=int((vf["data"] > d365).sum()),
                     ncli=vf.loc[vf["data"] > d365, "cliente"].astype(str).replace("", np.nan).nunique(),
                     perd=pf["qtd"].sum(), pu=pf.sort_values("data").iloc[-1] if not pf.empty else None,
                     prev=prev["filiais"].get(f, {}).get("f", [0.0])[0])
            resumo[f] = r
            for k in ("v30", "v90", "v12", "vtot", "perd", "prev"):
                tot[k] += r[k]
            tot["estoque"] += est
        self.atual["resumo"] = resumo

        for i, f in enumerate(filiais):
            r = resumo[f]
            ult, pu = r["ult"], r["pu"]
            unit = (ult["valor"] / ult["qtd"]) if ult is not None and ult["qtd"] else None
            valores = (nome_filial(f), fmt_qtd(r["est"]),
                       fmt_data(ult["data"]) if ult is not None else "—",
                       fmt_num((hoje - ult["data"]).days, 0) if ult is not None else "",
                       fmt_qtd(ult["qtd"]) if ult is not None else "",
                       str(ult["cliente"]) if ult is not None else "",
                       fmt_moeda(unit) if unit else "", fmt_moeda(r["custo"]) if r["custo"] else "",
                       fmt_qtd(r["v30"]), fmt_qtd(r["v90"]), fmt_qtd(r["v12"]), fmt_qtd(r["vtot"]),
                       str(r["ncli"]), fmt_prev(r["prev"]),
                       fmt_data(pu["data"]) if pu is not None else "—",
                       fmt_qtd(r["perd"]) if r["perd"] else "")
            if ult is None and pu is None and r["est"] == 0:
                tag = "parado"
            elif r["est"] <= 0 and (r["n12"] > 0 or r["perd"] > 0):
                tag = "semestoque"
            else:
                tag = "par" if i % 2 else "impar"
            self.tv_res.insert("", "end", iid=f"f{f}", values=valores, tags=(tag,))
        ult_geral = vendas_pos.iloc[-1] if not vendas_pos.empty else None
        cli12 = vendas_pos.loc[vendas_pos["data"] > d365, "cliente"].astype(str).replace("", np.nan)
        self.tv_res.insert("", "end", iid="total", tags=("total",), values=(
            "TOTAL REDE", fmt_qtd(tot["estoque"]),
            fmt_data(ult_geral["data"]) if ult_geral is not None else "—",
            fmt_num((hoje - ult_geral["data"]).days, 0) if ult_geral is not None else "",
            "", "", "", fmt_moeda(cm_rede) if cm_rede else "",
            fmt_qtd(tot["v30"]), fmt_qtd(tot["v90"]), fmt_qtd(tot["v12"]),
            fmt_qtd(tot["vtot"]), str(cli12.nunique()), fmt_prev(tot["prev"]),
            fmt_data(p["data"].max()) if not p.empty else "—",
            fmt_qtd(tot["perd"]) if tot["perd"] else ""))

        # ---- KPIs
        n_fil_est = int((e.groupby("filial")["estoque"].sum() > 0).sum()) if not e.empty else 0
        n_fil_12 = sum(1 for r in resumo.values() if r["n12"] > 0)
        self._kpi("estoque", fmt_qtd(tot["estoque"]),
                  f"em {n_fil_est} filial(is)" if n_fil_est else "sem estoque em nenhuma filial",
                  COR_TEXTO if tot["estoque"] > 0 else COR_VERMELHO)
        if ult_geral is not None:
            self._kpi("ult", fmt_data(ult_geral["data"]),
                      f"{sigla_filial(ult_geral['filial'])} · {fmt_qtd(ult_geral['qtd'])} un · "
                      f"há {(hoje - ult_geral['data']).days} dias")
        else:
            self._kpi("ult", "nunca", "sem venda desde 2024", COR_SUAVE)
        self._kpi("v12", fmt_qtd(tot["v12"]), f"90 dias: {fmt_qtd(tot['v90'])} · {n_fil_12} filial(is)")
        if cm_rede:
            margem = f" · margem {fmt_num((1 - cm_rede / pm_rede) * 100, 0)}%" if pm_rede else ""
            self._kpi("custo", fmt_moeda(cm_rede), f"por unidade, {cm_origem}{margem}")
        else:
            self._kpi("custo", "—", "sem custo registrado", COR_SUAVE)
        top_cli = vendas_pos.loc[vendas_pos["data"] > d365].groupby(
            vendas_pos["cliente"].astype(str), observed=True)["qtd"].sum().sort_values()
        top_cli = top_cli[top_cli.index != ""]
        self._kpi("cli", str(cli12.nunique()),
                  f"maior: {top_cli.index[-1][:22]}" if len(top_cli) else "nenhum cliente")
        self._kpi("perd", fmt_qtd(tot["perd"]) if not p.empty else "0",
                  f"{len(p)} registro(s) desde {fmt_data(self.base.perdidas['data'].min())}"
                  if not self.base.perdidas.empty else "",
                  COR_LARANJA if not p.empty else COR_TEXTO)
        valor_compra = ""
        nprev = len(prev["meses"])
        tot_n = sum(sum(x["f"]) for x in prev["filiais"].values())
        compra = sum((x.get("sug") or 0) for x in prev["filiais"].values()
                     if x.get("sug") is not None and not pd.isna(x.get("sug")))
        if compra and cm_rede:
            valor_compra = f" ({fmt_moeda(compra * cm_rede)})"
        if prev["fonte"] == "periodo":
            sub = f"base {nome_mes(prev['base'][0])}–{nome_mes(prev['base'][1])} · comprar {fmt_qtd(compra)} un{valor_compra}"
        else:
            sub = ("modelo" if prev["fonte"] == "modelo" else "estimativa simples") + \
                  (f" · comprar {fmt_qtd(compra)}{valor_compra}" if compra else "")
        self._kpi("prev", fmt_prev(tot_n), sub, COR_ACENTO,
                  titulo=f"Previsão {nprev} {'mês' if nprev == 1 else 'meses'} (rede)")

        rel = self.base.relacionados(cod)
        self._links("Relacionados:" if rel else "", rel)
        self._preencher_previsao()
        if manter and self.filial_sel is not None and self.tv_res.exists(f"f{self.filial_sel}"):
            self.tv_res.selection_set(f"f{self.filial_sel}")
        elif manter:
            self.filial_sel = None
        self._atualizar_filtro()
        self.nb.select(aba_atual if manter else self.ABA_VISAO)
        self.status(f"{cod} · {len(v)} lançamento(s) no 1174 · {len(p)} venda(s) perdida(s) · "
                    f"estoque em {n_fil_est} filial(is)")

    # ---------------- filtro por filial
    def _atualizar_filtro(self):
        if not self.atual:
            return
        self.lbl_filtro.config(text=f"Filial: {nome_filial(self.filial_sel)}" if self.filial_sel
                               else "Filial: todas")
        self._preencher_detalhes()
        self._preencher_clientes()
        self._desenhar_graficos()

    def _ao_selecionar_filial(self, _e):
        sel = self.tv_res.selection()
        novo = None if not sel or sel[0] == "total" else int(sel[0][1:])
        if novo != self.filial_sel:
            self.filial_sel = novo
            self._atualizar_filtro()

    def _selecionar_filial(self, f):
        if f == self.filial_sel:
            self._todas_filiais()
            return
        self.tv_res.selection_set(f"f{f}")
        self.tv_res.see(f"f{f}")

    def _todas_filiais(self):
        self.tv_res.selection_remove(self.tv_res.selection())
        self.filial_sel = None
        self._atualizar_filtro()

    def _filtrar(self, df):
        return df[df["filial"] == self.filial_sel] if self.filial_sel is not None else df

    # ---------------- tabelas
    def _preencher_detalhes(self):
        a = self.atual
        v, p = self._filtrar(a["v"]), self._filtrar(a["p"])
        for tv in (self.tv_vend, self.tv_perd):
            tv.delete(*tv.get_children())
        v = v.iloc[::-1].head(1000)
        for i, r in enumerate(v.itertuples(index=False)):
            unit = r.valor / r.qtd if r.qtd else None
            cu = r.custo / r.qtd if r.qtd and r.custo else None
            tags = ["par" if i % 2 else "impar"] + (["negativo"] if r.qtd < 0 else [])
            self.tv_vend.insert("", "end", values=(
                fmt_data(r.data), nome_filial(r.filial), str(r.cliente), fmt_qtd(r.qtd),
                fmt_moeda(unit) if unit else "", fmt_moeda(cu) if cu else "", fmt_moeda(r.valor),
                str(r.tipo_lanc),
                str(r.vendedor), str(r.cidade), str(r.operacao), str(r.nf or "")), tags=tags)
        p = p.sort_values("data", ascending=False)
        for i, r in enumerate(p.itertuples(index=False)):
            self.tv_perd.insert("", "end", values=(
                fmt_data(r.data), nome_filial(r.filial), fmt_qtd(r.qtd), fmt_moeda(r.preco),
                fmt_moeda(r.valor), str(r.cliente), str(r.vendedor), str(r.origem), str(r.motivo)),
                tags=("par" if i % 2 else "impar",))
        self._aba_txt(self.ABA_VENDAS, f"Últimas vendas  {len(v)}")
        self._aba_txt(self.ABA_PERDIDAS, f"Vendas perdidas  {len(p)}")

    def _preencher_clientes(self):
        a = self.atual
        tv = self.tv_cli
        tv.delete(*tv.get_children())
        vp = self._filtrar(a["vendas_pos"])
        n_cli_total = 0
        for f in a["filiais"]:
            vf = vp[vp["filial"] == f]
            if vf.empty:
                continue
            vf = vf.assign(cli=vf["cliente"].astype(str).replace("", "(sem nome)"))
            g = vf.groupby("cli", sort=False)
            tab = pd.DataFrame({
                "conta": g["cliente_cod"].last(), "compras": g.size(), "qtd": g["qtd"].sum(),
                "valor": g["valor"].sum(), "ult": g["data"].max(),
                "preco": g.apply(lambda d: d["valor"].iloc[-1] / d["qtd"].iloc[-1], include_groups=False),
                "vendedor": g["vendedor"].last().astype(str), "cidade": g["cidade"].last().astype(str),
            }).sort_values(["ult", "qtd"], ascending=False)
            n_cli_total += len(tab)
            pai = tv.insert("", "end", iid=f"g{f}", open=self.filial_sel is not None or len(tab) <= 15,
                            text=f"{nome_filial(f)}   ·   {len(tab)} cliente(s)", tags=("grupo",),
                            values=("", str(int(tab["compras"].sum())), fmt_qtd(tab["qtd"].sum()),
                                    fmt_moeda(tab["valor"].sum()), fmt_data(tab["ult"].max()),
                                    fmt_num((a["hoje"] - tab["ult"].max()).days, 0), "", "", ""))
            for i, (nome, r) in enumerate(tab.iterrows()):
                tv.insert(pai, "end", text=nome, tags=("par" if i % 2 else "impar",), values=(
                    str(r["conta"] or ""), str(r["compras"]), fmt_qtd(r["qtd"]), fmt_moeda(r["valor"]),
                    fmt_data(r["ult"]), fmt_num((a["hoje"] - r["ult"]).days, 0),
                    fmt_moeda(r["preco"]), r["vendedor"], r["cidade"]))
        self._aba_txt(self.ABA_CLIENTES, f"Clientes por filial  {n_cli_total}")

    def _expandir_cli(self, abrir: bool):
        for k in self.tv_cli.get_children(""):
            self.tv_cli.item(k, open=abrir)

    def _definir_colunas(self, tv, cols):
        tv["columns"] = [c[0] for c in cols]
        tv["displaycolumns"] = "#all"
        for cid, titulo, larg, anc in cols:
            tv.heading(cid, text=titulo, anchor=anc, command=lambda c=cid, t=tv: self._ordenar(t, c))
            tv.column(cid, width=larg, anchor=anc, stretch=True)

    def _preencher_previsao(self):
        a = self.atual
        pv = a["prev"]
        tv = self.tv_prev
        tv.delete(*tv.get_children())
        n = len(pv["meses"])
        horiz = f"{n} {'mês' if n == 1 else 'meses'}"
        ini_fim = f"{nome_mes(pv['meses'][0])} a {nome_mes(pv['meses'][-1])}" if n > 1 else nome_mes(pv["meses"][0])
        periodo = pv["fonte"] == "periodo"
        if periodo:
            de, ate = pv["base"]
            self._definir_colunas(tv, [
                ("filial", "Filial", 80, "w"), ("estoque", "Estoque", 70, "e"),
                ("vend", f"Vendido {nome_mes(de)}–{nome_mes(ate)}", 165, "e"),
                ("perd", "Perdido no período", 140, "e"), ("media", "Média/mês", 85, "e"),
                ("tot", f"Previsão {horiz}", 120, "e"), ("faixa", "Faixa provável", 120, "center"),
                ("cob", "Cobertura (dias)", 110, "e"), ("falta", "Falta prevista", 105, "center"),
                ("status", "Situação", 100, "w"), ("sug", "Comprar", 80, "e"),
                ("vsug", "Valor (custo médio)", 145, "e")])
            nota = (f"Previsão pelo PERÍODO DO 1174: média mensal de {nome_mes(de)} a {nome_mes(ate)} "
                    f"({pv['n_base']} meses) × {horiz} a atender ({ini_fim}). Demanda = venda líquida + venda "
                    f"perdida (vendas perdidas existem desde {fmt_data(self.base.perdidas['data'].min())}).  "
                    "Comprar = previsão − estoque atual.  Faixa provável = variação mês a mês no período.")
        else:
            self._definir_colunas(tv, [
                ("filial", "Filial", 80, "w"), ("estoque", "Estoque", 70, "e"),
                ("perfil", "Perfil de demanda", 125, "w")] +
                [(f"m{h}", nome_mes(m), 75, "e") for h, m in enumerate(pv["meses"])] + [
                ("tot", f"Total {horiz}", 100, "e"), ("faixa", "Faixa provável", 120, "center"),
                ("cob", "Cobertura (dias)", 110, "e"), ("falta", "Falta prevista", 105, "center"),
                ("status", "Situação", 110, "w"), ("sug", "Sugestão compra", 115, "e"),
                ("vsug", "Valor (custo médio)", 145, "e")])
            if pv["fonte"] == "modelo":
                meta = self.base.prev_meta
                nota = (f"Previsão do MODELO do projeto (gerado em {fmt_data(meta.get('gerado'))}, vendas até "
                        f"{fmt_data(meta.get('fim_vendas'))}), {horiz} ({ini_fim}; o modelo vai até 4 meses). "
                        "Situação e sugestão de compra vêm do modelo (lead time + segurança).")
                if pv["cod_modelo"] != self.cod_atual:
                    nota += f"\nEste código foi substituído: a previsão está consolidada no código {pv['cod_modelo']}."
            else:
                nota = ("Esta peça não está no modelo — ESTIMATIVA SIMPLES: metade da média dos últimos 3 meses + "
                        "metade do mesmo mês do ano anterior (faixa ±50%).")
        cm = a.get("custo")

        def vcompra(q):
            return fmt_moeda(q * cm) if q and cm else ""
        if cm:
            nota += f"  Valor = quantidade a comprar × custo médio vendido ({fmt_moeda(cm)}/un)."
        self.lbl_prev.config(text=nota)
        tot_m = np.zeros(n)
        var_tot = tot_sug = tot_vend = tot_perd = tot_media = 0.0
        for i, f in enumerate(a["filiais"]):
            x = pv["filiais"].get(f) or dict(f=[0.0] * n, meia=[0.0] * n, meia_tot=0.0, classe="sem histórico",
                                             status="", falta=None, sug=None, cob=None)
            est = a["resumo"][f]["est"]
            fs = (list(x["f"]) + [0.0] * n)[:n]
            tn = sum(fs)
            tot_m += np.array(fs)
            mn, mx = max(0.0, tn - x["meia_tot"]), tn + x["meia_tot"]
            var_tot += x["meia_tot"] ** 2
            sug = x["sug"] if x["sug"] is not None and not pd.isna(x["sug"]) else None
            tot_sug += sug or 0
            cob = x["cob"]
            st = x["status"] if x["status"] and x["status"] != "nan" else ""
            stu = st.upper()
            tag = ("st_rup" if "RUPT" in stu else "st_crit" if "CRÍT" in stu or "CRIT" in stu
                   else "st_aten" if ("ATEN" in stu or "COMPRAR" in stu) else ("par" if i % 2 else "impar"))
            if tn < 0.05 and est == 0:
                tag = "parado"
            txt_faixa = f"{fmt_prev(mn)} a {fmt_prev(mx)}" if tn >= 0.05 else ""
            txt_cob = (fmt_num(cob, 0) if cob is not None and not pd.isna(cob) and cob < 9999 else
                       ("—" if tn < 0.05 else "sem limite"))
            txt_falta = fmt_data(x["falta"]) if x["falta"] is not None and not pd.isna(x["falta"]) else ""
            if periodo:
                tot_vend += x["base_vend"]
                tot_perd += x["base_perd"]
                tot_media += x["media"]
                vals = (nome_filial(f), fmt_qtd(est), fmt_qtd(x["base_vend"]),
                        fmt_qtd(x["base_perd"]) if x["base_perd"] else "", fmt_prev(x["media"]), fmt_prev(tn),
                        txt_faixa, txt_cob, txt_falta, st, fmt_qtd(sug) if sug else "", vcompra(sug))
            else:
                vals = (nome_filial(f), fmt_qtd(est), x["classe"].capitalize() if x["classe"] else "",
                        *[fmt_prev(q) for q in fs], fmt_prev(tn), txt_faixa, txt_cob, txt_falta,
                        st.capitalize(), fmt_qtd(sug) if sug else "", vcompra(sug))
            tv.insert("", "end", iid=f"p{f}", tags=(tag,), values=vals)
        est_tot = fmt_qtd(sum(a["resumo"][f]["est"] for f in a["filiais"]))
        faixa_tot = f"{fmt_prev(max(0.0, tot_m.sum() - math.sqrt(var_tot)))} a {fmt_prev(tot_m.sum() + math.sqrt(var_tot))}"
        if periodo:
            vals = ("TOTAL REDE", est_tot, fmt_qtd(tot_vend), fmt_qtd(tot_perd) if tot_perd else "",
                    fmt_prev(tot_media), fmt_prev(tot_m.sum()), faixa_tot, "", "", "", fmt_qtd(tot_sug) if tot_sug else "",
                    vcompra(tot_sug))
        else:
            vals = ("TOTAL REDE", est_tot, "", *[fmt_prev(q) for q in tot_m], fmt_prev(tot_m.sum()), faixa_tot,
                    "", "", "", fmt_qtd(tot_sug) if tot_sug else "", vcompra(tot_sug))
        tv.insert("", "end", iid="total", tags=("total",), values=vals)

    # ---------------- gráficos
    def _graficos_vazios(self, texto):
        for g in (self.g_mensal, self.g_filiais, self.g_cli):
            g.definir(lambda c, w, h, t=texto: c.vazio(t))

    def _desenhar_graficos(self):
        a = self.atual
        fsel = self.filial_sel
        liq, p = self._filtrar(a["liquida"]), self._filtrar(a["p"])
        per_ref = pd.Period(a["ref"], "M")
        hist = [per_ref - k for k in range(MESES_HIST - 1, -1, -1)]
        vm = liq.groupby(liq["data"].dt.to_period("M"))["qtd"].sum() if not liq.empty else pd.Series(dtype=float)
        pm = p.groupby(p["data"].dt.to_period("M"))["qtd"].sum() if not p.empty else pd.Series(dtype=float)
        vend = [max(0.0, float(vm.get(m, 0.0))) for m in hist]
        perd = [float(pm.get(m, 0.0)) for m in hist]
        pv = a["prev"]
        fil = [fsel] if fsel is not None else list(pv["filiais"])
        npv = len(pv["meses"])
        prev = [sum(pv["filiais"][f]["f"][h] for f in fil if f in pv["filiais"]) for h in range(npv)]
        # faixa: erros das filiais se compensam em parte -> raiz da soma dos quadrados
        meia = [math.sqrt(sum(pv["filiais"][f]["meia"][h] ** 2 for f in fil if f in pv["filiais"]))
                for h in range(npv)]
        pmin = [max(0.0, q - m) for q, m in zip(prev, meia)]
        pmax = [q + m for q, m in zip(prev, meia)]
        titulo = f"Vendas mês a mês — {nome_filial(fsel) if fsel else 'Rede (todas as filiais)'}"
        dados = dict(titulo=titulo, hist=[nome_mes(m) for m in hist], vend=vend, perd=perd,
                     prev_rot=[nome_mes(m) for m in pv["meses"]], prev=prev, pmin=pmin, pmax=pmax,
                     fonte=pv["fonte"],
                     base=[i for i, m in enumerate(hist) if pv["fonte"] == "periodo"
                           and pv["base"][0] <= m <= pv["base"][1]])
        self.g_mensal.definir(lambda c, w, h: self._g_mensal(c, w, h, dados))

        linhas = [(f, a["resumo"][f]["v12"], a["resumo"][f]["est"], sum(pv["filiais"].get(f, {"f": [0]})["f"]))
                  for f in a["filiais"]]
        self.g_filiais.definir(lambda c, w, h: self._g_filiais(c, w, h, linhas, npv))

        vp = self._filtrar(a["vendas_pos"])
        if vp.empty:
            top = []
        else:
            chave = vp["cliente"].astype(str)
            if fsel is None:
                chave = chave + " (" + vp["filial"].map(sigla_filial) + ")"
            s = vp.groupby(chave)["qtd"].sum().sort_values(ascending=False)
            top = [(k, float(q)) for k, q in s.head(12).items() if not k.startswith(" (")]
        tit = f"Quantidade comprada desde {fmt_data(self.base.vendas['data'].min())}" + \
              (f" — {sigla_filial(fsel)}" if fsel else "")
        self.g_cli.definir(lambda c, w, h: self._g_top(c, w, h, tit, top))

    def _g_mensal(self, c: Grafico, W, H, d):
        c.titulo(d["titulo"], "Unidades por mês · " + {"simples": "previsão por estimativa simples",
                                                        "periodo": "previsão pela média do período-base"}
                 .get(d["fonte"], "previsão do modelo"))
        c.legenda([(G_VENDIDO, "vendido", None), (G_PERDIDO, "venda perdida", None),
                   (G_PREV, "previsão", G_PREV), (G_FAIXA, "faixa provável", None)], y=22)
        L, R, T, B = 50, 20, 84, H - 52
        if B - T < 40:
            return
        nh, npv = len(d["hist"]), len(d["prev"])
        n = nh + npv
        tops = [a + b for a, b in zip(d["vend"], d["perd"])] + d["pmax"]
        topo, ticks = escala(max(tops + [1]))

        def y(v):
            return B - (v / topo) * (B - T)
        for t in ticks:
            c.create_line(L, y(t), W - R, y(t), fill=COR_GRADE)
            c.create_text(L - 8, y(t), text=fmt_prev(t), anchor="e", font=(FONTE, 8), fill=COR_FRACO)
        slot = (W - L - R) / n
        bw = max(3, min(28, slot * 0.58))
        passo_rot = max(1, math.ceil(46 / slot))
        mostrar_val = slot >= 26
        xsep = L + nh * slot
        if d.get("base"):  # período-base escolhido, destacado no histórico
            xb0, xb1 = L + d["base"][0] * slot, L + (d["base"][-1] + 1) * slot
            c.create_rectangle(xb0, T - 14, xb1, B, fill="#FFF9EE", outline="")
            c.create_text(xb0 + 6, T - 12, text="período-base", anchor="nw", font=(FONTE, 8),
                          fill=COR_LARANJA)
        c.create_rectangle(xsep, T - 14, W - R, B, fill="#F6F8FE", outline="")
        c.create_line(xsep, T - 14, xsep, B, fill="#C9CED6", dash=(3, 3))
        c.create_text(xsep + 6, T - 12, text="previsão", anchor="nw", font=(FONTE, 8), fill=G_FAIXA)
        for i in range(n):
            x0 = L + i * slot + (slot - bw) / 2
            x1 = x0 + bw
            tag = f"b{i}"
            if i < nh:
                vv, pp = d["vend"][i], d["perd"][i]
                c.create_rectangle(x0, T - 6, x1, B, fill="", outline="", tags=tag)  # área de hover
                if vv > 0:
                    c.create_rectangle(x0, y(vv), x1, B, fill=G_VENDIDO, outline="", tags=tag)
                if pp > 0:
                    c.create_rectangle(x0, y(vv + pp), x1, y(vv), fill=G_PERDIDO, outline="", tags=tag)
                if mostrar_val and vv + pp > 0:
                    c.create_text((x0 + x1) / 2, y(vv + pp) - 8, text=fmt_qtd(vv + pp), font=(FONTE, 8),
                                  fill=COR_SUAVE, tags=tag)
                txt = f"{d['hist'][i]}: vendido {fmt_qtd(vv)} un" + (f" · perdido {fmt_qtd(pp)} un" if pp else "")
                rot = d["hist"][i]
            else:
                h = i - nh
                q, mn, mx = d["prev"][h], d["pmin"][h], d["pmax"][h]
                c.create_rectangle(x0, T - 6, x1, B, fill="", outline="", tags=tag)
                c.create_rectangle(x0, y(q), x1, B, fill=G_PREV, outline="", tags=tag)
                if mx > mn:
                    xm = (x0 + x1) / 2
                    c.create_line(xm, y(mn), xm, y(mx), fill=G_FAIXA, width=1, tags=tag)
                    c.create_line(xm - bw / 5, y(mx), xm + bw / 5, y(mx), fill=G_FAIXA, width=1, tags=tag)
                    c.create_line(xm - bw / 5, y(mn), xm + bw / 5, y(mn), fill=G_FAIXA, width=1, tags=tag)
                if mostrar_val:
                    c.create_text((x0 + x1) / 2, y(max(q, mx)) - 8, text=fmt_prev(q), font=(FONTE_FORTE, 8),
                                  fill=G_FAIXA, tags=tag)
                txt = f"{d['prev_rot'][h]} (previsão): {fmt_prev(q)} un · faixa {fmt_prev(mn)} a {fmt_prev(mx)}"
                rot = d["prev_rot"][h]
            c.dica(tag, txt)
            if (nh - 1 - i) % passo_rot == 0:
                c.create_text((x0 + x1) / 2, B + 14, text=rot, font=(FONTE, 8),
                              fill=G_FAIXA if i >= nh else COR_FRACO)
        c.create_line(L, B, W - R, B, fill=COR_BORDA)
        soma12 = sum(d["vend"][-12:])
        c.create_text(L, H - 14, anchor="w", font=(FONTE, 9), fill=COR_SUAVE,
                      text=f"Últimos 12 meses: {fmt_qtd(soma12)} un vendidas · média {fmt_prev(soma12 / 12)}/mês"
                           f"   ·   Previsão {len(d['prev'])} {'mês' if len(d['prev']) == 1 else 'meses'}: "
                           f"{fmt_prev(sum(d['prev']))} un")

    def _g_filiais(self, c: Grafico, W, H, linhas, npv=4):
        c.titulo("Por filial", f"Vendido 12 meses · estoque · previsão {npv} {'mês' if npv == 1 else 'meses'}"
                               " — clique para filtrar")
        L, R, T, B = 70, 150, 84, H - 12
        n = len(linhas)
        if n == 0 or B - T < 40:
            return
        rh = (B - T) / n
        vmax = max([max(v, e, p) for _, v, e, p in linhas] + [1])
        bh = max(3, min(9, (rh - 8) / 3))
        cols = [(W - 112, "vend.", G_VENDIDO), (W - 64, "est.", G_ESTOQUE), (W - 16, "prev.", G_FAIXA)]
        for x, t, cor in cols:
            tx =c.create_text(x, T - 12, text=t, anchor="e", font=(FONTE, 8), fill=COR_SUAVE)
            bx = c.bbox(tx)[0]
            c.create_oval(bx - 10, T - 15, bx - 4, T - 9, fill=cor, outline="")

        def xw(v):
            return L + (max(v, 0) / vmax) * (W - L - R)
        for k, (f, v12, est, p4) in enumerate(linhas):
            y0 = T + k * rh
            tag = f"row{f}"
            sel = f == self.filial_sel
            c.create_rectangle(0, y0, W, y0 + rh, fill=COR_ACENTO_CLARO if sel else COR_CARD,
                               outline="", tags=tag)
            if k:
                c.create_line(16, y0, W - 16, y0, fill=COR_GRADE)
            ym = y0 + rh / 2
            c.create_text(L - 8, ym, text=f"{f:02d} {sigla_filial(f)}", anchor="e",
                          font=(FONTE_FORTE if sel else FONTE, 9), fill=COR_TEXTO, tags=tag)
            for j, (val, cor, borda) in enumerate([(v12, G_VENDIDO, None), (est, G_ESTOQUE, None),
                                                   (p4, G_PREV, G_VENDIDO)]):
                yb = ym - 1.5 * bh - 1 + j * (bh + 1)
                if val > 0:
                    c.create_rectangle(L, yb, max(L + 2, xw(val)), yb + bh, fill=cor, outline="", tags=tag)
            sem_est = est <= 0 and (v12 > 0 or p4 >= 0.5)
            for (x, _, cor), val, txt in zip(cols, (v12, est, p4),
                                             (fmt_qtd(v12), fmt_qtd(est), fmt_prev(p4))):
                c.create_text(x, ym, text=txt, anchor="e", tags=tag,
                              font=(FONTE_FORTE if (sem_est and val is est) else FONTE, 9),
                              fill=COR_VERMELHO if (sem_est and val is est) else COR_TEXTO)
            c.dica(tag, f"{nome_filial(f)} — vendido 12m: {fmt_qtd(v12)} · estoque: {fmt_qtd(est)} · "
                        f"previsão {npv}m: {fmt_prev(p4)}" + ("  ⚠ sem estoque" if sem_est else "") +
                        "\n(clique para filtrar por esta filial)")
            c.tag_bind(tag, "<Button-1>", lambda e, f=f: self._selecionar_filial(f))
            c.tag_bind(tag, "<Enter>", lambda e: c.config(cursor="hand2"), add="+")
            c.tag_bind(tag, "<Leave>", lambda e: c.config(cursor=""), add="+")

    def _g_top(self, c: Grafico, W, H, titulo, top):
        c.titulo("Maiores compradores", titulo)
        if not top:
            c.vazio("Nenhuma venda para mostrar.")
            return
        L, R, T, B = 20, 48, 64, H - 12
        rh = min(34, (B - T) / len(top))
        vmax = max(q for _, q in top)
        xb = L + max(150, min(260, (W - L - R) * 0.5))  # coluna dos nomes | barras
        larg_nome = max(14, int((xb - L - 10) / 7))  # ~7 px por caractere em 9 pt
        for k, (nome, q) in enumerate(top):
            ym = T + k * rh + rh / 2
            if k:
                c.create_line(L, ym - rh / 2, W - 16, ym - rh / 2, fill=COR_GRADE)
            nm = nome.title()
            c.create_text(L, ym, text=nm if len(nm) <= larg_nome else nm[:larg_nome - 1] + "…",
                          anchor="w", font=(FONTE, 9), fill=COR_TEXTO)
            x1 = max(xb + 2, xb + (q / vmax) * (W - xb - R))
            bh = min(10, rh * 0.4)
            c.create_rectangle(xb, ym - bh / 2, x1, ym + bh / 2, fill=G_VENDIDO, outline="")
            c.create_text(x1 + 6, ym, text=fmt_qtd(q), anchor="w", font=(FONTE, 8), fill=COR_SUAVE)

    # ================= páginas da rede: Transferências · Saúde do estoque
    def _ir_pagina(self, pag):
        if pag == self.pagina:
            return
        self.pagina = pag
        for k, fr in self.paginas.items():
            if k == pag:
                fr.pack(fill="both", expand=True)
            else:
                fr.pack_forget()
        for k, b in self.btn_pag.items():
            ativo = k == pag
            b.config(bg=COR_ACENTO_CLARO if ativo else COR_CARD, fg=COR_ACENTO if ativo else COR_SUAVE)
        if pag == "consulta":
            self.ent.focus_set()
        else:
            self._atualizar_pagina()

    # ----- peças de interface reutilizadas nas duas páginas
    @staticmethod
    def _rotulo(pai, texto, padx=(0, 6)):
        tk.Label(pai, text=texto, bg=pai.cget("bg"), fg=COR_SUAVE, font=(FONTE, 9)).pack(side="left", padx=padx)

    @staticmethod
    def _combo(pai, var, valores, largura, ao_mudar, editavel=False):
        cb = ttk.Combobox(pai, textvariable=var, values=valores, width=largura, font=(FONTE, 10),
                          state="normal" if editavel else "readonly")
        cb.pack(side="left")
        cb.bind("<<ComboboxSelected>>", lambda e: ao_mudar())
        return cb

    def _entrada(self, pai, var, largura, ao_mudar):
        caixa = tk.Frame(pai, bg=COR_CARD, highlightthickness=1, highlightbackground="#D0D5DD")
        caixa.pack(side="left")
        e = tk.Entry(caixa, textvariable=var, width=largura, relief="flat", bd=0, font=(FONTE, 10),
                     bg=COR_CARD, fg=COR_TEXTO, highlightthickness=0, insertwidth=1)
        e.pack(padx=8, ipady=3)
        e.bind("<FocusIn>", lambda ev: caixa.config(highlightbackground=COR_ACENTO))
        e.bind("<FocusOut>", lambda ev: caixa.config(highlightbackground="#D0D5DD"))

        def adiar(_ev=None):
            if getattr(self, "_apos_busca", None):
                self.after_cancel(self._apos_busca)
            self._apos_busca = self.after(350, ao_mudar)
        e.bind("<KeyRelease>", adiar)
        return e

    def _linha_kpis(self, pai, itens):
        fr = tk.Frame(pai, bg=COR_FUNDO)
        fr.pack(fill="x", pady=12)
        out = {}
        for i, (chave, titulo) in enumerate(itens):
            f = cartao(fr, padx=14, pady=10)
            f.grid(row=0, column=i, sticky="nsew", padx=(0 if i == 0 else 5, 0 if i == len(itens) - 1 else 5))
            fr.columnconfigure(i, weight=1, uniform="k")
            t = tk.Label(f, text=titulo, bg=COR_CARD, fg=COR_SUAVE, font=(FONTE, 9))
            t.pack(anchor="w")
            v = tk.Label(f, text="—", bg=COR_CARD, fg=COR_TEXTO, font=(FONTE_FORTE, 17))
            v.pack(anchor="w", pady=(2, 0))
            s = tk.Label(f, text="", bg=COR_CARD, fg=COR_FRACO, font=(FONTE, 8), anchor="w", justify="left")
            s.pack(anchor="w", fill="x")
            s.bind("<Configure>", lambda e, lb=s: lb.config(wraplength=max(80, e.width)))
            out[chave] = (v, s, t)
        return out

    @staticmethod
    def _kpi_em(kp, chave, valor, sub="", cor=COR_TEXTO):
        v, s, _ = kp[chave]
        v.config(text=valor, fg=cor)
        s.config(text=sub)

    def _filtros_rede(self, pai, nota):
        """Cartão de filtros comum: período-base, atender, subgrupo e recorrência (+ linha extra)."""
        card = cartao(pai)
        card.pack(fill="x")
        l1 = tk.Frame(card, bg=COR_CARD)
        l1.pack(fill="x", padx=16, pady=(12, 0))
        self._rotulo(l1, "Período-base")
        for i, var in enumerate((self.var_de, self.var_ate)):
            if i:
                self._rotulo(l1, "até", (6, 6))
            cb = ttk.Combobox(l1, textvariable=var, width=7, state="readonly", font=(FONTE, 10))
            cb.pack(side="left")
            cb.bind("<<ComboboxSelected>>", self._ao_mudar_periodo)
            self.cbs_periodo.append(cb)
        self._rotulo(l1, "Atender", (16, 6))
        self._combo(l1, self.var_atender, list(self.cb_atender["values"]), 9, self._recalcular)
        self._rotulo(l1, "Subgrupo", (16, 6))
        cbs = self._combo(l1, self.var_subgrupo, [self.TODOS_SUB], 46, self._ao_mudar_subgrupo, editavel=True)
        cbs.bind("<KeyRelease>", lambda e, c=cbs: self._digitar_subgrupo(e, c))
        cbs.bind("<Return>", lambda e, c=cbs: self._confirmar_subgrupo(c))
        self.cbs_subgrupo.append(cbs)
        _Dica(cbs, "Digite parte do código ou do nome para filtrar a lista; Enter aplica.")
        self._rotulo(l1, "Recorrência", (16, 6))
        cbr = self._combo(l1, self.var_recorr, ["1 mês ou mais", "2 meses ou mais", "3 meses ou mais",
                                               "6 meses ou mais"], 15, self._atualizar_pagina)
        _Dica(cbr, "Falta (ruptura, abaixo) e destino de transferência só contam para a loja que vendeu\n"
                   "ou perdeu venda da peça em pelo menos N meses do período-base.")
        l2 = tk.Frame(card, bg=COR_CARD)
        l2.pack(fill="x", padx=16, pady=(10, 4))
        lb = tk.Label(card, text=nota, bg=COR_CARD, fg=COR_FRACO, font=(FONTE, 8), anchor="w", justify="left")
        lb.pack(fill="x", padx=16, pady=(0, 10))
        lb.bind("<Configure>", lambda e: lb.config(wraplength=max(200, e.width - 4)))
        return l2

    def _digitar_subgrupo(self, e, cb):
        if e.keysym in ("Return", "KP_Enter", "Escape", "Up", "Down", "Tab"):
            return
        t = cb.get().strip().upper()
        rot = getattr(self, "subgrupos_rot", [self.TODOS_SUB])
        cb["values"] = ([s for s in rot if t in s.upper()] or rot) if t else rot

    def _confirmar_subgrupo(self, cb):
        rot = getattr(self, "subgrupos_rot", [self.TODOS_SUB])
        t = cb.get().strip()
        if t in rot:
            escolha = t
        else:
            cands = [s for s in rot if t.upper() in s.upper()] if t else [self.TODOS_SUB]
            escolha = cands[0] if cands else self._sub_aplicado
        self.var_subgrupo.set(escolha)
        self._ao_mudar_subgrupo()

    def _ao_mudar_subgrupo(self):
        for cb in self.cbs_subgrupo:
            cb["values"] = getattr(self, "subgrupos_rot", [self.TODOS_SUB])
        if self.var_subgrupo.get() == self._sub_aplicado:
            return
        self._sub_aplicado = self.var_subgrupo.get()
        self._atualizar_pagina()

    def _sub_codigo(self):
        s = self._sub_aplicado
        return None if not s or s == self.TODOS_SUB else s.split(" · ")[0].strip()

    def _recorr_n(self) -> int:
        m = re.match(r"\d+", self.var_recorr.get())
        return int(m.group()) if m else 2

    @staticmethod
    def _loja_cod(txt):
        m = re.match(r"\d+", txt or "")
        return int(m.group()) if m else None

    def _painel_filtrado(self):
        de, ate = self._periodo_base()
        if de is None:
            return None
        pn = rede.painel_rede(self.base, de, ate)
        cod = self._sub_codigo()
        if cod:
            attrs = dict(pn.attrs)
            pn = pn[pn["subgrupo"] == cod]
            pn.attrs.update(attrs)
        return pn

    def _atualizar_pagina(self):
        if self.pagina not in ("transf", "saude") or self.overlay.winfo_ismapped():
            return
        if self.base.vendas.empty and self.base.estoque.empty:
            return
        self.config(cursor="watch")
        self.status("Calculando a rede...")
        self.update_idletasks()
        try:
            if self.pagina == "transf":
                self._atualizar_transf()
            else:
                self._atualizar_saude()
        finally:
            self.config(cursor="")

    def _abrir_peca(self, cod):
        self._ir_pagina("consulta")
        self.var_busca.set(cod)
        self.lista_codigos = []
        self._chips()
        self.mostrar(cod)

    def _abrir_da_tabela(self, tv, col):
        k = tv.focus()
        vals = tv.item(k, "values") if k else ()
        if len(vals) > col and vals[col]:
            self._abrir_peca(str(vals[col]))

    def _exportar(self, folhas: dict, nome: str):
        from tkinter import filedialog
        arq = filedialog.asksaveasfilename(parent=self, title="Exportar para Excel", defaultextension=".xlsx",
                                           initialdir=str(PASTA), initialfile=nome,
                                           filetypes=[("Planilha Excel", "*.xlsx")])
        if not arq:
            return
        try:
            with pd.ExcelWriter(arq) as xw:
                for aba, df in folhas.items():
                    df.to_excel(xw, sheet_name=aba[:31], index=False)
        except Exception as e:  # noqa: BLE001  (arquivo aberto no Excel, sem permissão...)
            messagebox.showerror("Não foi possível exportar", str(e))
            return
        self.status(f"Exportado: {arq}")

    def _periodo_txt(self):
        de, ate = self._periodo_base()
        n = self._meses_atender()
        sub = self._sub_codigo()
        return (f"base {nome_mes(de)}–{nome_mes(ate)} · atender {n} {'mês' if n == 1 else 'meses'}"
                + (f" · subgrupo {sub}" if sub else ""))

    # ----- página Transferências
    def _montar_transf(self, pg):
        l2 = self._filtros_rede(pg, "A loja de origem guarda a própria previsão de Atender + 1 mês; só o que "
                                    "passar disso é sugerido para a loja que tem falta nos meses de Atender. "
                                    "Peça parada (sem venda no período) sai primeiro. Valor = custo unitário "
                                    "do 1211 na origem.")
        lojas = ["Todas"] + [nome_filial(f) for f in sorted(NOMES_FILIAIS)]
        self.var_t_orig, self.var_t_dest = tk.StringVar(value="Todas"), tk.StringVar(value="Todas")
        self.var_t_vmin, self.var_t_busca = tk.StringVar(value="R$ 0"), tk.StringVar()
        self._rotulo(l2, "Origem")
        self._combo(l2, self.var_t_orig, lojas, 9, self._preencher_transf)
        self._rotulo(l2, "Destino", (16, 6))
        self._combo(l2, self.var_t_dest, lojas, 9, self._preencher_transf)
        self._rotulo(l2, "Valor mínimo", (16, 6))
        self._combo(l2, self.var_t_vmin, ["R$ 0", "R$ 50", "R$ 100", "R$ 500", "R$ 1.000"], 8,
                    self._preencher_transf)
        self._rotulo(l2, "Buscar peça", (16, 6))
        self._entrada(l2, self.var_t_busca, 22, self._preencher_transf)
        BotaoPlano(l2, "Exportar Excel", self._exportar_transf, estilo="secundario", fonte=(FONTE_FORTE, 9),
                   padx=12, pady=4).pack(side="right")
        BotaoPlano(l2, "Copiar tabela", lambda: self._copiar_tv(self.tv_transf, "Transferências sugeridas"),
                   estilo="secundario", fonte=(FONTE_FORTE, 9), padx=12, pady=4).pack(side="right", padx=8)
        self.kpis_t = self._linha_kpis(pg, [("n", "Sugestões"), ("un", "Unidades"),
                                            ("valor", "Valor a transferir"), ("pecas", "Peças diferentes"),
                                            ("rup", "Rupturas resolvidas")])
        area = cartao(pg)
        area.pack(fill="both", expand=True)
        area.columnconfigure(0, weight=74)
        area.columnconfigure(2, weight=26)
        area.rowconfigure(1, weight=1)
        topo = tk.Frame(area, bg=COR_CARD)
        topo.grid(row=0, column=0, sticky="ew", padx=12, pady=(10, 0))
        self.lbl_t_info = tk.Label(topo, text="", bg=COR_CARD, fg=COR_FRACO, font=(FONTE, 9))
        self.lbl_t_info.pack(side="left")
        self._legenda(topo, [(COR_SEM_ESTOQUE, "destino sem estoque")])
        fr = tk.Frame(area, bg=COR_CARD)
        fr.grid(row=1, column=0, sticky="nsew")
        self.tv_transf = self._tabela(fr, [
            ("produto", "Produto", 118, "w"), ("desc", "Descrição", 170, "w"), ("sub", "Subgr.", 62, "w"),
            ("de", "De", 66, "w"), ("est_de", "Estoque", 74, "e"), ("cob_de", "Cobre (m)", 84, "e"),
            ("para", "Para", 66, "w"), ("est_para", "Estoque", 74, "e"),
            ("dem_para", "Dem./mês", 82, "e"), ("qtd", "Transferir", 86, "e"),
            ("custo", "Custo unit.", 96, "e"), ("valor", "Valor", 100, "e")], altura=10)
        self._fixar_colunas(self.tv_transf, "desc")
        self.tv_transf.bind("<Double-1>", lambda e: self._abrir_da_tabela(self.tv_transf, 0))
        self.tv_transf.bind("<<TreeviewSelect>>", lambda e: self._motivo_transf())
        tk.Frame(area, bg=COR_BORDA, width=1).grid(row=0, column=1, rowspan=2, sticky="ns", pady=16)
        self.g_transf = Grafico(area)
        self.g_transf.grid(row=0, column=2, rowspan=2, sticky="nsew", padx=(0, 4), pady=6)
        self.g_transf_dados = []

    def _atualizar_transf(self):
        pn = self._painel_filtrado()
        if pn is None:
            return
        self.df_transf = rede.transferencias(pn, self._meses_atender(), self._recorr_n())
        tr = self.df_transf
        env = tr.groupby("de")["valor"].sum() if not tr.empty else pd.Series(dtype=float)
        rec = tr.groupby("para")["valor"].sum() if not tr.empty else pd.Series(dtype=float)
        self.g_transf_dados = [(f, float(env.get(f, 0.0)), float(rec.get(f, 0.0))) for f in sorted(NOMES_FILIAIS)]
        self._preencher_transf()
        self.status(f"Transferências · {self._periodo_txt()} · {len(tr):,} sugestões na rede".replace(",", "."))

    def _filtro_transf(self):
        tr = getattr(self, "df_transf", None)
        if tr is None:
            return None
        f = tr
        o, d = self._loja_cod(self.var_t_orig.get()), self._loja_cod(self.var_t_dest.get())
        if o:
            f = f[f["de"] == o]
        if d:
            f = f[f["para"] == d]
        vmin = float(re.sub(r"\D", "", self.var_t_vmin.get()) or 0)
        if vmin:
            f = f[f["valor"] >= vmin]
        b = self.var_t_busca.get().strip().upper()
        if b:
            f = f[f["produto"].str.upper().str.contains(b, regex=False) |
                  f["descricao"].astype(str).str.upper().str.contains(b, regex=False)]
        return f

    def _preencher_transf(self):
        f = self._filtro_transf()
        if f is None:
            return
        self.df_transf_vis = f
        tv = self.tv_transf
        tv.delete(*tv.get_children())
        for i, r in enumerate(f.head(self.MAX_LINHAS).itertuples(index=False)):
            cob_de = "parada" if r.media_de <= 0 else fmt_num(r.est_de / r.media_de, 1)
            tag = "semestoque" if r.est_para <= 0 else ("par" if i % 2 else "impar")
            tv.insert("", "end", iid=f"r{i}", tags=(tag,), values=(
                r.produto, r.descricao, r.subgrupo, nome_filial(r.de), fmt_qtd(r.est_de), cob_de,
                nome_filial(r.para), fmt_qtd(r.est_para), fmt_prev(r.media_para), fmt_qtd(r.qtd),
                fmt_moeda(r.custo), fmt_moeda(r.valor)))
        n_rup = f.loc[f["est_para"] <= 0, ["para", "produto"]].drop_duplicates().shape[0]
        self._kpi_em(self.kpis_t, "n", fmt_num(len(f), 0), "linhas origem → destino")
        self._kpi_em(self.kpis_t, "un", fmt_qtd(f["qtd"].sum()), "unidades a mover")
        self._kpi_em(self.kpis_t, "valor", fmt_compacto(f["valor"].sum()), "a custo da origem",
                     cor=COR_ACENTO if len(f) else COR_TEXTO)
        self._kpi_em(self.kpis_t, "pecas", fmt_num(f["produto"].nunique(), 0), "códigos com sugestão")
        self._kpi_em(self.kpis_t, "rup", fmt_num(n_rup, 0), "loja × peça hoje sem estoque",
                     cor=COR_VERMELHO if n_rup else COR_TEXTO)
        extra = (f" · mostrando as {fmt_num(self.MAX_LINHAS, 0)} de maior valor (o Excel leva todas)"
                 if len(f) > self.MAX_LINHAS else "")
        self._t_info = f"{fmt_num(len(f), 0)} sugestões{extra} · clique mostra o motivo · duplo clique abre a peça"
        self.lbl_t_info.config(text=self._t_info, fg=COR_FRACO)
        self.g_transf.definir(self._g_transf)

    def _motivo_transf(self):
        sel = self.tv_transf.selection()
        f = getattr(self, "df_transf_vis", None)
        if not sel or f is None or not sel[0].startswith("r"):
            return
        r = f.iloc[int(sel[0][1:])]
        self.lbl_t_info.config(fg=COR_TEXTO, text=f"{r['produto']}: {nome_filial(r['de'])} → "
                                                  f"{nome_filial(r['para'])} · {r['motivo']}")

    @staticmethod
    def _fixar_colunas(tv, elastica):
        """Só a coluna 'elastica' estica; as demais mantêm a largura (evita texto cortado)."""
        for c in tv["columns"]:
            tv.column(c, stretch=c == elastica, minwidth=40)

    def _g_transf(self, c: Grafico, W, H):
        dados = self.g_transf_dados
        c.titulo("Por loja", "Clique: o que envia · botão direito: o que recebe")
        L, R, T, B = 64, 128, 84, H - 12
        if not dados or B - T < 40:
            return
        cols = [(W - 72, "envia", G_VENDIDO), (W - 14, "recebe", G_ESTOQUE)]
        for x, t, cor in cols:
            tx = c.create_text(x, T - 12, text=t, anchor="e", font=(FONTE, 8), fill=COR_SUAVE)
            bx = c.bbox(tx)[0]
            c.create_oval(bx - 10, T - 15, bx - 4, T - 9, fill=cor, outline="")
        vmax = max([max(e, r) for _, e, r in dados] + [1.0])
        rh = (B - T) / len(dados)
        bh = max(3, min(9, (rh - 8) / 2))
        o_sel, d_sel = self._loja_cod(self.var_t_orig.get()), self._loja_cod(self.var_t_dest.get())
        for k, (f, env, rec) in enumerate(dados):
            y0 = T + k * rh
            tag = f"t{f}"
            sel = f in (o_sel, d_sel)
            c.create_rectangle(0, y0, W, y0 + rh, fill=COR_ACENTO_CLARO if sel else COR_CARD, outline="", tags=tag)
            if k:
                c.create_line(16, y0, W - 16, y0, fill=COR_GRADE)
            ym = y0 + rh / 2
            c.create_text(L - 8, ym, text=f"{f:02d} {sigla_filial(f)}", anchor="e", tags=tag,
                          font=(FONTE_FORTE if sel else FONTE, 9), fill=COR_TEXTO)
            for j, (val, cor) in enumerate([(env, G_VENDIDO), (rec, G_ESTOQUE)]):
                yb = ym - bh - 0.5 + j * (bh + 1)
                if val > 0:
                    c.create_rectangle(L, yb, max(L + 2, L + val / vmax * (W - L - R)), yb + bh,
                                       fill=cor, outline="", tags=tag)
                c.create_text(cols[j][0], ym, text=fmt_compacto(val).replace("R$ ", "") if val else "–",
                              anchor="e", font=(FONTE, 9), fill=COR_TEXTO, tags=tag)
            c.dica(tag, f"{nome_filial(f)} — envia {fmt_moeda(env)} · recebe {fmt_moeda(rec)}\n"
                        "(clique: o que envia · botão direito: o que recebe)")
            c.tag_bind(tag, "<Button-1>", lambda e, f=f: self._loja_transf(f, "orig"))
            c.tag_bind(tag, "<Button-3>", lambda e, f=f: self._loja_transf(f, "dest"))
            c.tag_bind(tag, "<Enter>", lambda e: c.config(cursor="hand2"), add="+")
            c.tag_bind(tag, "<Leave>", lambda e: c.config(cursor=""), add="+")

    def _loja_transf(self, f, lado):
        alvo, outro = (self.var_t_orig, self.var_t_dest) if lado == "orig" else (self.var_t_dest, self.var_t_orig)
        alvo.set("Todas" if self._loja_cod(alvo.get()) == f else nome_filial(f))
        outro.set("Todas")
        self._preencher_transf()

    def _exportar_transf(self):
        f = getattr(self, "df_transf_vis", None)
        if f is None or f.empty:
            messagebox.showinfo("Nada para exportar", "Nenhuma sugestão com os filtros atuais.")
            return
        df = pd.DataFrame({
            "Produto": f["produto"], "Descrição": f["descricao"], "Subgrupo": f["subgrupo"],
            "De (loja)": f["de"].map(nome_filial), "Estoque origem": f["est_de"],
            "Demanda/mês origem": f["media_de"].round(2), "Para (loja)": f["para"].map(nome_filial),
            "Estoque destino": f["est_para"], "Demanda/mês destino": f["media_para"].round(2),
            "Transferir": f["qtd"], "Custo unit.": f["custo"].round(2), "Valor": f["valor"].round(2),
            "Motivo": f["motivo"]})
        self._exportar({"Transferências": df, "Parâmetros": pd.DataFrame({"Filtro": [self._periodo_txt()]})},
                       f"transferencias_{date.today():%Y-%m-%d}.xlsx")

    # ----- página Saúde do estoque
    def _montar_saude(self, pg):
        l2 = self._filtros_rede(pg, "Situação da peça em cada loja: Ruptura = sem estoque, com demanda recorrente · "
                                    "Abaixo = estoque menor que a previsão de Atender · Excesso = cobertura acima "
                                    "do limite (só a parte que passa dele entra no valor) · Parada = estoque sem "
                                    "nenhuma venda nem venda perdida no período-base. Valores a custo (1211).")
        self.var_limite, self.var_s_busca = tk.StringVar(value="6 meses"), tk.StringVar()
        self.loja_s, self.st_sel = None, "Todos"
        self._rotulo(l2, "Excesso acima de")
        self._combo(l2, self.var_limite, ["3 meses", "4 meses", "6 meses", "9 meses", "12 meses", "18 meses",
                                          "24 meses"], 9, self._atualizar_pagina)
        self._rotulo(l2, "de cobertura", (6, 0))
        self._rotulo(l2, "Buscar peça", (16, 6))
        self._entrada(l2, self.var_s_busca, 22, self._preencher_itens)
        BotaoPlano(l2, "Exportar Excel", self._exportar_saude, estilo="secundario", fonte=(FONTE_FORTE, 9),
                   padx=12, pady=4).pack(side="right")
        BotaoPlano(l2, "Copiar peças", lambda: self._copiar_tv(self.tv_itens, "Saúde do estoque — peças"),
                   estilo="secundario", fonte=(FONTE_FORTE, 9), padx=12, pady=4).pack(side="right", padx=8)
        BotaoPlano(l2, "Copiar lojas", lambda: self._copiar_tv(self.tv_lojas, "Saúde do estoque — lojas"),
                   estilo="secundario", fonte=(FONTE_FORTE, 9), padx=12, pady=4).pack(side="right")
        self.kpis_s = self._linha_kpis(pg, [("valor", "Valor em estoque"), ("cob", "Cobertura"),
                                            ("giro", "Giro 12 meses"), ("parado", "Parado"),
                                            ("excesso", "Excesso"), ("rup", "Rupturas"),
                                            ("transf", "Transferível")])
        corpo = tk.Frame(pg, bg=COR_FUNDO)
        corpo.pack(fill="both", expand=True)
        corpo.columnconfigure(0, weight=70)
        corpo.columnconfigure(1, weight=30)
        corpo.rowconfigure(0, weight=45, uniform="s")
        corpo.rowconfigure(1, weight=55, uniform="s")
        sup = cartao(corpo)
        sup.grid(row=0, column=0, sticky="nsew", pady=(0, 10))
        self.tv_lojas = self._tabela(sup, [
            ("loja", "Loja", 70, "w"), ("valor", "Estoque R$", 92, "e"), ("itens", "Itens", 56, "e"),
            ("cob", "Cobre (m)", 82, "e"), ("giro", "Giro", 50, "e"), ("parado", "Parado R$", 88, "e"),
            ("parado_p", "%", 44, "e"), ("excesso", "Excesso R$", 92, "e"), ("excesso_p", "%", 44, "e"),
            ("abaixo", "Abaixo", 64, "e"), ("rup", "Rupturas", 76, "e"), ("perd", "Perdida", 70, "e"),
            ("transf", "Transferível", 96, "e")], altura=6)
        self._fixar_colunas(self.tv_lojas, "loja")
        self.tv_lojas.bind("<<TreeviewSelect>>", self._ao_sel_loja)
        lado = cartao(corpo)
        lado.grid(row=0, column=1, rowspan=2, sticky="nsew", padx=(10, 0))
        self.g_saude = Grafico(lado)
        self.g_saude.pack(fill="both", expand=True, padx=4, pady=6)
        inf = cartao(corpo)
        inf.grid(row=1, column=0, sticky="nsew")
        cab = tk.Frame(inf, bg=COR_CARD)
        cab.pack(fill="x", padx=12, pady=(10, 0))
        self.lbl_itens = tk.Label(cab, text="Peças", bg=COR_CARD, fg=COR_TEXTO, font=(FONTE_FORTE, 10))
        self.lbl_itens.pack(side="left")
        self.fr_chips_s = tk.Frame(cab, bg=COR_CARD)
        self.fr_chips_s.pack(side="left", padx=16)
        self.lbl_itens_info = tk.Label(cab, text="", bg=COR_CARD, fg=COR_FRACO, font=(FONTE, 9))
        self.lbl_itens_info.pack(side="right")
        self.tv_itens = self._tabela(inf, [
            ("loja", "Loja", 70, "w"), ("produto", "Produto", 118, "w"), ("desc", "Descrição", 200, "w"),
            ("sub", "Subgr.", 62, "w"), ("st", "Situação", 80, "w"), ("est", "Estoque", 74, "e"),
            ("dem", "Dem./mês", 82, "e"), ("cob", "Cobre (m)", 84, "e"),
            ("valor", "Estoque R$", 100, "e"), ("exc", "Excesso R$", 100, "e"),
            ("ult", "Últ. venda", 96, "center")], altura=6)
        self._fixar_colunas(self.tv_itens, "desc")
        self.tv_itens.tag_configure("st_Ruptura", background=COR_SEM_ESTOQUE)
        self.tv_itens.tag_configure("st_Abaixo", background="#FFF6ED")
        self.tv_itens.tag_configure("st_Excesso", background="#FFFAEB")
        self.tv_itens.tag_configure("st_Parada", foreground=COR_SUAVE, background=COR_LINHA_PAR)
        self.tv_itens.bind("<Double-1>", lambda e: self._abrir_da_tabela(self.tv_itens, 1))

    def _limite(self) -> float:
        m = re.match(r"\d+", self.var_limite.get())
        return float(m.group()) if m else 6.0

    def _atualizar_saude(self):
        pn = self._painel_filtrado()
        if pn is None:
            return
        n, rc = self._meses_atender(), self._recorr_n()
        cl = rede.classificar(pn, n, self._limite(), rc)
        tr = rede.transferencias(pn, n, rc)
        rl = rede.resumo_lojas(cl, tr)
        self.df_cl, self.df_rl = cl, rl
        tv = self.tv_lojas
        tv.delete(*tv.get_children())

        def pct(a, b):
            return f"{fmt_num(100 * a / b, 0)}%" if b else ""

        def valores(nome, r):
            return (nome, fmt_num(r["valor_est"], 0), fmt_num(r["saldo"], 0),
                    fmt_num(r["cob"], 1) if pd.notna(r["cob"]) else "—",
                    fmt_num(r["giro"], 1) if pd.notna(r["giro"]) else "—",
                    fmt_num(r["parado_valor"], 0), pct(r["parado_valor"], r["valor_est"]),
                    fmt_num(r["excesso_valor"], 0), pct(r["excesso_valor"], r["valor_est"]),
                    fmt_num(r["n_Abaixo"], 0), fmt_num(r["n_Ruptura"], 0), fmt_num(r["perd_qtd"], 0),
                    fmt_num(r["transf"], 0))
        lojas = [f for f in sorted(NOMES_FILIAIS) if f in rl.index] + \
                [f for f in rl.index if f != -1 and f not in NOMES_FILIAIS]
        for i, f in enumerate(lojas):
            tv.insert("", "end", iid=f"l{f}", values=valores(nome_filial(f), rl.loc[f]),
                      tags=("par" if i % 2 else "impar",))
        if -1 in rl.index:
            tv.insert("", "end", iid="total", values=valores("TOTAL REDE", rl.loc[-1]), tags=("total",))
            t = rl.loc[-1]
            self._kpi_em(self.kpis_s, "valor", fmt_compacto(t["valor_est"]),
                         f"{fmt_num(t['saldo'], 0)} itens com saldo")
            self._kpi_em(self.kpis_s, "cob", f"{fmt_num(t['cob'], 1)} meses" if pd.notna(t["cob"]) else "—",
                         "estoque ÷ demanda/mês")
            self._kpi_em(self.kpis_s, "giro", f"{fmt_num(t['giro'], 1)}×" if pd.notna(t["giro"]) else "—",
                         "custo vendido 12m ÷ estoque")
            self._kpi_em(self.kpis_s, "parado", fmt_compacto(t["parado_valor"]),
                         f"{pct(t['parado_valor'], t['valor_est'])} · {fmt_num(t['n_Parada'], 0)} itens",
                         cor=COR_LARANJA if t["parado_valor"] else COR_TEXTO)
            self._kpi_em(self.kpis_s, "excesso", fmt_compacto(t["excesso_valor"]),
                         f"{pct(t['excesso_valor'], t['valor_est'])} · {fmt_num(t['n_Excesso'], 0)} itens")
            self._kpi_em(self.kpis_s, "rup", fmt_num(t["n_Ruptura"], 0),
                         f"{fmt_num(t['n_Abaixo'], 0)} itens abaixo",
                         cor=COR_VERMELHO if t["n_Ruptura"] else COR_TEXTO)
            self._kpi_em(self.kpis_s, "transf", fmt_compacto(t["transf"]), "pode ir para outra loja",
                         cor=COR_ACENTO if t["transf"] else COR_TEXTO)
        else:
            for k in self.kpis_s:
                self._kpi_em(self.kpis_s, k, "—")
        if self.loja_s is not None and f"l{self.loja_s}" in tv.get_children():
            tv.selection_set(f"l{self.loja_s}")
        else:
            self.loja_s = None
        self._preencher_itens()
        self.g_saude.definir(self._g_saude)
        self.status(f"Saúde do estoque · {self._periodo_txt()} · excesso acima de "
                    f"{fmt_num(self._limite(), 0)} meses")

    def _ao_sel_loja(self, _e=None):
        sel = self.tv_lojas.selection()
        novo = None if not sel or sel[0] == "total" else int(sel[0][1:])
        if novo != self.loja_s:
            self.loja_s = novo
            self._preencher_itens()
            self.g_saude.redesenhar()

    def _sel_loja_saude(self, f):
        if f == self.loja_s:
            self.tv_lojas.selection_remove(self.tv_lojas.selection())
        elif f"l{f}" in self.tv_lojas.get_children():
            self.tv_lojas.selection_set(f"l{f}")
            self.tv_lojas.see(f"l{f}")

    def _status_itens(self, st):
        self.st_sel = st
        self._preencher_itens()

    def _preencher_itens(self):
        cl = getattr(self, "df_cl", None)
        if cl is None:
            return
        f = cl if self.loja_s is None else cl[cl["filial"] == self.loja_s]
        b = self.var_s_busca.get().strip().upper()
        if b:
            f = f[f["produto"].str.upper().str.contains(b, regex=False) |
                  f["descricao"].astype(str).str.upper().str.contains(b, regex=False)]
        cont = f["status"].value_counts()
        for w in self.fr_chips_s.winfo_children():
            w.destroy()
        for st in ["Todos"] + rede.STATUS:
            n = len(f) if st == "Todos" else int(cont.get(st, 0))
            BotaoPlano(self.fr_chips_s, f"{st}  {fmt_num(n, 0)}", lambda s=st: self._status_itens(s),
                       estilo="primario" if st == self.st_sel else "secundario", fonte=(FONTE_FORTE, 9),
                       padx=8, pady=2).pack(side="left", padx=2)
        if self.st_sel != "Todos":
            f = f[f["status"] == self.st_sel]
        f = f.sort_values(["valor_est", "dem_valor"], ascending=False, kind="stable")
        self.df_itens_vis = f
        tv = self.tv_itens
        tv.delete(*tv.get_children())
        for r in f.head(self.MAX_LINHAS).itertuples(index=False):
            tv.insert("", "end", tags=(f"st_{r.status}",), values=(
                nome_filial(r.filial), r.produto, r.descricao, r.subgrupo, r.status, fmt_qtd(r.est0),
                fmt_prev(r.media), "—" if not np.isfinite(r.cob) else fmt_num(r.cob, 1),
                fmt_moeda(r.valor_est) if r.valor_est else "", fmt_moeda(r.excesso_valor) if r.excesso_valor else "",
                fmt_data(r.ult_venda) if pd.notna(r.ult_venda) else "—"))
        self.lbl_itens.config(text=f"Peças — {nome_filial(self.loja_s) if self.loja_s else 'todas as lojas'}")
        extra = f"mostrando {fmt_num(self.MAX_LINHAS, 0)} de {fmt_num(len(f), 0)} · " if len(f) > self.MAX_LINHAS else ""
        self.lbl_itens_info.config(text=f"{extra}duplo clique abre a peça")

    def _g_saude(self, c: Grafico, W, H):
        rl = getattr(self, "df_rl", None)
        c.titulo("Composição do estoque", "Valor a custo por loja · clique para ver as peças")
        c.legenda([(G_VENDIDO, "em giro", None), (G_PERDIDO, "excesso", None), ("#C9CED6", "parado", None)], y=66)
        if rl is None:
            return
        lojas = [f for f in sorted(NOMES_FILIAIS) if f in rl.index]
        L, R, T, B = 64, 170, 104, H - 10
        if not lojas or B - T < 40:
            return
        rh = (B - T) / len(lojas)
        bh = max(6, min(14, rh - 12))
        for x, t in [(W - R + 10, "estoque"), (W - 14, "rupturas")]:
            c.create_text(x, T - 12, text=t, anchor="w" if x < W - 20 else "e", font=(FONTE, 8), fill=COR_SUAVE)
        for k, f in enumerate(lojas):
            r = rl.loc[f]
            y0 = T + k * rh
            ym = y0 + rh / 2
            tag = f"s{f}"
            sel = f == self.loja_s
            c.create_rectangle(0, y0, W, y0 + rh, fill=COR_ACENTO_CLARO if sel else COR_CARD, outline="", tags=tag)
            c.create_text(L - 8, ym, text=f"{f:02d} {sigla_filial(f)}", anchor="e", tags=tag,
                          font=(FONTE_FORTE if sel else FONTE, 9), fill=COR_TEXTO)
            tot = r["valor_est"]
            x = L
            partes = [(r["em_giro_valor"], G_VENDIDO), (r["excesso_valor"], G_PERDIDO), (r["parado_valor"], "#C9CED6")]
            if tot > 0:
                for val, cor in partes:
                    w = val / tot * (W - L - R)
                    if w > 0.5:
                        c.create_rectangle(x, ym - bh / 2, x + w, ym + bh / 2, fill=cor, outline=COR_CARD, tags=tag)
                    x += w
            else:
                c.create_rectangle(L, ym - bh / 2, W - R, ym + bh / 2, fill=COR_GRADE, outline="", tags=tag)
            c.create_text(W - R + 10, ym, text=fmt_compacto(tot), anchor="w", font=(FONTE, 9),
                          fill=COR_TEXTO, tags=tag)
            if k:
                c.create_line(16, y0, W - 16, y0, fill=COR_GRADE)
            c.create_text(W - 14, ym, text=fmt_num(r["n_Ruptura"], 0) if r["n_Ruptura"] else "–", anchor="e",
                          font=(FONTE, 9), fill=COR_VERMELHO if r["n_Ruptura"] else COR_FRACO, tags=tag)

            def p(v, t=tot):
                return f"{fmt_num(100 * v / t, 0)}%" if t else "0%"
            c.dica(tag, f"{nome_filial(f)} — {fmt_moeda(tot)} em estoque\n"
                        f"em giro {p(r['em_giro_valor'])} · excesso {p(r['excesso_valor'])} · "
                        f"parado {p(r['parado_valor'])} · {fmt_num(r['n_Ruptura'], 0)} rupturas")
            c.tag_bind(tag, "<Button-1>", lambda e, f=f: self._sel_loja_saude(f))
            c.tag_bind(tag, "<Enter>", lambda e: c.config(cursor="hand2"), add="+")
            c.tag_bind(tag, "<Leave>", lambda e: c.config(cursor=""), add="+")

    def _exportar_saude(self):
        rl, f = getattr(self, "df_rl", None), getattr(self, "df_itens_vis", None)
        if rl is None:
            return
        lojas = pd.DataFrame({
            "Loja": [("TOTAL REDE" if i == -1 else nome_filial(i)) for i in rl.index],
            "Valor em estoque": rl["valor_est"].round(2), "Itens com saldo": rl["saldo"].astype(int),
            "Cobertura (meses)": rl["cob"].round(1), "Giro 12m": rl["giro"].round(2),
            "Valor em giro": rl["em_giro_valor"].round(2), "Valor parado": rl["parado_valor"].round(2),
            "Valor em excesso": rl["excesso_valor"].round(2), "Itens abaixo": rl["n_Abaixo"].astype(int),
            "Rupturas": rl["n_Ruptura"].astype(int), "Itens parados": rl["n_Parada"].astype(int),
            "Itens em excesso": rl["n_Excesso"].astype(int), "Venda perdida (un)": rl["perd_qtd"],
            "Venda perdida (R$)": rl["perd_valor"].round(2), "Transferível (R$)": rl["transf"].round(2)})
        pecas = pd.DataFrame({
            "Loja": f["filial"].map(nome_filial), "Produto": f["produto"], "Descrição": f["descricao"],
            "Subgrupo": f["subgrupo"], "Situação": f["status"], "Estoque": f["est0"],
            "Demanda/mês": f["media"].round(2),
            "Cobertura (meses)": f["cob"].where(np.isfinite(f["cob"])).round(1),
            "Valor em estoque": f["valor_est"].round(2), "Valor do excesso": f["excesso_valor"].round(2),
            "Última venda": f["ult_venda"]}) if f is not None else pd.DataFrame()
        self._exportar({"Lojas": lojas, "Peças": pecas,
                        "Parâmetros": pd.DataFrame({"Filtro": [self._periodo_txt() + f" · excesso acima de "
                                                              f"{fmt_num(self._limite(), 0)} meses"]})},
                       f"saude_estoque_{date.today():%Y-%m-%d}.xlsx")

    # ---------------- copiar
    def copiar_aba(self):
        if not self.cod_atual or not self.atual:
            return
        aba = self.nb.index(self.nb.select())
        tv = {self.ABA_RESUMO: self.tv_res, self.ABA_CLIENTES: self.tv_cli, self.ABA_PREV: self.tv_prev,
              self.ABA_VENDAS: self.tv_vend, self.ABA_PERDIDAS: self.tv_perd}.get(aba, self.tv_res)
        self._copiar_tv(tv, f"Peça\t{self.cod_atual}\t{self.base.descricoes.get(self.cod_atual, '')}")

    def _copiar_tv(self, tv, titulo):
        arvore = tv is self.tv_cli
        cab = ([tv.heading("#0")["text"]] if arvore else []) + [tv.heading(c)["text"] for c in tv["columns"]]
        linhas = [titulo if "\t" in titulo else f"{titulo}\t{self._periodo_txt()}", "\t".join(cab)]

        def add(k):
            linhas.append("\t".join(([tv.item(k, "text")] if arvore else []) +
                                    [str(x) for x in tv.item(k, "values")]))
            for filho in tv.get_children(k):
                add(filho)
        for k in tv.get_children(""):
            add(k)
        self.clipboard_clear()
        self.clipboard_append("\n".join(linhas))
        self.status("Tabela copiada — é só colar no Excel (Ctrl+V).")


def main():
    try:
        from ctypes import windll
        windll.shcore.SetProcessDpiAwareness(1)  # texto nítido em telas com escala
    except Exception:  # noqa: BLE001
        pass
    app = App()
    try:  # fecha a tela de abertura do .exe (PyInstaller --splash)
        import pyi_splash  # type: ignore
        app.after(200, pyi_splash.close)
    except ImportError:
        pass
    app.mainloop()


if __name__ == "__main__":
    main()
