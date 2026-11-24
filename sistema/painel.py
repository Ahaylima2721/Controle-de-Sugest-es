"""Dados do painel + geração do HTML offline (mesma página do aplicativo, com os dados embutidos)."""
from __future__ import annotations

import json
import math
from pathlib import Path

import pandas as pd

import config as C

APP_HTML = Path(__file__).parent / "web" / "app.html"
STATUS = ["RUPTURA", "CRÍTICO", "ATENÇÃO"]
EVID = ["VENDA PERDIDA", "VENDA RECENTE", "SÓ SAZONAL (ano ant.)", ""]


def _r(x, n=1):
    if x is None:
        return None
    try:
        x = float(x)
    except (TypeError, ValueError):
        return None
    if math.isnan(x) or math.isinf(x):
        return None
    return round(x, n)


def nomes_subgrupos(estoque: pd.DataFrame) -> dict:
    return (estoque.drop_duplicates("subgrupo").set_index("subgrupo")["subgrupo_desc"]
            .astype(str).str.strip().to_dict())


def dados_painel(df: pd.DataFrame, perdidas: pd.DataFrame, estoque: pd.DataFrame, meta: dict) -> dict:
    nomes_sg = nomes_subgrupos(estoque)
    al = df[df["status"].isin(STATUS)]
    sgs = sorted(set(al["subgrupo"].dropna().astype(int)) | set(perdidas["subgrupo"].dropna().astype(int)))
    st_idx = {s: i for i, s in enumerate(STATUS)}
    ev_idx = {e: i for i, e in enumerate(EVID)}
    itens = []
    for r in al.itertuples(index=False):
        fal = r.data_prevista_falta
        itens.append([
            int(r.filial), r.produto, str(r.descricao or "")[:60],
            int(r.subgrupo) if pd.notna(r.subgrupo) else 0,
            st_idx[r.status], 1 if r.recorrencia == "RECORRENTE" else 0,
            _r(r.estoque, 0), _r(r.demanda_mes, 1), _r(r.cobertura_dias, 0),
            fal.strftime("%d/%m/%Y") if pd.notna(fal) else "",
            _r(r.sugestao_compra, 0), r.sobra_outras_filiais or "",
            _r(r.valor_em_risco, 0), ev_idx.get(r.evidencia, 3), _r(r.perdido_qtd, 0),
        ])
    vp = perdidas[perdidas["data"] > meta["data_base"] - pd.Timedelta(days=meta["janela"])]
    vpg = vp.groupby(["filial", "subgrupo"]).agg(v=("valor", "sum"), n=("qtd", "size")).reset_index()
    return {
        "geradoEm": pd.Timestamp.now().strftime("%d/%m/%Y %H:%M"),
        "vendasAte": meta["data_base"].strftime("%d/%m/%Y"),
        "estoqueDe": meta["data_estoque_ts"].strftime("%d/%m/%Y"),
        "perdidasDe": meta["perdidas_periodo"][0].strftime("%d/%m/%Y"),
        "perdidasAte": meta["perdidas_periodo"][1].strftime("%d/%m/%Y"),
        "params": {"lt": C.LEAD_TIME_DIAS, "seg": C.SEGURANCA_DIAS},
        "status": STATUS, "evid": EVID,
        "subgrupos": [[int(s), nomes_sg.get(s, f"SUBGRUPO {s}")] for s in sgs],
        "filiais": [[f, f"{f:02d} {C.NOMES_FILIAIS.get(f, '')}".strip()] for f in sorted(set(df["filial"].astype(int)))],
        "itens": itens,
        "perdidas": [[int(a), int(b), _r(c, 0), int(d)] for a, b, c, d in vpg.itertuples(index=False)],
    }


def gerar(df, perdidas, estoque, meta, destino: Path):
    """HTML offline: o painel funciona sem o aplicativo; as telas Item/Acurácia/Importar pedem o app."""
    js = json.dumps(dados_painel(df, perdidas, estoque, meta), ensure_ascii=False,
                    separators=(",", ":")).replace("</", "<\\/")
    html = APP_HTML.read_text(encoding="utf-8").replace("/*__DADOS__*/null", js)
    Path(destino).write_text(html, encoding="utf-8")
    return destino
