"""Pipeline: importa relatórios -> atualiza a base -> roda o modelo -> gera saídas.

Pode ser usado de 3 jeitos:
  - Aplicativo (INICIAR.bat): arrastar os arquivos na tela "Importar arquivos".
  - ATUALIZAR.bat: processa o que estiver na pasta entrada\\.
  - python atualizar.py --carga-inicial : importa os relatórios da pasta raiz.

O tipo do relatório é reconhecido pelo CONTEÚDO (cabeçalho), não pelo nome do arquivo.

Regras de mesclagem:
  - 1174 e Vendas Perdidas: o período coberto pelo arquivo novo SUBSTITUI o mesmo
    período na base (pode reenviar o mesmo dia sem duplicar).
  - 1211: vira o estoque atual; uma cópia fica em base\\estoque_historico.
"""
from __future__ import annotations

import argparse
import pickle
import shutil
import sys
import time
from datetime import datetime
from pathlib import Path

import pandas as pd

AQUI = Path(__file__).resolve().parent
sys.path.insert(0, str(AQUI))

import config as C  # noqa: E402
from leitores import ler_1174, ler_1211, ler_vendas_perdidas  # noqa: E402
import modelo  # noqa: E402
import motor  # noqa: E402
import painel  # noqa: E402
import relatorio  # noqa: E402

RAIZ = AQUI.parent
BASE = RAIZ / "base"
ENTRADA = RAIZ / "entrada"
PROCESSADOS = ENTRADA / "processados"
NAO_RECONHECIDOS = ENTRADA / "nao_reconhecidos"
SAIDA = RAIZ / "saida"

ARQ_VENDAS = BASE / "vendas_1174.pkl"
ARQ_PERDIDAS = BASE / "vendas_perdidas.pkl"
ARQ_ESTOQUE = BASE / "estoque_atual.pkl"
HIST_ESTOQUE = BASE / "estoque_historico"
ARQ_MODELO = BASE / "modelo.pkl"
ARQ_PREVISAO = BASE / "previsao.pkl"
HIST_PREVISOES = BASE / "previsoes_historico"
ARQ_EQUIV = AQUI / "equivalencias_manual.csv"
ARQ_HIST_IMPORT = BASE / "importacoes.csv"

NOMES_TIPO = {"1174": "Vendas (1174)", "1211": "Estoque (1211)", "perdidas": "Vendas perdidas"}

_log_fn = print


def log(msg):
    _log_fn(f"[{datetime.now():%H:%M:%S}] {msg}")


# ----------------------------------------------------------------- reconhecimento
def detectar_tipo(arq: Path) -> str | None:
    """Reconhece o relatório pelo cabeçalho. Retorna '1174', '1211', 'perdidas' ou None."""
    ext = arq.suffix.lower()
    try:
        if ext in (".csv", ".txt"):
            with open(arq, "rb") as f:
                cab = f.read(4000)
            txt = cab.decode("utf-8-sig", errors="ignore")
            if "Produto" in txt and ("Custo unit" in txt or "Valor total" in txt) and ";" in txt:
                return "1211"
            return None
        if ext in (".xlsx", ".xlsm"):
            import openpyxl
            wb = openpyxl.load_workbook(arq, read_only=True, data_only=True)
            ws = wb.worksheets[0]
            linhas = []
            for i, r in enumerate(ws.iter_rows(values_only=True)):
                linhas.append([str(x) for x in r if x is not None])
                if i >= 8:
                    break
            wb.close()
            plano = " | ".join(" ".join(l) for l in linhas)
            if "Venda Perdida" in plano:
                return "perdidas"
            if "Tipo de Lançamento" in plano and "Produto" in plano and "Quantidade" in plano:
                return "1174"
    except Exception as e:  # arquivo aberto/corrompido
        log(f"  não consegui ler {arq.name}: {e}")
    return None


def resumo_arquivo(arq: Path, tipo: str) -> str:
    return f"{NOMES_TIPO.get(tipo, tipo)} · {arq.stat().st_size/1e6:.1f} MB"


# ----------------------------------------------------------------- importação
def _mesclar_periodo(base, novo):
    if base is None or base.empty:
        return novo
    ini, fim = novo["data"].min(), novo["data"].max()
    manter = base[(base["data"] < ini) | (base["data"] > fim)]
    out = pd.concat([manter, novo], ignore_index=True)
    for c in out.columns:
        if c in base and base[c].dtype.name == "category":
            out[c] = out[c].astype("category")
    return out


def _registrar(tipo, nome, ini, fim, linhas):
    reg = pd.DataFrame([{"quando": datetime.now().strftime("%d/%m/%Y %H:%M"), "tipo": NOMES_TIPO[tipo],
                         "arquivo": nome, "periodo_ini": ini, "periodo_fim": fim, "linhas": linhas}])
    if ARQ_HIST_IMPORT.exists():
        reg = pd.concat([pd.read_csv(ARQ_HIST_IMPORT, sep=";"), reg], ignore_index=True)
    reg.to_csv(ARQ_HIST_IMPORT, sep=";", index=False)


def importar_1174(arq: Path):
    log(f"Lendo vendas {arq.name} (arquivo grande pode levar alguns minutos)...")
    t = time.time()
    novo = ler_1174(arq)
    base = pd.read_pickle(ARQ_VENDAS) if ARQ_VENDAS.exists() else None
    out = _mesclar_periodo(base, novo)
    out.to_pickle(ARQ_VENDAS)
    log(f"  {len(novo):,} linhas de {novo.data.min():%d/%m/%Y} a {novo.data.max():%d/%m/%Y} "
        f"-> base com {len(out):,} linhas ({time.time()-t:.0f}s)")
    _registrar("1174", arq.name, f"{novo.data.min():%d/%m/%Y}", f"{novo.data.max():%d/%m/%Y}", len(novo))


def importar_perdidas(arq: Path):
    log(f"Lendo vendas perdidas {arq.name}...")
    novo = ler_vendas_perdidas(arq)
    base = pd.read_pickle(ARQ_PERDIDAS) if ARQ_PERDIDAS.exists() else None
    out = _mesclar_periodo(base, novo).drop_duplicates("cod_vp", keep="last")
    out.to_pickle(ARQ_PERDIDAS)
    log(f"  {len(novo):,} registros de {novo.data.min():%d/%m/%Y} a {novo.data.max():%d/%m/%Y} "
        f"-> base com {len(out):,}")
    _registrar("perdidas", arq.name, f"{novo.data.min():%d/%m/%Y}", f"{novo.data.max():%d/%m/%Y}", len(novo))


def importar_1211(arq: Path, data_estoque=None):
    log(f"Lendo estoque {arq.name}...")
    e = ler_1211(arq)
    dia = pd.Timestamp(data_estoque).date() if data_estoque else datetime.fromtimestamp(arq.stat().st_mtime).date()
    e.attrs["data"] = str(dia)
    e.to_pickle(ARQ_ESTOQUE)
    HIST_ESTOQUE.mkdir(parents=True, exist_ok=True)
    e.to_pickle(HIST_ESTOQUE / f"estoque_{dia}.pkl")
    log(f"  {len(e):,} itens com saldo, estoque de {dia:%d/%m/%Y}")
    _registrar("1211", arq.name, f"{dia:%d/%m/%Y}", f"{dia:%d/%m/%Y}", len(e))


IMPORTADORES = {"1174": importar_1174, "1211": importar_1211, "perdidas": importar_perdidas}


def processar_arquivos(arquivos: list[Path], mover: bool):
    # estoque por último não importa; vendas primeiro (mais demorado) para falhar cedo
    ordem = {"1174": 0, "perdidas": 1, "1211": 2}
    tipados = []
    for a in arquivos:
        if a.name.startswith("~$") or not a.is_file():
            continue
        tipo = detectar_tipo(a)
        if tipo is None:
            log(f"  ⚠ {a.name}: não reconhecido como 1174, 1211 ou vendas perdidas — ignorado")
            if mover:
                NAO_RECONHECIDOS.mkdir(parents=True, exist_ok=True)
                shutil.move(str(a), NAO_RECONHECIDOS / a.name)
            continue
        tipados.append((ordem[tipo], a.stat().st_mtime, tipo, a))
    for _, _, tipo, a in sorted(tipados):
        IMPORTADORES[tipo](a)
        if mover:
            PROCESSADOS.mkdir(parents=True, exist_ok=True)
            shutil.move(str(a), PROCESSADOS / f"{datetime.now():%Y-%m-%d_%H%M}_{a.name}")
    return len(tipados)


# ----------------------------------------------------------------- cálculo e saídas
def gerar_saidas(excel=True):
    for arq in (ARQ_VENDAS, ARQ_ESTOQUE, ARQ_PERDIDAS):
        if not arq.exists():
            raise SystemExit(f"Base incompleta: falta {arq.name}. Importe o relatório correspondente.")
    vendas = pd.read_pickle(ARQ_VENDAS)
    estoque = pd.read_pickle(ARQ_ESTOQUE)
    perdidas = pd.read_pickle(ARQ_PERDIDAS)

    log("Rodando o modelo de previsão (backtest + previsão)...")
    R = modelo.rodar(vendas, perdidas, estoque, ARQ_EQUIV, log=log)
    with open(ARQ_MODELO, "wb") as f:
        pickle.dump(R, f, protocol=pickle.HIGHEST_PROTOCOL)
    _guardar_previsao_do_mes(R)

    log("Calculando risco de falta...")
    data_est = pd.Timestamp(estoque.attrs["data"]) if estoque.attrs.get("data") else None
    df, meta = motor.calcular(vendas, estoque, perdidas, data_estoque=data_est, modelo_R=R)
    meta["data_estoque"] = estoque.attrs.get("data", "?")
    meta["vendas_ate"] = vendas["data"].max()
    meta["perdidas_periodo"] = (perdidas["data"].min(), perdidas["data"].max())
    with open(ARQ_PREVISAO, "wb") as f:
        pickle.dump((df, meta), f, protocol=pickle.HIGHEST_PROTOCOL)

    SAIDA.mkdir(exist_ok=True)
    R["equivalencias"].to_csv(SAIDA / "codigos_substituidos.csv", sep=";", index=False, encoding="utf-8-sig")
    html = SAIDA / "Painel_Falta_Pecas.html"
    painel.gerar(df, perdidas, estoque, meta, html)
    log(f"Painel offline: {html.name}")
    if excel:
        log("Gerando planilha Excel (cerca de 1 minuto)...")
        destino = SAIDA / f"Previsao_Falta_Pecas_BETA_{meta['data_estoque_ts']:%Y-%m-%d}.xlsx"
        try:
            relatorio.gerar(df, perdidas, meta, destino)
        except PermissionError:
            destino = destino.with_name(destino.stem + f"_{datetime.now():%H%M}.xlsx")
            log("  (a planilha do dia estava aberta no Excel; gravei com outro nome)")
            relatorio.gerar(df, perdidas, meta, destino)
        log(f"Planilha: {destino.name}")
    log("Pronto ✔")


def _guardar_previsao_do_mes(R):
    """Guarda a 1ª previsão feita para cada mês, para medir a acurácia REAL quando o mês fechar."""
    HIST_PREVISOES.mkdir(parents=True, exist_ok=True)
    arq = HIST_PREVISOES / f"previsao_feita_em_{R['meses'][-1] + 1}.pkl"
    if arq.exists():
        return
    snap = R["info"][["filial", "produto", "classe", "preco"]].copy()
    for h, m in enumerate(R["meses_prev"]):
        snap[str(m)] = R["F"][:, h]
    snap.attrs["feita_em"] = str(pd.Timestamp.now().date())
    snap.to_pickle(arq)
    log(f"  previsão guardada para acompanhamento real: {arq.name}")


def executar(arquivos=None, mover=True, excel=True, log_fn=None):
    global _log_fn
    if log_fn:
        _log_fn = log_fn
    try:
        for p in (BASE, ENTRADA, SAIDA):
            p.mkdir(exist_ok=True)
        if arquivos is None:
            arquivos = sorted(ENTRADA.iterdir())
        n = processar_arquivos(list(arquivos), mover)
        log(f"{n} relatório(s) importado(s)")
        gerar_saidas(excel=excel)
    finally:
        _log_fn = print


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--carga-inicial", action="store_true",
                    help="importa os relatórios da pasta raiz do projeto, sem movê-los")
    ap.add_argument("--so-relatorio", action="store_true")
    ap.add_argument("--sem-excel", action="store_true")
    a = ap.parse_args()
    if a.carga_inicial:
        executar(sorted(RAIZ.iterdir()), mover=False, excel=not a.sem_excel)
    elif a.so_relatorio:
        gerar_saidas(excel=not a.sem_excel)
    else:
        executar(mover=True, excel=not a.sem_excel)


if __name__ == "__main__":
    main()
