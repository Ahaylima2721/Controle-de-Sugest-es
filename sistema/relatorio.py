"""Gera a planilha Excel da previsão (BETA)."""
from __future__ import annotations

import numpy as np
import pandas as pd
from openpyxl.formatting.rule import FormulaRule
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.utils import get_column_letter

import config as C

CORES = {
    "RUPTURA": "F8CBAD", "CRÍTICO": "FCE4D6", "ATENÇÃO": "FFF2CC",
    "OK": "E2EFDA", "SEM GIRO": "EDEDED", "SUBSTITUÍDO": "DDEBF7",
}
AZUL = PatternFill("solid", fgColor="1F4E78")
BRANCO_B = Font(bold=True, color="FFFFFF")

COLUNAS_PREV = [
    ("filial_nome", "Filial", 9, None),
    ("produto", "Produto", 16, None),
    ("descricao", "Descrição", 34, None),
    ("subgrupo", "Grupo", 7, "0"),
    ("status", "Status", 11, None),
    ("recorrencia", "Recorrência", 17, None),
    ("evidencia", "Evidência da demanda", 17, None),
    ("estoque", "Estoque", 9, "#,##0"),
    ("demanda_mes", "Demanda prev./mês", 11, "#,##0.0"),
    ("cobertura_dias", "Cobertura (dias)", 10, "#,##0"),
    ("data_prevista_falta", "Falta prevista em", 12, "dd/mm/yyyy"),
    ("sugestao_compra", "Sugestão compra", 10, "#,##0"),
    ("sobra_outras_filiais", "Sobra em outras filiais (qtd)", 26, None),
    ("valor_em_risco", "Venda em risco (R$)", 13, "#,##0.00"),
    ("vendido_30d", "Vendido 30d", 9, "#,##0"),
    ("vendido_recente", "Vendido 90d", 9, "#,##0"),
    ("vendido_12m", "Vendido 12m", 9, "#,##0"),
    ("vendido_ano_ant_prox", "Ano ant. próx. 45d", 10, "#,##0"),
    ("meses_com_venda_12m", "Meses c/ venda (12m)", 9, "0"),
    ("perdido_qtd", "Venda perdida qtd (90d)", 10, "#,##0"),
    ("perdido_ocorrencias", "Venda perdida ocorr.", 9, "0"),
    ("classe", "Perfil de demanda", 13, None),
    ("metodo", "Método de previsão", 12, None),
    ("acur_item", "Acurácia do item (backtest)", 10, "0%"),
    ("codigos_antigos", "Códigos antigos unidos", 18, None),
    ("ultima_venda", "Última venda", 11, "dd/mm/yyyy"),
    ("dias_sem_venda", "Dias sem venda", 8, "0"),
    ("preco_medio", "Preço médio", 10, "#,##0.00"),
]


def _fmt_tabela(ws, n_linhas, colunas, linha_cab=1):
    for j, (_, titulo, larg, fmt) in enumerate(colunas, start=1):
        c = ws.cell(row=linha_cab, column=j)
        c.fill, c.font = AZUL, BRANCO_B
        c.alignment = Alignment(wrap_text=True, vertical="center", horizontal="center")
        ws.column_dimensions[get_column_letter(j)].width = larg
        if fmt:
            for i in range(linha_cab + 1, linha_cab + 1 + n_linhas):
                ws.cell(row=i, column=j).number_format = fmt
    ws.row_dimensions[linha_cab].height = 45
    ult = get_column_letter(len(colunas))
    ws.auto_filter.ref = f"A{linha_cab}:{ult}{linha_cab + n_linhas}"
    ws.freeze_panes = ws.cell(row=linha_cab + 1, column=4)


def _cores_status(ws, col_status, n, ncol, linha_ini=2):
    ult = get_column_letter(ncol)
    rng = f"A{linha_ini}:{ult}{linha_ini + n - 1}"
    for st, cor in CORES.items():
        ws.conditional_formatting.add(
            rng, FormulaRule(formula=[f'${col_status}{linha_ini}="{st}"'],
                             fill=PatternFill("solid", fgColor=cor)))


def _escrever_prev(xw, df, aba):
    cols = [c for c, *_ in COLUNAS_PREV]
    df = df.copy()
    for c in cols:
        if c not in df:
            df[c] = None
    if "acur_item" in df:
        df["acur_item"] = pd.to_numeric(df["acur_item"], errors="coerce").clip(lower=0)
    out = df[cols].copy()
    out["cobertura_dias"] = out["cobertura_dias"].replace(np.inf, np.nan).round(0)
    out.columns = [t for _, t, *_ in COLUNAS_PREV]
    out.to_excel(xw, sheet_name=aba, index=False)
    ws = xw.sheets[aba]
    _fmt_tabela(ws, len(out), COLUNAS_PREV)
    _cores_status(ws, "E", len(out), len(COLUNAS_PREV))


def gerar(df: pd.DataFrame, perdidas: pd.DataFrame, meta: dict, destino):
    hoje = meta["data_base"]
    alerta = df[df["status"].isin(["RUPTURA", "CRÍTICO", "ATENÇÃO"])]
    alerta_rec = alerta[alerta["recorrencia"] == "RECORRENTE"]

    with pd.ExcelWriter(destino, engine="openpyxl") as xw:
        # ------------------------------------------------ Resumo
        ordem = ["RUPTURA", "CRÍTICO", "ATENÇÃO", "OK", "SEM GIRO"]
        base = df[df["recorrencia"] == "RECORRENTE"]
        cnt = base.pivot_table(index="filial_nome", columns="status", values="produto",
                               aggfunc="count", fill_value=0).reindex(columns=ordem, fill_value=0)
        risco = base[base["status"].isin(ordem[:3])].groupby("filial_nome")["valor_em_risco"].sum()
        vp = perdidas[perdidas["data"] > hoje - pd.Timedelta(days=meta["janela"])]
        vp_f = vp.groupby("filial")["valor"].sum()
        vp_f.index = [f"{f:02d} {C.NOMES_FILIAIS.get(f, '')}" for f in vp_f.index]
        resumo = cnt.copy()
        resumo["Venda em risco (R$)"] = risco
        resumo["Vendas perdidas 90d (R$)"] = vp_f
        resumo = resumo.fillna(0)
        resumo.loc["TOTAL"] = resumo.sum()
        resumo.index.name = "Filial"
        resumo = resumo.reset_index()

        topo = [
            ["PREVISÃO DE FALTA DE PEÇAS — VERSÃO BETA"],
            [f"Data-base (última venda no 1174): {hoje:%d/%m/%Y}   |   Estoque (1211) de: "
             f"{meta['data_estoque']}   |   Vendas perdidas: "
             f"{meta['perdidas_periodo'][0]:%d/%m/%Y} a {meta['perdidas_periodo'][1]:%d/%m/%Y}"],
            [f"Parâmetros: lead time {C.LEAD_TIME_DIAS}d, segurança {C.SEGURANCA_DIAS}d, "
             f"ciclo de compra {C.CICLO_COMPRA_DIAS}d, janela recente {C.JANELA_RECENTE_DIAS}d, "
             f"peso sazonal {C.PESO_SAZONAL:.0%}"],
            ["Contagem de itens RECORRENTES (vendidos em ≥ "
             f"{C.MIN_MESES_RECORRENCIA} meses nos últimos 12, ou com ≥2 vendas perdidas)."],
        ]
        pd.DataFrame(topo).to_excel(xw, sheet_name="Resumo", index=False, header=False)
        resumo.to_excel(xw, sheet_name="Resumo", index=False, startrow=5)
        ws = xw.sheets["Resumo"]
        ws["A1"].font = Font(bold=True, size=14, color="1F4E78")
        cols_res = [(c, c, 14, "#,##0") for c in resumo.columns]
        cols_res[0] = ("Filial", "Filial", 12, None)
        for j in (len(cols_res) - 2, len(cols_res) - 1):
            cols_res[j] = (resumo.columns[j], resumo.columns[j], 18, "#,##0.00")
        for j, (_, _, larg, fmt) in enumerate(cols_res, start=1):
            c = ws.cell(row=6, column=j)
            c.fill, c.font = AZUL, BRANCO_B
            c.alignment = Alignment(wrap_text=True, horizontal="center")
            ws.column_dimensions[get_column_letter(j)].width = larg
            for i in range(7, 7 + len(resumo)):
                if fmt:
                    ws.cell(row=i, column=j).number_format = fmt
        for j in range(1, len(cols_res) + 1):
            ws.cell(row=6 + len(resumo), column=j).font = Font(bold=True)
        for st, cor in CORES.items():
            if st in resumo.columns:
                j = list(resumo.columns).index(st) + 1
                ws.cell(row=6, column=j).fill = PatternFill("solid", fgColor=cor)
                ws.cell(row=6, column=j).font = Font(bold=True)

        # legenda
        lin = 9 + len(resumo)
        leg = [
            ("RUPTURA", "Sem estoque e com demanda (venda recente, sazonal ou venda perdida)."),
            ("CRÍTICO", f"Estoque acaba antes do lead time ({C.LEAD_TIME_DIAS} dias): já não dá tempo de repor."),
            ("ATENÇÃO", f"Estoque acaba antes de {C.LEAD_TIME_DIAS + C.SEGURANCA_DIAS} dias (lead time + segurança): pedir agora."),
            ("OK", "Cobertura acima do horizonte de proteção."),
            ("SEM GIRO", "Tem estoque mas nenhuma demanda prevista (candidato a transferência)."),
        ]
        ws.cell(row=lin, column=1, value="Legenda").font = Font(bold=True)
        for i, (st, txt) in enumerate(leg, start=1):
            c = ws.cell(row=lin + i, column=1, value=st)
            c.fill = PatternFill("solid", fgColor=CORES[st])
            c.font = Font(bold=True)
            ws.cell(row=lin + i, column=2, value=txt)

        # ------------------------------------------------ Alertas (recorrentes)
        _escrever_prev(xw, alerta_rec, "Alertas recorrentes")
        # ------------------------------------------------ Alertas baixa recorrência
        _escrever_prev(xw, alerta[alerta["recorrencia"] != "RECORRENTE"], "Alertas baixa recorr.")
        # ------------------------------------------------ Transferências
        tr = alerta_rec[(alerta_rec["sobra_rede_qtd"] > 0)]
        _escrever_prev(xw, tr, "Transferências possíveis")
        # ------------------------------------------------ Base completa
        _escrever_prev(xw, df, "Todos os itens")

        # ------------------------------------------------ Vendas perdidas x estoque hoje
        k = ["filial", "produto"]
        vpa = vp.groupby(k).agg(ocorrencias=("qtd", "size"), qtd=("qtd", "sum"),
                                valor=("valor", "sum"), primeira=("data", "min"),
                                ultima=("data", "max"), descricao=("descricao", "first")).reset_index()
        vpa = vpa.merge(df[k + ["estoque", "status", "vendido_12m", "sobra_outras_filiais"]],
                        on=k, how="left")
        vpa["filial"] = vpa["filial"].map(lambda f: f"{f:02d} {C.NOMES_FILIAIS.get(f, '')}")
        vpa["situacao_hoje"] = np.where(vpa["estoque"].fillna(0) > 0, "JÁ TEM ESTOQUE",
                                        "CONTINUA SEM ESTOQUE")
        vpa = vpa.sort_values(["situacao_hoje", "valor"], ascending=[False, False])
        cols_vp = [
            ("filial", "Filial", 9, None), ("produto", "Produto", 16, None),
            ("descricao", "Descrição", 34, None), ("situacao_hoje", "Situação hoje", 20, None),
            ("ocorrencias", "Ocorrências", 10, "0"), ("qtd", "Qtd perdida", 9, "#,##0"),
            ("valor", "Valor perdido (R$)", 13, "#,##0.00"),
            ("primeira", "Primeira", 11, "dd/mm/yyyy"), ("ultima", "Última", 11, "dd/mm/yyyy"),
            ("estoque", "Estoque hoje", 9, "#,##0"), ("vendido_12m", "Vendido 12m", 9, "#,##0"),
            ("status", "Status previsão", 11, None),
            ("sobra_outras_filiais", "Sobra em outras filiais", 26, None),
        ]
        o = vpa[[c for c, *_ in cols_vp]].copy()
        o.columns = [t for _, t, *_ in cols_vp]
        o.to_excel(xw, sheet_name="Vendas perdidas x estoque", index=False)
        ws = xw.sheets["Vendas perdidas x estoque"]
        _fmt_tabela(ws, len(o), cols_vp)
        ws.conditional_formatting.add(
            f"A2:{get_column_letter(len(cols_vp))}{len(o)+1}",
            FormulaRule(formula=['$D2="CONTINUA SEM ESTOQUE"'],
                        fill=PatternFill("solid", fgColor=CORES["RUPTURA"])))

        # ------------------------------------------------ Metodologia
        met = [
            "COMO A PREVISÃO É CALCULADA (por filial + produto)",
            "",
            "1-3) Demanda prevista = MODELO DE PREVISÃO (sistema\\modelo.py): previsão mensal por filial × produto.",
            "     Vários métodos competem num backtest dos últimos 9 meses; para cada perfil de demanda",
            "     (coluna 'Perfil de demanda') fica o método de menor erro com viés de até ±30%.",
            "     Demanda = venda líquida + garantia + vendas perdidas; picos isolados são cortados;",
            "     códigos substituídos (ex.: sufixo DSH/OF/FP) têm o histórico unido ao código novo.",
            "     Acurácia detalhada: tela 'Acurácia' do aplicativo (INICIAR.bat).",
            "4) Cobertura (dias) = estoque atual / demanda prevista por dia.",
            f"5) Horizonte de proteção = lead time ({C.LEAD_TIME_DIAS}) + segurança ({C.SEGURANCA_DIAS}) = {meta['horizonte']} dias.",
            f"6) Sugestão de compra = demanda/dia × ({meta['horizonte']} + ciclo {C.CICLO_COMPRA_DIAS}) − estoque.",
            "7) Venda em risco = (demanda no horizonte − estoque) × preço médio de venda 12m.",
            "8) Sobra em outras filiais = estoque que excede a necessidade da própria filial no horizonte.",
            "9) Evidência da demanda: VENDA PERDIDA (pedido recusado nos últimos 90d) > VENDA RECENTE (90d) > "
            "SÓ SAZONAL (vendeu no mesmo período do ano passado). 'Só sazonal' com muitos dias sem venda "
            "pode ser item substituído/descontinuado — conferir antes de comprar.",
            "",
            "Vendas líquidas: lançamentos 'Vendas' somam, demais tipos (devoluções) abatem.",
            "O 1211 lista apenas itens com saldo > 0; item ausente no 1211 = estoque zero.",
            "",
            "LIMITAÇÕES CONHECIDAS DO BETA",
            "- Lead time é único para todos os itens (não há dado de fornecedor/marca ainda).",
            "- Não considera pedidos de compra em aberto / mercadoria em trânsito.",
            "- Códigos com sufixo (OF, BR, PA...) são tratados como produtos distintos.",
            "- Vendas perdidas só existem a partir de 01/09/2026; antes disso a demanda reprimida não é medida.",
            "- Itens muito esporádicos (BAIXA RECORRÊNCIA) ficam em aba separada para não poluir os alertas.",
            "",
            "Parâmetros ficam em sistema\\config.py.",
        ]
        pd.DataFrame({"Metodologia": met}).to_excel(xw, sheet_name="Metodologia", index=False)
        ws = xw.sheets["Metodologia"]
        ws.column_dimensions["A"].width = 120
        ws["A1"].fill, ws["A1"].font = AZUL, BRANCO_B
