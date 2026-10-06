"""
utils/relatorio.py — Geração do relatório executivo (HTML) e do workbook de
acompanhamento mensal (XLSX) a partir de um ComparisonResult.

É chamada pelo app no final de cada processamento. Produz dois artefatos em
memória (sem gravar disco):

  • build_html_report(result, ...)  → str (HTML autocontido, gráficos embutidos)
  • build_xlsx_workbook(result, ...) → bytes (workbook editável do mês)

Os dois consomem o mesmo ComparisonResult — a fonte da verdade continua
sendo o cálculo feito em comparator.py.
"""

from __future__ import annotations

import base64
import io
import re
import unicodedata
import warnings
from datetime import datetime
from typing import Optional

import pandas as pd

warnings.filterwarnings("ignore")

import matplotlib
matplotlib.use("Agg")
import matplotlib.pyplot as plt
import openpyxl
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.worksheet.datavalidation import DataValidation

from .comparator import ComparisonResult, result_to_dataframe
from .parser import _format_brl

plt.rcParams.update({"font.size": 10, "figure.dpi": 110})

AZUL = "#0F2A5C"
LAR = "#FF6B1A"
VERDE = "#1a8a4a"
CINZA = "#6b7891"
ROXO = "#6c3483"

CONTRATOS_AGRUPA = [
    "PMV", "SEDU", "SESA", "CETURB", "INOVA", "UNIMED", "IASES", "PCES",
    "POLICIA CIVIL", "POLICIA FEDERAL", "POLICIA MILITAR", "SEME",
    "SEMPDEC", "SEMESP", "SEMSU", "ECT", "CESAN", "ORLA", "ADM",
]

REGRAS_DESCONTO_TEXTO = (
    "• FALTA 30 DIAS — desconto do mês inteiro de VT (21–22 dias úteis).<br>"
    "• ATESTADO MÉDICO — desconto proporcional aos dias afastados.<br>"
    "• FALTA INJUSTIFICADA — desconto dos dias de falta.<br>"
    "• FÉRIAS — regra geral: desconta o VT do período; <b>exceção: vigilantes "
    "(VSP) NÃO descontam VT em férias</b>.<br>"
    "• Cobertura 'Sem Cobertura' predomina — verificar se o posto ficou "
    "descoberto ou se houve cobertura informal."
)


def extrai_contrato(posto) -> str:
    if posto is None or (isinstance(posto, float) and pd.isna(posto)):
        return "(sem posto)"
    p = str(posto).strip().upper()
    if not p:
        return "(sem posto)"
    for c in sorted(CONTRATOS_AGRUPA, key=len, reverse=True):
        if p == c or p.startswith(c + " ") or p.startswith(c + "-") or p.startswith(c + " -"):
            return c
    if " - " in p:
        return p.split(" - ")[0].strip()
    return p


def _norm_name(s) -> str:
    if s is None:
        return ""
    s = unicodedata.normalize("NFD", str(s))
    return re.sub(r"\s+", " ", "".join(c for c in s if unicodedata.category(c) != "Mn")).upper().strip()


def _brl(v: float) -> str:
    return f"R$ {v:,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")


# ---------------------------------------------------------------------------
# Parser do relatório de ocorrências (AppLider — "Relatorio_Ocorrencias…")
# ---------------------------------------------------------------------------

_OCOR_COLS = [
    "PARCEIROID", "NOME", "TIPO_ID", "DATA_OC", "MOTIVO_C", "ATIVO", "COD_COB",
    "NOME_COB", "DESCRICAO", "EMPRESA", "ESCALA", "POSTO", "DATA_INI",
    "DIAS_AFAST", "DATA_FIM", "TIPO_AUSENCIA", "MEDICO", "CRM", "CID", "ANO",
    "MES_ANO", "TIPO_COB", "DIAS_DESC",
]


def parse_ocorrencias(content: bytes, filename: str) -> Optional[pd.DataFrame]:
    """
    Lê o relatório de ocorrências do AppLider (export .xls ou .xlsx).
    Esperado: header na linha 4 (índice 3), colunas PARCEIROID … DIAS DESC.
    Se o arquivo não tiver esse formato, devolve None (o app avisa e segue sem).
    """
    try:
        if filename.lower().endswith(".xlsx"):
            df = pd.read_excel(io.BytesIO(content), header=3, engine="openpyxl", dtype=object)
        else:
            df = pd.read_excel(io.BytesIO(content), header=3, engine="xlrd", dtype=object)
    except Exception:
        return None
    if df.shape[1] < 5:
        return None
    # remove rodapé "TOTAL DE REGISTROS: …"
    df = df[~df.iloc[:, 0].astype(str).str.contains("TOTAL DE REGISTROS", na=False)]
    if df.empty:
        return None
    n_cols = min(df.shape[1], len(_OCOR_COLS))
    df = df.iloc[:, :n_cols]
    df.columns = _OCOR_COLS[:n_cols]
    for col in _OCOR_COLS:
        if col not in df.columns:
            df[col] = ""
    df["NOME_N"] = df["NOME"].map(_norm_name)
    df["DIAS_DESC_N"] = pd.to_numeric(df["DIAS_DESC"], errors="coerce").fillna(0)
    df["DESCRICAO_N"] = df["DESCRICAO"].astype(str).str.upper().str.strip()
    return df


# ---------------------------------------------------------------------------
# Montagem do dataframe "enriquecido" usado pelos dois relatórios
# ---------------------------------------------------------------------------

def _enrich_dataframe(
    result: ComparisonResult,
    ocorrencias_df: Optional[pd.DataFrame],
) -> pd.DataFrame:
    df = result_to_dataframe(result)
    # o dataframe do comparator chama a coluna de "Empresa"; os relatórios usam
    # "EmpresaGrupo" — padroniza aqui uma única vez para os dois geradores
    if "EmpresaGrupo" not in df.columns and "Empresa" in df.columns:
        df = df.rename(columns={"Empresa": "EmpresaGrupo"})
    df["Contrato"] = df["Posto"].map(extrai_contrato)

    if ocorrencias_df is not None and not ocorrencias_df.empty:
        agg = (
            ocorrencias_df.groupby("NOME_N")
            .agg(
                n_oc=("DESCRICAO_N", "count"),
                dias_desc=("DIAS_DESC_N", "sum"),
                motivos=("DESCRICAO_N", lambda s: ", ".join(sorted(set(
                    x for x in s if x and x != "NAN")))),
            )
            .reset_index()
        )
        agg_map = {r["NOME_N"]: r for _, r in agg.iterrows()}
        df["NOME_N"] = df["Nome"].map(_norm_name)
        df["Ocorrencias_resumo"] = df["NOME_N"].map(
            lambda n: "" if agg_map.get(n) is None else
            f"{agg_map[n]['motivos']} ({int(agg_map[n]['n_oc'])} ocorr.)")
        df["Dias_descontados_oc"] = df["NOME_N"].map(
            lambda n: 0 if agg_map.get(n) is None else int(agg_map[n]["dias_desc"]))
    else:
        df["Ocorrencias_resumo"] = ""
        df["Dias_descontados_oc"] = 0
    return df


# ---------------------------------------------------------------------------
# Gráficos (base64 para embutir no HTML)
# ---------------------------------------------------------------------------

def _fig_b64(fig) -> str:
    buf = io.BytesIO()
    fig.savefig(buf, format="png", bbox_inches="tight", facecolor="white")
    plt.close(fig)
    return base64.b64encode(buf.getvalue()).decode()


def _charts(df: pd.DataFrame, resumo_emp: pd.DataFrame,
            oc_mot: Optional[pd.DataFrame]) -> list:
    charts = []

    fig, ax = plt.subplots(figsize=(9, 4.2))
    x = range(len(resumo_emp))
    ax.bar([i - .2 for i in x], resumo_emp["Total_TXT"], width=.4, label="TXT bruto", color=CINZA)
    ax.bar([i + .2 for i in x], resumo_emp["A_Depositar"], width=.4, label="A depositar", color=LAR)
    for i, r in resumo_emp.iterrows():
        ax.text(i + .2, r["A_Depositar"], _brl(r["A_Depositar"]), ha="center",
                fontsize=8, color=AZUL, fontweight="bold")
    ax.set_xticks(list(x))
    ax.set_xticklabels(resumo_emp["EmpresaGrupo"], rotation=8)
    ax.set_title("TXT vs A depositar — por empresa", color=AZUL, fontweight="bold")
    ax.legend(); ax.grid(axis="y", alpha=.25)
    charts.append(("TXT vs Depósito por empresa", _fig_b64(fig)))

    fig, ax = plt.subplots(figsize=(9, 3.6))
    bars = ax.barh(resumo_emp["EmpresaGrupo"], resumo_emp["Economia_%"], color=VERDE)
    for b, v in zip(bars, resumo_emp["Economia_%"]):
        ax.text(b.get_width() + .5, b.get_y() + b.get_height() / 2, f"{v}%",
                va="center", fontweight="bold", color=AZUL)
    ax.set_xlim(0, max(resumo_emp["Economia_%"]) * 1.25 if len(resumo_emp) else 1)
    ax.set_title("% de redução vs TXT (saldo absorvendo)", color=AZUL, fontweight="bold")
    ax.grid(axis="x", alpha=.25)
    charts.append(("Economia % por empresa", _fig_b64(fig)))

    contrato = df.groupby("Contrato").agg(
        Qtde=("Matrícula", "count"), TXT=("Valor TXT (R$)", "sum"),
        Saldo=("Saldo PDF (R$)", "sum"), Depositar=("A depositar (R$)", "sum")
    ).reset_index()
    contrato["Economia"] = contrato["TXT"] - contrato["Depositar"]
    top_c = contrato.sort_values("Depositar", ascending=False).head(14)
    fig, ax = plt.subplots(figsize=(9, 5.6))
    bars = ax.barh(top_c["Contrato"].astype(str), top_c["Depositar"], color=AZUL)
    ax.invert_yaxis()
    for b, v, q, e in zip(bars, top_c["Depositar"], top_c["Qtde"], top_c["Economia"]):
        ax.text(b.get_width() + max(top_c["Depositar"]) * .01,
                b.get_y() + b.get_height() / 2,
                f"{_brl(v)} · {q} pessoas · eco. {_brl(e)}",
                va="center", fontsize=8, color=CINZA)
    ax.set_xlim(0, max(top_c["Depositar"]) * 1.8 if len(top_c) else 1)
    ax.set_title("Top contratos por valor a depositar", color=AZUL, fontweight="bold")
    ax.grid(axis="x", alpha=.25)
    charts.append(("Top contratos", _fig_b64(fig)))

    if oc_mot is not None and not oc_mot.empty:
        fig, ax = plt.subplots(figsize=(9, 5.2))
        tm = oc_mot.head(12)
        bars = ax.barh(tm["DESCRICAO_N"].astype(str).str[:40],
                       tm["Dias_descontados"], color=ROXO)
        ax.invert_yaxis()
        for b, v, n in zip(bars, tm["Dias_descontados"], tm["Registros"]):
            ax.text(b.get_width() + max(tm["Dias_descontados"]) * .01,
                    b.get_y() + b.get_height() / 2,
                    f"{int(v)} dias · {int(n)} registros", va="center",
                    fontsize=8, color=CINZA)
        ax.set_xlim(0, max(tm["Dias_descontados"]) * 1.35 if len(tm) else 1)
        ax.set_title("Dias descontados por motivo de ocorrência",
                     color=AZUL, fontweight="bold")
        ax.grid(axis="x", alpha=.25)
        charts.append(("Dias descontados por motivo", _fig_b64(fig)))

    fig, ax = plt.subplots(figsize=(8, 4))
    sc = df["Escala"].value_counts()
    bars = ax.bar(sc.index.astype(str), sc.values,
                  color=[LAR, AZUL, VERDE, "#8e44ad", "#f39c12", "#95a5a6"][:len(sc)])
    for b, v in zip(bars, sc.values):
        ax.text(b.get_x() + b.get_width() / 2, v, str(int(v)),
                ha="center", fontweight="bold", color=AZUL)
    ax.set_title("Distribuição de escalas (geral)", color=AZUL, fontweight="bold")
    ax.grid(axis="y", alpha=.25); plt.xticks(rotation=12)
    charts.append(("Escalas", _fig_b64(fig)))

    status_counts = df["Status"].value_counts()
    order = ["✅ Saldo cobre 100%", "💰 Complemento", "—  Já 0 no TXT",
             "🆕 Sem saldo no cartão", "🖐 Revisão manual (2x2)", "⚠️ Sem escala (revisar)"]
    order = [s for s in order if s in status_counts.index] + \
            [s for s in status_counts.index if s not in order]
    colors_st = ["#1a8a4a", LAR, "#b7c3d6", "#e67e22", "#f39c12", "#b32a2a"]
    piv = df.pivot_table(index="EmpresaGrupo", columns="Status",
                         values="Matrícula", aggfunc="count", fill_value=0)
    fig, ax = plt.subplots(figsize=(9, 4.2))
    bottom = [0] * len(piv)
    for s, c in zip(order, colors_st):
        vals = piv[s].values if s in piv.columns else [0] * len(piv)
        ax.bar(piv.index, vals, bottom=bottom, label=s, color=c)
        bottom = [bb + v for bb, v in zip(bottom, vals)]
    ax.set_title("Status por empresa", color=AZUL, fontweight="bold")
    ax.legend(fontsize=8, loc="upper right")
    ax.grid(axis="y", alpha=.25); plt.xticks(rotation=8)
    charts.append(("Status por empresa", _fig_b64(fig)))

    return charts


# ---------------------------------------------------------------------------
# HTML
# ---------------------------------------------------------------------------

def build_html_report(
    result: ComparisonResult,
    ocorrencias_df: Optional[pd.DataFrame] = None,
    *,
    titulo_extra: str = "",
) -> str:
    df = _enrich_dataframe(result, ocorrencias_df)

    resumo_emp = df.groupby("EmpresaGrupo" if "EmpresaGrupo" in df.columns else "Empresa").agg(
        Colaboradores=("Matrícula", "count"),
        Total_TXT=("Valor TXT (R$)", "sum"),
        Saldo_PDF=("Saldo PDF (R$)", "sum"),
        A_Depositar=("A depositar (R$)", "sum"),
    ).reset_index().rename(columns={"Empresa": "EmpresaGrupo"})
    resumo_emp["Economia_R$"] = resumo_emp["Total_TXT"] - resumo_emp["A_Depositar"]
    resumo_emp["Economia_%"] = (resumo_emp["Economia_R$"] /
                                 resumo_emp["Total_TXT"].replace(0, pd.NA) * 100).round(1)

    contrato = df.groupby("Contrato").agg(
        Qtde=("Matrícula", "count"), TXT=("Valor TXT (R$)", "sum"),
        Saldo=("Saldo PDF (R$)", "sum"), Depositar=("A depositar (R$)", "sum")
    ).reset_index()
    contrato["Economia"] = contrato["TXT"] - contrato["Depositar"]
    contrato = contrato.sort_values("Depositar", ascending=False)

    status_counts = df["Status"].value_counts()
    escala_counts = df["Escala"].value_counts()
    TOTAL = {
        "colab": len(df), "txt": df["Valor TXT (R$)"].sum(),
        "dep": df["A depositar (R$)"].sum(), "saldo": df["Saldo PDF (R$)"].sum(),
        "consumo_atual": df["Consumo mês atual (R$)"].sum(),
        "ajustado": df["Saldo ajustado (R$)"].sum(),
    }
    TOTAL["eco"] = TOTAL["txt"] - TOTAL["dep"]
    TOTAL["eco_pct"] = TOTAL["eco"] / TOTAL["txt"] * 100 if TOTAL["txt"] else 0

    oc_stats = None
    oc_emp = None
    top_pess = None
    if ocorrencias_df is not None and not ocorrencias_df.empty:
        oc_stats = {
            "registros": len(ocorrencias_df),
            "dias": int(ocorrencias_df["DIAS_DESC_N"].sum()),
            "pessoas": int(ocorrencias_df["NOME_N"].nunique()),
        }
        oc_mot = (
            ocorrencias_df.groupby("DESCRICAO_N")
            .agg(Registros=("DESCRICAO_N", "count"), Pessoas=("NOME_N", "nunique"),
                 Dias_descontados=("DIAS_DESC_N", "sum"))
            .reset_index()
            .sort_values("Dias_descontados", ascending=False)
        )
        oc_mot = oc_mot[oc_mot["DESCRICAO_N"].ne("NAN")]
        oc_emp = (
            ocorrencias_df.groupby("EmpresaGrupo" if "EmpresaGrupo" in ocorrencias_df.columns
                                    else "EMPRESA")
            .agg(Registros=("DESCRICAO_N", "count"),
                 Dias_descontados=("DIAS_DESC_N", "sum"))
            .reset_index().sort_values("Dias_descontados", ascending=False)
        )
        top_pess = (
            ocorrencias_df.groupby(["NOME_N"])
            .agg(Ocorrencias=("DESCRICAO_N", "count"),
                 Dias_descontados=("DIAS_DESC_N", "sum"),
                 Motivos=("DESCRICAO_N", lambda s: ", ".join(sorted(set(
                     x for x in s if x and x != "NAN")))))
            .reset_index().sort_values("Dias_descontados", ascending=False).head(20)
        )
    else:
        oc_mot = None

    n_oc_match = int((df["Dias_descontados_oc"] > 0).sum())

    inc = []
    inc.append(("Colaboradores SEM ESCALA (não achados no AppLider)",
                df[df["Status"].str.contains("Sem escala", na=False)],
                "Sem escala o app não calcula dias do período — sai valor cheio do TXT sem abater saldo. Maior risco de pagamento a maior."))
    inc.append(("Colaboradores SEM SALDO no cartão (matrícula ausente no PDF do GVBUS)",
                df[df["Status"].str.contains("Sem saldo no cartão", na=False)],
                "Matrícula do TXT não existe no relatório de saldo. Pode ser cartão novo/substituído/bloqueado. Depósito sai cheio — conferir no portal GVBUS."))
    inc.append(("Escala 2x2 (revisão manual)",
                df[df["Status"].str.contains("2x2", na=False)],
                "Por desenho o app não calcula 2x2 automaticamente. Revisar manualmente o consumo do período."))
    inc.append(("Regra especial CETURB (1 vale/dia)",
                df[df["Regra especial"].notna() & (df["Regra especial"] != "")],
                "Consumo calculado com R$ 5,10/dia em vez de R$ 10,20. Listados para conferência do conjunto."))
    inc.append(("TXT zerado MAS saldo alto no cartão (>R$100)",
                df[(df["Valor TXT (R$)"] == 0) & (df["Saldo PDF (R$)"] > 100)],
                "Férias/afastamento/FLT zeraram o TXT, mas há saldo parado considerável. Monitorar uso."))
    if ocorrencias_df is not None:
        inc.append(("Colaboradores com ocorrências que descontaram dias",
                    df[df["Dias_descontados_oc"] > 0],
                    "Tiveram atestado/falta/afastamento com dias descontados conforme o relatório de ocorrências — o valor do TXT já veio reduzido por isso. Detalhe na aba Ocorrências do workbook."))

    def _tbl(d_, money=(), pct=(), max_rows=None):
        d = d_.copy()
        if max_rows: d = d.head(max_rows)
        h = '<table><thead><tr>' + "".join(f"<th>{c}</th>" for c in d.columns) + "</tr></thead><tbody>"
        for _, r in d.iterrows():
            h += "<tr>"
            for c in d.columns:
                v = r[c]
                if c in money: v = _brl(float(v))
                elif c in pct: v = f"{float(v):.1f}%"
                elif pd.isna(v) or v == "": v = "—"
                h += f"<td>{v}</td>"
            h += "</tr>"
        return h + "</tbody></table>"

    inc_html = ""
    for titulo, dfi, expl in inc:
        n = len(dfi)
        cls = "ok" if n == 0 else ("warn" if "ocorrênc" in titulo.lower() or "revis" in titulo.lower() else "alert")
        inc_html += (f'<div class="inc {cls}"><div class="inc-head">'
                     f'<span class="inc-n">{n}</span><h3>{titulo}</h3></div>'
                     f'<p class="why"><b>Por que acontece / impacto:</b> {expl}</p>')
        if n > 0:
            cols = [c for c in ["Matrícula", "Nome", "EmpresaGrupo", "Contrato", "Posto",
                                "Escala", "Ocorrencias_resumo", "Dias_descontados_oc",
                                "Valor TXT (R$)", "Saldo PDF (R$)", "A depositar (R$)",
                                "OBS", "Status"] if c in dfi.columns]
            inc_html += _tbl(dfi[cols], money=("Valor TXT (R$)", "Saldo PDF (R$)", "A depositar (R$)"), max_rows=10)
            if n > 10:
                inc_html += (f"<details class='more-all'><summary>▶ Ver todas as {n} linhas "
                             f"(clique para expandir)</summary>"
                             + _tbl(dfi[cols], money=("Valor TXT (R$)", "Saldo PDF (R$)", "A depositar (R$)"))
                             + "</details>")
        inc_html += "</div>"

    # dataset completo para tabela interativa
    FT_COLS = ["Matrícula", "Nome", "EmpresaGrupo", "Contrato", "Posto", "Escala",
               "Status", "Valor TXT (R$)", "Saldo PDF (R$)", "Saldo ajustado (R$)",
               "A depositar (R$)", "Ocorrencias_resumo", "Dias_descontados_oc", "OBS"]
    ft_rows = []
    for _, r in df.iterrows():
        row = []
        for c in FT_COLS:
            v = r.get(c, "")
            if pd.isna(v): v = ""
            elif isinstance(v, float) and c.endswith("(R$)"): v = round(v, 2)
            elif isinstance(v, float): v = int(v) if v == int(v) else v
            row.append(v)
        ft_rows.append(row)
    import json
    ft_json = json.dumps(ft_rows, ensure_ascii=False)
    ft_headers = json.dumps(FT_COLS, ensure_ascii=False)
    ft_emps = json.dumps(sorted(df["EmpresaGrupo"].unique().tolist()), ensure_ascii=False)
    ft_ctrs = json.dumps(sorted(df["Contrato"].astype(str).unique().tolist()), ensure_ascii=False)
    ft_sts = json.dumps(sorted(df["Status"].astype(str).unique().tolist()), ensure_ascii=False)
    ft_escs = json.dumps(sorted(df["Escala"].astype(str).unique().tolist()), ensure_ascii=False)

    charts = _charts(df, resumo_emp, oc_mot)
    charts_html = "".join(
        f'<div class="chart"><h3>{t}</h3><img src="data:image/png;base64,{b}"/></div>'
        for t, b in charts)

    oc_sec_html = ""
    if oc_stats is not None:
        oc_sec_html = f"""
<h2>4 · Ocorrências — por que cada pessoa foi descontada</h2>
<div class="note"><b>📖 Regras de desconto identificadas (conferir com RH):</b><br>{REGRAS_DESCONTO_TEXTO}</div>
<h3>4.1 · Dias descontados por motivo</h3>{_tbl(oc_mot)}
<h3>4.2 · Ocorrências por empresa</h3>{_tbl(oc_emp)}
<h3>4.3 · Top 20 pessoas com mais dias descontados</h3>{_tbl(top_pess)}
"""
    else:
        oc_sec_html = """
<h2>4 · Ocorrências</h2>
<div class="note" style="background:#fdeaea;border-color:#f5b7b1"><b>⚠️ Relatório de ocorrências não anexado.</b>
O relatório foi gerado sem o cruzamento com as ocorrências do mês (atestados, faltas, afastamentos…).
Os valores abaixo refletem apenas o cálculo do VT. Reprocesse anexando o relatório de ocorrências para ter
o motivo do desconto de cada pessoa.</div>
"""

    full_table = f"""
<h2>7 · Dados completos (tabela interativa — todas as linhas)</h2>
<div class="note">Todos os <b>{len(df)} colaboradores</b> direto no navegador: busca livre, filtros por empresa/contrato/status/escala, ordenação por coluna e paginação — sem abrir Excel.</div>
<div class="fulltable-toolbar">
  <input id="ftSearch" type="text" placeholder="🔍 Buscar por nome, matrícula, posto, OBS…" oninput="ftApply()"/>
  <select id="ftEmp" onchange="ftApply()"><option value="">Empresa: todas</option></select>
  <select id="ftCtr" onchange="ftApply()"><option value="">Contrato: todos</option></select>
  <select id="ftSt" onchange="ftApply()"><option value="">Status: todos</option></select>
  <select id="ftEsc" onchange="ftApply()"><option value="">Escala: todas</option></select>
  <span class="ft-count" id="ftCount"></span>
</div>
<div class="ft-wrap"><table id="fullTable"><thead></thead><tbody></tbody></table></div>
<div class="ft-pager">
  <button onclick="ftPage(-1)" id="ftPrev">◀ Anterior</button>
  <span class="pg" id="ftPg"></span>
  <button onclick="ftPage(1)" id="ftNext">Próxima ▶</button>
</div>
<script>
const FT_HEADERS = {ft_headers};
const FT_DATA = {ft_json};
const FT_EMP = {ft_emps}, FT_CTR = {ft_ctrs}, FT_ST = {ft_sts}, FT_ESC = {ft_escs};
const MONEY = new Set(["Valor TXT (R$)","Saldo PDF (R$)","Saldo ajustado (R$)","A depositar (R$)"]);
const PAGE = 50;
let ftFiltered = FT_DATA, ftPageI = 0, ftSortCol = -1, ftSortAsc = true;
function ftFill(id, arr) {{ const s = document.getElementById(id); for (const v of arr) {{ const o = document.createElement("option"); o.value = v; o.textContent = v; s.appendChild(o); }} }}
ftFill("ftEmp", FT_EMP); ftFill("ftCtr", FT_CTR); ftFill("ftSt", FT_ST); ftFill("ftEsc", FT_ESC);
function fmtBRL(v) {{ return "R$ " + Number(v).toLocaleString("pt-BR", {{minimumFractionDigits:2, maximumFractionDigits:2}}); }}
function ftApply() {{
  const q = document.getElementById("ftSearch").value.toLowerCase().trim();
  const e = document.getElementById("ftEmp").value, c = document.getElementById("ftCtr").value;
  const s = document.getElementById("ftSt").value, x = document.getElementById("ftEsc").value;
  ftFiltered = FT_DATA.filter(r => (!e || r[2] === e) && (!c || r[3] === c) && (!s || r[6] === s) && (!x || r[5] === x) && (!q || r.join(" ").toLowerCase().includes(q)));
  if (ftSortCol >= 0) ftSort(ftSortCol, ftSortAsc, false);
  ftPageI = 0; ftRender();
}}
function ftSort(col, asc, reapply=true) {{
  ftSortCol = col; ftSortAsc = asc;
  ftFiltered.sort((a, b) => {{
    let va = a[col], vb = b[col];
    if (MONEY.has(FT_HEADERS[col]) || FT_HEADERS[col] === "Dias_descontados_oc") {{ va = Number(va) || 0; vb = Number(vb) || 0; return asc ? va - vb : vb - va; }}
    va = String(va).toLowerCase(); vb = String(vb).toLowerCase();
    return asc ? va.localeCompare(vb, "pt-BR") : vb.localeCompare(va, "pt-BR");
  }});
  if (reapply) {{ ftPageI = 0; ftRender(); }}
}}
function ftPage(d) {{ const maxP = Math.max(0, Math.ceil(ftFiltered.length / PAGE) - 1); ftPageI = Math.min(maxP, Math.max(0, ftPageI + d)); ftRender(); }}
function ftRender() {{
  const thead = document.querySelector("#fullTable thead");
  thead.innerHTML = "<tr>" + FT_HEADERS.map((h, i) => `<th onclick="ftSort(${{i}}, ftSortCol===${{i}} ? !ftSortAsc : true)">${{h}}${{ftSortCol===i ? (ftSortAsc?" ▲":" ▼") : ""}}</th>`).join("") + "</tr>";
  const tbody = document.querySelector("#fullTable tbody");
  const start = ftPageI * PAGE;
  tbody.innerHTML = ftFiltered.slice(start, start + PAGE).map(r => "<tr>" + r.map((v, i) => {{
    let t = (v === "" || v === null) ? "—" : String(v);
    if (MONEY.has(FT_HEADERS[i]) && t !== "—") t = fmtBRL(v);
    if (FT_HEADERS[i] === "Dias_descontados_oc") t = Number(v) || 0;
    return `<td>${{t}}</td>`;
  }}).join("") + "</tr>").join("");
  const total = ftFiltered.length;
  const maxP = Math.max(1, Math.ceil(total / PAGE));
  document.getElementById("ftPg").textContent = `Página ${{ftPageI+1}} de ${{maxP}}`;
  document.getElementById("ftCount").textContent = `${{total}} de ${{FT_DATA.length}} registros`;
  document.getElementById("ftPrev").disabled = ftPageI === 0;
  document.getElementById("ftNext").disabled = ftPageI >= maxP - 1;
}}
ftApply();
</script>
"""

    empresas_lista = sorted(df["EmpresaGrupo"].unique().tolist())
    n_empresas = len(empresas_lista)

    html = f"""<!doctype html><html><head><meta charset="utf-8"><title>Relatório GVBUS — Conferência + Ocorrências</title>
<style>
body{{font-family:Segoe UI,Arial,sans-serif;margin:0;background:#f4f6fb;color:#223}}
.hero{{background:linear-gradient(120deg,{AZUL},{LAR});color:#fff;padding:34px 44px;border-radius:0 0 22px 22px}}
.hero h1{{margin:0;font-size:1.7rem}} .hero p{{opacity:.9;margin:6px 0 0}}
.wrap{{max-width:1180px;margin:26px auto;padding:0 22px}}
.kpis{{display:grid;grid-template-columns:repeat(auto-fit,minmax(190px,1fr));gap:14px;margin:-44px auto 0;max-width:1180px;padding:0 22px}}
.kpi{{background:#fff;border-radius:14px;padding:16px 18px;box-shadow:0 8px 22px -10px rgba(15,42,92,.25);border:1px solid #e6ebf3}}
.kpi .l{{font-size:.8rem;color:{CINZA};font-weight:600}} .kpi .v{{font-size:1.35rem;font-weight:800;color:{AZUL};margin-top:4px}} .kpi .s{{font-size:.75rem;color:{CINZA};margin-top:2px}}
h2{{color:{AZUL};margin:36px 0 12px;border-left:6px solid {LAR};padding-left:12px}}
h3{{color:{AZUL}}}
table{{border-collapse:collapse;width:100%;background:#fff;border-radius:10px;overflow:hidden;font-size:.85rem;box-shadow:0 2px 8px rgba(0,0,0,.06)}}
th{{background:{AZUL};color:#fff;padding:8px 10px;text-align:left;font-size:.78rem}}
td{{padding:7px 10px;border-bottom:1px solid #eef0f5}} tr:nth-child(even){{background:#f8fafc}}
.chart{{background:#fff;border-radius:14px;padding:14px;margin:14px 0;box-shadow:0 2px 8px rgba(0,0,0,.06)}}
.chart img{{max-width:100%}} .chart h3{{margin:0 0 8px;color:{AZUL};font-size:1rem}}
.inc{{background:#fff;border-radius:14px;padding:16px 20px;margin:14px 0;border-left:6px solid #ccc;box-shadow:0 2px 8px rgba(0,0,0,.05)}}
.inc.alert{{border-left-color:#b32a2a}} .inc.warn{{border-left-color:#e67e22}} .inc.ok{{border-left-color:{VERDE}}}
.inc-head{{display:flex;align-items:center;gap:12px}}
.inc-n{{min-width:44px;height:44px;border-radius:50%;background:{AZUL};color:#fff;display:flex;align-items:center;justify-content:center;font-weight:800}}
.inc.alert .inc-n{{background:#b32a2a}} .inc.warn .inc-n{{background:#e67e22}} .inc.ok .inc-n{{background:{VERDE}}}
.inc h3{{margin:0;color:{AZUL};font-size:1.02rem}}
.why{{color:#444;font-size:.88rem;background:#f8fafc;padding:10px 12px;border-radius:8px}}
.more{{color:{CINZA};font-size:.8rem}}
.note{{background:#fff7e6;border:1px solid #ffd699;padding:14px 18px;border-radius:12px;font-size:.92rem;margin:14px 0}}
ul{{line-height:1.7}} .foot{{text-align:center;color:{CINZA};font-size:.8rem;margin:40px 0 30px}}
details.more-all{{margin-top:10px;background:#f8fafc;border:1px solid #e6ebf3;border-radius:10px;padding:10px 14px}}
details.more-all summary{{cursor:pointer;font-weight:700;color:{AZUL};padding:4px 0;user-select:none}}
details.more-all summary:hover{{color:{LAR}}}
details.more-all[open] summary{{color:{LAR};border-bottom:1px solid #e6ebf3;margin-bottom:8px}}
.fulltable-toolbar{{display:flex;flex-wrap:wrap;gap:10px;align-items:center;background:#fff;padding:14px 16px;border-radius:12px 12px 0 0;border:1px solid #e6ebf3}}
.fulltable-toolbar input,.fulltable-toolbar select{{padding:7px 10px;border:1px solid #c7d2e6;border-radius:8px;font-size:.88rem}}
.fulltable-toolbar input{{flex:1;min-width:220px}}
.ft-wrap{{overflow-x:auto;background:#fff;border:1px solid #e6ebf3;border-top:0;border-radius:0 0 12px 12px;max-height:640px;overflow-y:auto}}
#fullTable{{font-size:.8rem;box-shadow:none;border-radius:0}}
#fullTable th{{position:sticky;top:0;cursor:pointer;white-space:nowrap;user-select:none}}
#fullTable th:hover{{background:#1E3F8A}}
#fullTable td{{white-space:nowrap}}
.ft-pager{{display:flex;gap:8px;align-items:center;justify-content:center;padding:10px;background:#fff;border:1px solid #e6ebf3;border-top:0;border-radius:0 0 12px 12px}}
.ft-pager button{{background:{AZUL};color:#fff;border:0;border-radius:8px;padding:6px 14px;font-weight:600;cursor:pointer}}
.ft-pager button:disabled{{background:#c7d2e6;cursor:default}}
.ft-pager .pg{{font-size:.88rem;color:{CINZA};font-weight:600}}
.ft-count{{font-size:.8rem;color:{CINZA};font-weight:600}}
</style></head><body>
<div class="hero"><h1>🟧 Relatório GVBUS — Inconsistências, Ocorrências & Redução de Valores</h1>
<p>Por empresa, contrato, escala e ocorrências · {n_empresas} empresa(s) · {len(df)} colaboradores{f" · {oc_stats['registros']} ocorrências" if oc_stats else ""} · gerado em {datetime.now().strftime('%d/%m/%Y %H:%M')}{f" · {titulo_extra}" if titulo_extra else ""}</p></div>
<div class="kpis">
<div class="kpi"><div class="l">👥 Colaboradores</div><div class="v">{TOTAL['colab']}</div></div>
<div class="kpi"><div class="l">📄 TXT bruto</div><div class="v">{_brl(TOTAL['txt'])}</div></div>
<div class="kpi"><div class="l">🟠 A DEPOSITAR</div><div class="v">{_brl(TOTAL['dep'])}</div></div>
<div class="kpi"><div class="l">💸 ECONOMIA</div><div class="v" style="color:{VERDE}">{_brl(TOTAL['eco'])}</div><div class="s">{TOTAL['eco_pct']:.1f}%</div></div>
{"<div class='kpi'><div class='l'>📋 Ocorrências</div><div class='v'>" + str(oc_stats['registros']) + "</div><div class='s'>" + str(oc_stats['dias']) + " dias descontados</div></div>" if oc_stats else ""}
{"<div class='kpi'><div class='l'>🔗 Cruzadas</div><div class='v'>" + str(n_oc_match) + "</div><div class='s'>colaboradores com ocorrência</div></div>" if oc_stats else ""}
</div>
<div class="wrap">
<h2>1 · Resumo executivo</h2>
<div class="note"><b>Como ler:</b> o TXT pede um valor por colaborador; o app abate o saldo remanescente do cartão (saldo PDF − consumo do restante do mês atual, por escala e posto — CETURB consome 1 vale/dia). Ocorrências de atestado/falta/afastamento já vieram refletidas no valor do TXT pela folha.</div>
<ul>
<li><b>TXT bruto:</b> {_brl(TOTAL['txt'])} · <b>A depositar:</b> {_brl(TOTAL['dep'])} · <b>Economia:</b> {_brl(TOTAL['eco'])} ({TOTAL['eco_pct']:.1f}%).</li>
<li><b>Status:</b> {', '.join(f"{s}: {n}" for s, n in status_counts.items())}.</li>
<li><b>Escalas:</b> {', '.join(f"{s}: {n}" for s, n in escala_counts.items())}.</li>
</ul>
<h2>2 · Por empresa</h2>
{_tbl(resumo_emp, money=("Total_TXT","Saldo_PDF","A_Depositar","Economia_R$"), pct=("Economia_%",))}
<h2>3 · Por contrato</h2>
{_tbl(contrato, money=("TXT","Saldo","Depositar","Economia"))}
{oc_sec_html}
<h2>5 · Gráficos</h2>
{charts_html}
<h2>6 · Por escala</h2>
{_tbl(df.groupby("Escala").agg(Qtde=("Matrícula","count"), TXT=("Valor TXT (R$)","sum"), Depositar=("A depositar (R$)","sum")).assign(Economia=lambda d: d.TXT - d.Depositar).reset_index(), money=("TXT","Depositar","Economia"))}
{full_table}
<h2>8 · Inconsistências — com o porquê de cada uma (tabelas expansíveis)</h2>
{inc_html}
<div class="foot">Líder Limpe · GVBUS Comparator · conferência mensal · gerado automaticamente pelo app</div>
</div></body></html>"""
    return html


# ---------------------------------------------------------------------------
# XLSX
# ---------------------------------------------------------------------------

def build_xlsx_workbook(
    result: ComparisonResult,
    ocorrencias_df: Optional[pd.DataFrame] = None,
) -> bytes:
    df = _enrich_dataframe(result, ocorrencias_df)

    wb = openpyxl.Workbook()
    wb.remove(wb.active)

    HB = PatternFill("solid", fgColor="0F2A5C")
    HO = PatternFill("solid", fgColor="FF6B1A")
    HF = Font(bold=True, color="FFFFFF", size=10)
    EDIT = PatternFill("solid", fgColor="FFF4E0")
    THIN = Border(*[Side(style="thin", color="D9E0EA")] * 4)

    # ---- Detalhado ----
    ws = wb.create_sheet("Detalhado")
    cols = ["Matrícula", "Nome", "Escala", "Empresa", "Contrato", "Posto", "Regra especial",
            "Vales/dia", "Custo/dia", "Valor TXT (R$)", "Saldo PDF (R$)", "Dias mês atual",
            "Consumo mês atual (R$)", "Saldo ajustado (R$)", "Dias mês seg.",
            "Consumo mês seg. (R$)", "Dias período (total)", "A depositar (R$)", "OBS",
            "Ocorrências do mês (motivos)", "Dias descontados (ocorr.)", "Status",
            "Corrigida?", "Matr. original (AppLider)",
            "Situação conferência", "Observação interna", "Complemento extra (R$)",
            "Motivo do complemento", "Data do registro", "Registrado por", "TOTAL FINAL (R$)"]
    ws.append(cols)

    det = df.copy()
    det["Custo/dia"] = (
        det["Custo/dia"].astype(str).str.replace("R$ ", "", regex=False)
        .str.replace(",", ".", regex=False).astype(float)
    )
    det["Corrigida?"] = det["Corrigida?"].fillna("")
    det["Matr. original (AppLider)"] = det["Matr. original (AppLider)"].fillna("")
    n_rows = len(det); last = n_rows + 1

    emp_col_name = "EmpresaGrupo" if "EmpresaGrupo" in det.columns else "Empresa"
    for i, (_, r) in enumerate(det.iterrows(), start=2):
        ws.append([
            r["Matrícula"], r["Nome"], r["Escala"], r[emp_col_name], r["Contrato"], r["Posto"],
            r["Regra especial"] if pd.notna(r["Regra especial"]) else "", r["Vales/dia"], r["Custo/dia"],
            r["Valor TXT (R$)"], r["Saldo PDF (R$)"], r["Dias mês atual"],
            r["Consumo mês atual (R$)"], r["Saldo ajustado (R$)"], r["Dias mês seg"],
            r["Consumo mês seg. (R$)"], r["Dias no período (total)"],
            r["A depositar (R$)"], r["OBS"] if pd.notna(r["OBS"]) else "",
            r["Ocorrencias_resumo"], int(r["Dias_descontados_oc"]), r["Status"],
            r["Corrigida?"], r["Matr. original (AppLider)"],
            "Pendente", "", "", "", "", "",
            f"=R{i}+N(AA{i})",
        ])
    for j, c in enumerate(cols, start=1):
        cell = ws.cell(row=1, column=j)
        cell.font = HF
        cell.fill = HO if 25 <= j <= 30 else HB
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
    for j in range(25, 31):
        for i in range(2, last + 1):
            ws.cell(row=i, column=j).fill = EDIT
    for i in range(2, last + 1):
        for j in [10, 11, 13, 14, 16, 18, 27, 31]:
            ws.cell(row=i, column=j).number_format = '#,##0.00'
        for j in range(1, 32):
            ws.cell(row=i, column=j).border = THIN
    dv = DataValidation(type="list", formula1='"Pendente,OK,Ajustado,Em análise,Rejeitado"', allow_blank=True)
    ws.add_data_validation(dv)
    dv.add(f"Y2:Y{last}")
    ws.freeze_panes = "C2"
    ws.auto_filter.ref = f"A1:AE{last}"
    widths = {"A": 10, "B": 32, "C": 14, "D": 26, "E": 15, "F": 36, "G": 10, "H": 9,
              "I": 10, "J": 12, "K": 12, "L": 11, "M": 15, "N": 14, "O": 10, "P": 15,
              "Q": 12, "R": 13, "S": 20, "T": 36, "U": 12, "V": 20, "W": 10, "X": 15,
              "Y": 17, "Z": 32, "AA": 15, "AB": 28, "AC": 13, "AD": 15, "AE": 14}
    for c, w in widths.items():
        ws.column_dimensions[c].width = w

    # ---- Ocorrências ----
    if ocorrencias_df is not None and not ocorrencias_df.empty:
        wso = wb.create_sheet("Ocorrências")
        oc_cols = ["Nome", "Empresa", "Posto de trabalho", "Escala", "Motivo",
                   "Data ocorrência", "Data início", "Data fim", "Dias afastamento",
                   "Dias descontados", "Tipo cobertura", "Responsável cobertura", "CID",
                   "Mês/Ano", "PARCEIROID"]
        wso.append(oc_cols)
        for _, r in ocorrencias_df.iterrows():
            wso.append([
                str(r["NOME"]).strip() if pd.notna(r["NOME"]) else "",
                str(r["EmpresaGrupo"]) if "EmpresaGrupo" in ocorrencias_df.columns else str(r["EMPRESA"]),
                str(r["POSTO"]).strip() if pd.notna(r["POSTO"]) else "",
                str(r["ESCALA"]).strip() if pd.notna(r["ESCALA"]) else "",
                str(r["DESCRICAO"]).strip() if pd.notna(r["DESCRICAO"]) else "",
                str(r["DATA_OC"]) if pd.notna(r["DATA_OC"]) else "",
                str(r["DATA_INI"]) if pd.notna(r["DATA_INI"]) else "",
                str(r["DATA_FIM"]) if pd.notna(r["DATA_FIM"]) else "",
                r["DIAS_AFAST"] if pd.notna(r["DIAS_AFAST"]) else "",
                r["DIAS_DESC_N"],
                str(r["TIPO_COB"]).strip() if pd.notna(r["TIPO_COB"]) else "",
                str(r["NOME_COB"]).strip() if pd.notna(r["NOME_COB"]) else "",
                str(r["CID"]).strip() if pd.notna(r["CID"]) else "",
                str(r["MES_ANO"]).strip() if pd.notna(r["MES_ANO"]) else "",
                str(r["PARCEIROID"]).strip() if pd.notna(r["PARCEIROID"]) else "",
            ])
        for j in range(1, 16):
            c = wso.cell(row=1, column=j); c.font = HF; c.fill = HB
            c.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        wso.freeze_panes = "A2"
        wso.auto_filter.ref = f"A1:O{len(ocorrencias_df) + 1}"
        for c, w in {"A": 32, "B": 24, "C": 34, "D": 26, "E": 24, "F": 14, "G": 12,
                     "H": 12, "I": 12, "J": 12, "K": 16, "L": 30, "M": 10, "N": 10, "O": 12}.items():
            wso.column_dimensions[c].width = w

    # ---- Intercorrências ----
    wsi = wb.create_sheet("Intercorrências")
    wsi.append(["Data", "Empresa", "Matrícula", "Nome", "Contrato", "Posto",
                "Complemento (R$)", "Motivo / categoria", "Detalhes do ocorrido",
                "Status", "Registrado por", "Resolvido em"])
    for j in range(1, 13):
        c = wsi.cell(row=1, column=j); c.font = HF
        c.fill = HO if j in (7, 8, 9, 10, 11, 12) else HB
        c.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
    dv2 = DataValidation(type="list", formula1='"Aberto,Em tratativa,Resolvido,Lançado no depósito"', allow_blank=True)
    wsi.add_data_validation(dv2); dv2.add("J2:J2000")
    dv3 = DataValidation(type="list",
        formula1='"Saldo divergente,Cartão novo/trocado,Falta não prevista,Férias parcial,Afastamento,Erro de escala,Matrícula divergente,Atestado pós-corte,Outros"',
        allow_blank=True)
    wsi.add_data_validation(dv3); dv3.add("H2:H2000")
    wsi.freeze_panes = "A2"
    wsi.auto_filter.ref = "A1:L2000"
    for c, w in {"A": 12, "B": 26, "C": 11, "D": 34, "E": 14, "F": 36, "G": 16,
                 "H": 22, "I": 46, "J": 16, "K": 16, "L": 14}.items():
        wsi.column_dimensions[c].width = w
    for i in range(2, 62):
        wsi.cell(row=i, column=7).number_format = '#,##0.00'
        for j in range(1, 13):
            wsi.cell(row=i, column=j).border = THIN

    # ---- Resumo ----
    wsr = wb.create_sheet("Resumo")
    def put(ws_, r_, c_, v_, bold=False, size=10, color="223344", fill=None, numfmt=None):
        cell = ws_.cell(row=r_, column=c_, value=v_)
        cell.font = Font(bold=bold, size=size, color=color)
        if fill: cell.fill = fill
        if numfmt: cell.number_format = numfmt
        return cell

    put(wsr, 1, 1, "📊 RESUMO DO MÊS — GVBUS", bold=True, size=16, color="0F2A5C")
    put(wsr, 2, 1, "Valores dinâmicos: atualizam sozinhos quando você edita Detalhado / Intercorrências.",
        size=9, color="6b7891")
    kpis = [
        ("Colaboradores", f"=COUNTA(Detalhado!A2:A{last})"),
        ("Total TXT (R$)", f"=SUM(Detalhado!J2:J{last})"),
        ("Saldo PDF (R$)", f"=SUM(Detalhado!K2:K{last})"),
        ("A depositar (R$)", f"=SUM(Detalhado!R2:R{last})"),
        ("Complementos extras do mês (R$)", f"=SUM(Detalhado!AA2:AA{last})"),
        ("TOTAL FINAL (R$)", f"=SUM(Detalhado!AE2:AE{last})"),
        ("Intercorrências registradas", '=COUNTA(Intercorrências!C2:C2000)'),
        ("Complemento via intercorrências (R$)", "=SUM(Intercorrências!G2:G2000)"),
        ("Dias descontados por ocorrências (mês)", f"=SUM(Detalhado!U2:U{last})"),
    ]
    r_ = 4
    put(wsr, r_, 1, "INDICADOR", bold=True, color="FFFFFF", fill=HB)
    put(wsr, r_, 2, "VALOR", bold=True, color="FFFFFF", fill=HB)
    r_ += 1
    for nome, f in kpis:
        put(wsr, r_, 1, nome, bold=True)
        put(wsr, r_, 2, f, numfmt='#,##0.00' if "(R$)" in nome else '0')
        r_ += 1

    # por empresa
    r_ += 1
    put(wsr, r_, 1, "POR EMPRESA", bold=True, size=12, color="0F2A5C"); r_ += 1
    for j, h in enumerate(["Empresa", "Qtde", "Total TXT", "A depositar",
                           "Compl. extras", "TOTAL FINAL", "Economia"], start=1):
        put(wsr, r_, j, h, bold=True, color="FFFFFF", fill=HB)
    r_ += 1
    emp_start = r_
    for emp in sorted(df[emp_col_name].unique().tolist()):
        put(wsr, r_, 1, emp)
        put(wsr, r_, 2, f'=COUNTIF(Detalhado!D2:D{last},A{r_})', numfmt='0')
        put(wsr, r_, 3, f'=SUMIF(Detalhado!D2:D{last},A{r_},Detalhado!J2:J{last})', numfmt='#,##0.00')
        put(wsr, r_, 4, f'=SUMIF(Detalhado!D2:D{last},A{r_},Detalhado!R2:R{last})', numfmt='#,##0.00')
        put(wsr, r_, 5, f'=SUMIF(Detalhado!D2:D{last},A{r_},Detalhado!AA2:AA{last})', numfmt='#,##0.00')
        put(wsr, r_, 6, f'=SUMIF(Detalhado!D2:D{last},A{r_},Detalhado!AE2:AE{last})', numfmt='#,##0.00')
        put(wsr, r_, 7, f'=C{r_}-D{r_}', numfmt='#,##0.00')
        r_ += 1
    put(wsr, r_, 1, "TOTAL", bold=True)
    for j, col in [(2, 'B'), (3, 'C'), (4, 'D'), (5, 'E'), (6, 'F'), (7, 'G')]:
        put(wsr, r_, j, f'=SUM({col}{emp_start}:{col}{r_-1})', bold=True,
            numfmt='#,##0.00' if j > 2 else '0')
    r_ += 2

    # por contrato
    put(wsr, r_, 1, "POR CONTRATO", bold=True, size=12, color="0F2A5C"); r_ += 1
    for j, h in enumerate(["Contrato", "Qtde", "Total TXT", "A depositar",
                           "TOTAL FINAL", "Economia"], start=1):
        put(wsr, r_, j, h, bold=True, color="FFFFFF", fill=HB)
    r_ += 1
    for ct in contrato["Contrato"].tolist():
        put(wsr, r_, 1, ct)
        put(wsr, r_, 2, f'=COUNTIF(Detalhado!E2:E{last},A{r_})', numfmt='0')
        put(wsr, r_, 3, f'=SUMIF(Detalhado!E2:E{last},A{r_},Detalhado!J2:J{last})', numfmt='#,##0.00')
        put(wsr, r_, 4, f'=SUMIF(Detalhado!E2:E{last},A{r_},Detalhado!R2:R{last})', numfmt='#,##0.00')
        put(wsr, r_, 5, f'=SUMIF(Detalhado!E2:E{last},A{r_},Detalhado!AE2:AE{last})', numfmt='#,##0.00')
        put(wsr, r_, 6, f'=C{r_}-D{r_}', numfmt='#,##0.00')
        r_ += 1

    # ocorrências por motivo
    if ocorrencias_df is not None and not ocorrencias_df.empty:
        r_ += 2
        put(wsr, r_, 1, "OCORRÊNCIAS POR MOTIVO", bold=True, size=12, color="0F2A5C"); r_ += 1
        for j, h in enumerate(["Motivo", "Registros", "Pessoas", "Dias descontados"], start=1):
            put(wsr, r_, j, h, bold=True, color="FFFFFF", fill=HB)
        r_ += 1
        oc_mot = (
            ocorrencias_df.groupby("DESCRICAO_N")
            .agg(Registros=("DESCRICAO_N", "count"), Pessoas=("NOME_N", "nunique"),
                 Dias_descontados=("DIAS_DESC_N", "sum"))
            .reset_index().sort_values("Dias_descontados", ascending=False)
        )
        oc_mot = oc_mot[oc_mot["DESCRICAO_N"].ne("NAN")]
        for _, r in oc_mot.iterrows():
            put(wsr, r_, 1, r["DESCRICAO_N"])
            put(wsr, r_, 2, int(r["Registros"]), numfmt='0')
            put(wsr, r_, 3, int(r["Pessoas"]), numfmt='0')
            put(wsr, r_, 4, int(r["Dias_descontados"]), numfmt='0')
            r_ += 1

    wsr.column_dimensions["A"].width = 34
    for c in "BCDEFG":
        wsr.column_dimensions[c].width = 16

    # ---- Inconsistências ----
    wsx = wb.create_sheet("Inconsistências")
    wsx.append(["Categoria", "Por quê importa", "Qtde", "Matrículas afetadas"])
    for j in range(1, 5):
        c = wsx.cell(row=1, column=j); c.font = HF; c.fill = HB
    inc2 = [
        ("Colaboradores SEM ESCALA (não achados no AppLider)",
         df[df["Status"].str.contains("Sem escala", na=False)],
         "Sem escala o app não calcula dias do período — sai valor cheio do TXT sem abater saldo."),
        ("Colaboradores SEM SALDO no cartão",
         df[df["Status"].str.contains("Sem saldo no cartão", na=False)],
         "Matrícula do TXT não existe no PDF do GVBUS. Conferir no portal."),
        ("Escala 2x2 (revisão manual)",
         df[df["Status"].str.contains("2x2", na=False)],
         "Por desenho o app não calcula 2x2 automaticamente."),
        ("Regra especial CETURB (1 vale/dia)",
         df[df["Regra especial"].notna() & (df["Regra especial"] != "")],
         "Consumo calculado com R$ 5,10/dia em vez de R$ 10,20."),
        ("TXT zerado MAS saldo alto no cartão (>R$100)",
         df[(df["Valor TXT (R$)"] == 0) & (df["Saldo PDF (R$)"] > 100)],
         "Férias/afastamento zeraram o TXT, mas há saldo parado considerável."),
    ]
    if ocorrencias_df is not None:
        inc2.append(("Colaboradores com ocorrências que descontaram dias",
                     df[df["Dias_descontados_oc"] > 0],
                     "Tiveram atestado/falta/afastamento com dias descontados conforme o relatório de ocorrências."))
    for titulo, dfi, expl in inc2:
        mats = ", ".join(dfi["Matrícula"].astype(str).head(40).tolist()) if len(dfi) else "—"
        if len(dfi) > 40:
            mats += f" … (+{len(dfi) - 40})"
        wsx.append([titulo, expl, len(dfi), mats])
    wsx.column_dimensions["A"].width = 52
    wsx.column_dimensions["B"].width = 60
    wsx.column_dimensions["C"].width = 8
    wsx.column_dimensions["D"].width = 80
    for row in wsx.iter_rows(min_row=2):
        for cell in row:
            cell.alignment = Alignment(wrap_text=True, vertical="top")

    # ---- Como usar ----
    wsg = wb.create_sheet("Como usar")
    guia = [
        ["GUIA DO WORKBOOK DE ACOMPANHAMENTO MENSAL", ""], ["", ""],
        ["Aba Detalhado", "Todas as linhas do cálculo. Filtros ligados. Colunas com CABEÇALHO LARANJA são editáveis:"],
        ["", "• Situação conferência — dropdown: Pendente / OK / Ajustado / Em análise / Rejeitado"],
        ["", "• Observação interna — notas livres"],
        ["", "• Complemento extra (R$) — valor depositado fora do ciclo. TOTAL FINAL = A depositar + Complemento extra (fórmula automática)."],
        ["", "• Motivo do complemento / Data / Registrado por — rastreabilidade"],
        ["", "• Colunas 'Ocorrências do mês' e 'Dias descontados' vêm do relatório de ocorrências cruzado por nome."],
        ["Aba Ocorrências", "Todas as ocorrências do mês (relatório AppLider) com motivo, datas e dias descontados. Filtrável."],
        ["Aba Intercorrências", "Diário de problemas do mês (cartão novo, saldo divergente, falta não prevista...) com categoria e status em dropdown."],
        ["Aba Resumo", "100% fórmulas — atualiza sozinho ao editar as outras abas."],
        ["Aba Inconsistências", "Categorias, porquês e matrículas afetadas do ciclo."],
        ["", ""],
        ["Dica mensal", "Salve uma cópia como 'Acompanhamento GVBUS - MES ANO.xlsx' no dia do processamento e alimente durante o mês."],
    ]
    for row in guia:
        wsg.append(row)
    wsg.cell(row=1, column=1).font = Font(bold=True, size=14, color="0F2A5C")
    wsg.column_dimensions["A"].width = 26
    wsg.column_dimensions["B"].width = 110
    for row in wsg.iter_rows(min_row=2):
        for cell in row:
            cell.alignment = Alignment(wrap_text=True, vertical="top")

    buf = io.BytesIO()
    wb.save(buf)
    return buf.getvalue()
