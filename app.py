"""
Líder Limpe — GVBUS Comparator (v4 · Escalas + Período de Apuração)
====================================================================

App Streamlit que compara três fontes de dados:

1. **TXT comercial**  — valores a depositar no cartão GVBUS para o próximo mês
2. **Saldo GVBUS**    — PDF (preferido) ou planilha do sistema de cartão
3. **Planilha AppLider** — matrícula, nome oficial e escala de trabalho

E gera um TXT final com o valor exato a depositar por colaborador,
descontando **o que o colaborador ainda vai consumir do saldo atual** até o
fim do período de apuração informado.

Execute:
    streamlit run app.py
"""

from __future__ import annotations

import base64
import io
from datetime import date, datetime, timedelta
from pathlib import Path

import pandas as pd
import streamlit as st

from utils.parser import (
    AppLiderTable,
    FramesetXlsError,
    SaldoTable,
    TxtRow,
    parse_applider,
    parse_saldo,
    parse_txt,
    _format_brl,
)
from utils.comparator import (
    ComparisonResult,
    MatriculaConflict,
    SemEscalaCase,
    STATUS_LABELS,
    VALOR_VALE_PADRAO,
    compare,
    find_matricula_conflicts,
    find_sem_escala,
    result_to_2x2_dataframe,
    result_to_dataframe,
    result_to_txt,
)
from utils.escala import (
    LABELS_ESCALA,
    feriados_no_periodo,
    normalizar_escala,
)


# ---------------------------------------------------------------------------
# Configuração da página
# ---------------------------------------------------------------------------

st.set_page_config(
    page_title="Líder Limpe — GVBUS Comparator",
    page_icon="🟧",
    layout="wide",
    initial_sidebar_state="collapsed",
)


# ---------------------------------------------------------------------------
# CSS custom — visual moderno com cores da Líder Limpe
# ---------------------------------------------------------------------------

CUSTOM_CSS = """
<style>
.stApp { background: linear-gradient(180deg, #f8fafc 0%, #eef2f9 100%); }

/* hero header */
.hero {
    background: linear-gradient(120deg, #0F2A5C 0%, #1E3F8A 55%, #FF6B1A 130%);
    color: #fff;
    padding: 28px 32px;
    border-radius: 18px;
    margin-bottom: 24px;
    box-shadow: 0 12px 30px -10px rgba(15,42,92,.35);
    display: flex; align-items: center; gap: 22px;
}
.hero img { height: 64px; border-radius: 12px; background: #fff; padding: 4px; }
.hero h1 { margin: 0; font-size: 1.55rem; font-weight: 700; letter-spacing: -0.01em; }
.hero p  { margin: 4px 0 0; opacity: 0.88; font-size: 0.95rem; }

/* cards de métrica */
[data-testid="stMetric"] {
    background: #ffffff;
    border-radius: 14px;
    padding: 14px 18px;
    border: 1px solid #e6ebf3;
    box-shadow: 0 4px 14px -8px rgba(15,42,92,.15);
}
[data-testid="stMetricLabel"] { color: #4b5e80 !important; font-weight: 600; }
[data-testid="stMetricValue"] { color: #0F2A5C !important; font-weight: 700; }

/* uploaders */
[data-testid="stFileUploader"] section {
    border: 2px dashed #c7d2e6;
    background: #ffffff;
    border-radius: 14px;
    padding: 14px 12px;
    transition: all .15s ease;
}
[data-testid="stFileUploader"] section:hover {
    border-color: #FF6B1A;
    background: #fff7f1;
}

/* botões principais */
.stButton > button, .stDownloadButton > button {
    background: linear-gradient(120deg, #FF6B1A, #ff8a3d);
    color: #fff;
    border: 0;
    border-radius: 10px;
    font-weight: 600;
    padding: 0.55rem 1.2rem;
    box-shadow: 0 6px 18px -6px rgba(255,107,26,.55);
    transition: transform .08s ease, box-shadow .15s ease;
}
.stButton > button:hover, .stDownloadButton > button:hover {
    transform: translateY(-1px);
    box-shadow: 0 10px 22px -6px rgba(255,107,26,.65);
}

/* section titles */
.section-title {
    font-size: 1.05rem;
    font-weight: 700;
    color: #0F2A5C;
    margin: 22px 0 10px;
    display:flex; align-items:center; gap:8px;
}
.section-title .dot {
    width:10px; height:10px; border-radius:50%; background:#FF6B1A;
}

/* badges */
.badge { padding: 2px 10px; border-radius: 999px; font-size: .78rem; font-weight: 600; display:inline-block; }
.badge.ok    { background: #e6f7ee; color: #1a8a4a; }
.badge.warn  { background: #fff5e0; color: #b56a00; }
.badge.alert { background: #fde7e7; color: #b32a2a; }
.badge.info  { background: #e7eefb; color: #1e3f8a; }
.badge.purple { background: #f2e7fb; color: #6c3483; }
.badge.pink  { background: #fde4ee; color: #a83359; }

/* conflito card */
.conflict-card {
    background: #fff;
    border-left: 4px solid #FF6B1A;
    border-radius: 12px;
    padding: 12px 16px;
    margin-bottom: 8px;
    box-shadow: 0 4px 14px -10px rgba(15,42,92,.25);
}
.noesc-card {
    background: #fff8f0;
    border-left: 4px solid #b56a00;
    border-radius: 12px;
    padding: 12px 16px;
    margin-bottom: 8px;
}

/* footer */
.footer { color:#6b7891; font-size:.85rem; text-align:center; margin-top: 32px; padding: 12px; }

/* dataframe arredondado */
.stDataFrame, .stTable { border-radius: 12px; overflow: hidden; }

/* steps */
.steps { display:flex; gap:8px; margin: 0 0 20px; flex-wrap:wrap; }
.step  {
    background:#ffffff;
    border:1px solid #e6ebf3;
    padding:8px 14px;
    border-radius: 999px;
    font-size:.85rem;
    color:#6b7891;
    font-weight: 500;
}
.step.active { background:#0F2A5C; color:#fff; border-color:#0F2A5C; }
.step.done   { background:#e6f7ee; color:#1a8a4a; border-color:#bfe7d0; }

/* summary tiles (a depositar / economia / dias etc) */
.tile {
    background: #ffffff;
    border-radius: 14px;
    border: 1px solid #e6ebf3;
    padding: 16px 18px;
    box-shadow: 0 4px 14px -10px rgba(15,42,92,.15);
}
.tile .label { color:#4b5e80; font-size:.85rem; font-weight:600; margin-bottom:4px; }
.tile .value { color:#0F2A5C; font-size:1.4rem; font-weight:700; letter-spacing:-0.01em; }
.tile .subv  { color:#6b7891; font-size:.78rem; margin-top:2px; }

/* blocker */
.blocker {
    background: linear-gradient(120deg, #fff5e0, #ffe9e0);
    border-left: 5px solid #b56a00;
    padding: 14px 18px;
    border-radius: 12px;
    margin: 12px 0;
    color: #7a4a00;
}
.blocker b { color: #703b00; }
</style>
"""
st.markdown(CUSTOM_CSS, unsafe_allow_html=True)


# ---------------------------------------------------------------------------
# Header
# ---------------------------------------------------------------------------

def _logo_data_uri() -> str:
    for p in [Path(__file__).parent / "assets" / "logo.png",
              Path(__file__).parent / "assets" / "logo.jpg"]:
        if p.exists():
            data = p.read_bytes()
            mime = "image/png" if p.suffix == ".png" else "image/jpeg"
            return f"data:{mime};base64,{base64.b64encode(data).decode()}"
    return ""


logo_uri = _logo_data_uri()
logo_html = f'<img src="{logo_uri}" alt="Líder Limpe" />' if logo_uri else ""
st.markdown(f"""
<div class="hero">
    {logo_html}
    <div>
        <h1>GVBUS Comparator · Líder Limpe</h1>
        <p>Compara TXT comercial + saldo do cartão + escalas do AppLider e
        gera automaticamente o TXT final de depósito, respeitando o consumo
        que ainda ocorrerá no período de apuração.</p>
    </div>
</div>
""", unsafe_allow_html=True)


# ---------------------------------------------------------------------------
# Estado da sessão
# ---------------------------------------------------------------------------

DEFAULTS = {
    "txt_rows": None,
    "saldo_table": None,
    "applider_table": None,
    "conflicts": None,
    "sem_escala": None,
    "confirmed_overrides": {},   # dict[int, str]  (mat_ap por txt_index)
    "validated_pairs": {},       # decisões do usuário nos conflitos
    "sem_escala_decisions": {},  # {txt_index: "mat"|"IGNORE"|None}
    "min_similarity": 92.0,
    "periodo_inicio": None,
    "periodo_fim": None,
    "valor_vale": VALOR_VALE_PADRAO,
    "feriados": None,
    "feriados_manual": [],       # feriados adicionais
}
for k, v in DEFAULTS.items():
    st.session_state.setdefault(k, v)


def _reset():
    for k, v in DEFAULTS.items():
        st.session_state[k] = v


# ---------------------------------------------------------------------------
# Step indicator
# ---------------------------------------------------------------------------

def render_steps(current: int) -> None:
    labels = ["1 · Configurar", "2 · Uploads", "3 · Validação", "4 · Resultado"]
    html = '<div class="steps">'
    for i, label in enumerate(labels, start=1):
        cls = "step"
        if i < current:
            cls += " done"
        elif i == current:
            cls += " active"
        html += f'<div class="{cls}">{label}</div>'
    html += "</div>"
    st.markdown(html, unsafe_allow_html=True)


# ---------------------------------------------------------------------------
# Passo 1: período de apuração + configurações
# ---------------------------------------------------------------------------

step = 1
if st.session_state.txt_rows is not None:
    step = 3
render_steps(step)

st.markdown('<div class="section-title"><span class="dot"></span>1. Período de apuração</div>',
            unsafe_allow_html=True)
st.caption(
    "Informe o intervalo de dias em que os colaboradores ainda vão consumir o "
    "saldo atual do cartão. Normalmente pega um pedaço do mês atual e o "
    "início do próximo. Ex.: **26/06 → 25/07**."
)

hoje = date.today()
# defaults: início = hoje, fim = mesmo dia do mês seguinte
default_ini = st.session_state.periodo_inicio or hoje
if st.session_state.periodo_fim:
    default_fim = st.session_state.periodo_fim
else:
    prox = hoje.replace(day=1)
    if hoje.month == 12:
        default_fim = date(hoje.year + 1, 1, hoje.day)
    else:
        try:
            default_fim = date(hoje.year, hoje.month + 1, hoje.day)
        except ValueError:
            default_fim = date(hoje.year, hoje.month + 1, 28)

cfg_col1, cfg_col2, cfg_col3, cfg_col4 = st.columns([1.2, 1.2, 1, 1.4])
with cfg_col1:
    periodo_inicio = st.date_input(
        "🟢 Início do período",
        value=default_ini,
        format="DD/MM/YYYY",
        key="date_inicio",
    )
with cfg_col2:
    periodo_fim = st.date_input(
        "🔴 Fim do período",
        value=default_fim,
        format="DD/MM/YYYY",
        key="date_fim",
    )
with cfg_col3:
    valor_vale = st.number_input(
        "💵 Vale unitário (R$)",
        min_value=0.10, max_value=50.0,
        value=float(st.session_state.valor_vale),
        step=0.10, format="%.2f",
        help="Valor de um vale-transporte (ida OU volta). O consumo diário é "
             "2× esse valor.",
        key="input_vale",
    )
with cfg_col4:
    if periodo_inicio and periodo_fim and periodo_inicio <= periodo_fim:
        n_dias = (periodo_fim - periodo_inicio).days + 1
        feriados_calc = feriados_no_periodo(periodo_inicio, periodo_fim)
        st.markdown(f"""
        <div class="tile">
          <div class="label">📅 Janela</div>
          <div class="value">{n_dias} dias corridos</div>
          <div class="subv">{len(feriados_calc)} feriado(s) na Grande Vitória</div>
        </div>
        """, unsafe_allow_html=True)
    else:
        st.error("Período inválido")

st.session_state.periodo_inicio = periodo_inicio
st.session_state.periodo_fim = periodo_fim
st.session_state.valor_vale = valor_vale

with st.expander("⚙️ Feriados detectados no período (edite se quiser)"):
    if periodo_inicio and periodo_fim and periodo_inicio <= periodo_fim:
        base = sorted(feriados_no_periodo(periodo_inicio, periodo_fim))
        cols = st.columns(4)
        keeps = []
        for i, d in enumerate(base):
            with cols[i % 4]:
                lbl = d.strftime("%d/%m/%Y (%a)")
                if st.checkbox(lbl, value=True, key=f"fer_{d.isoformat()}"):
                    keeps.append(d)
        extra = st.text_input(
            "➕ Adicionar feriados extra (formato DD/MM/AAAA separado por vírgula)",
            key="fer_extra",
        )
        extras_parsed = []
        if extra.strip():
            for tok in extra.split(","):
                tok = tok.strip()
                try:
                    extras_parsed.append(datetime.strptime(tok, "%d/%m/%Y").date())
                except ValueError:
                    pass
        st.session_state.feriados = set(keeps) | set(extras_parsed)
    else:
        st.session_state.feriados = None


# ---------------------------------------------------------------------------
# Passo 2: uploads
# ---------------------------------------------------------------------------

st.markdown('<div class="section-title"><span class="dot"></span>2. Enviar arquivos</div>',
            unsafe_allow_html=True)

up_col1, up_col2, up_col3 = st.columns(3, gap="medium")
with up_col1:
    txt_file = st.file_uploader(
        "📄 TXT comercial",
        type=["txt"],
        key="txt_uploader",
        help="Arquivo `matrícula;nome;valor;obs` gerado pelo setor comercial.",
    )
with up_col2:
    xls_file = st.file_uploader(
        "📊 Saldo GVBUS (PDF/XLSX/XLS/CSV)",
        type=["pdf", "xlsx", "xls", "csv"],
        key="xls_uploader",
        help="Relatório do sistema do cartão. O PDF é o formato preferido.",
    )
with up_col3:
    applider_file = st.file_uploader(
        "🗂 Planilha AppLider (XLS/XLSX)",
        type=["xls", "xlsx"],
        key="applider_uploader",
        help="Exportação do AppLider com Matrícula + Nome + Tipo Escala.",
    )

st.markdown("")

with st.expander("⚙️ Ajuste fino"):
    st.session_state.min_similarity = st.slider(
        "Similaridade mínima (%) para casar nomes divergentes",
        min_value=70, max_value=100,
        value=int(st.session_state.min_similarity),
        step=1,
        help=(
            "Usado para sugerir correções de matrícula quando o TXT tem uma "
            "matrícula que não está no AppLider mas o nome bate. Diminua "
            "para pegar mais casos; aumente para ser mais rigoroso."
        ),
    )

go_col, reset_col = st.columns([1, 1])
with go_col:
    process = st.button("🚀 Processar arquivos", use_container_width=True, type="primary")
with reset_col:
    if st.button("🔄 Limpar tudo", use_container_width=True):
        _reset()
        st.rerun()


# ---------------------------------------------------------------------------
# Processamento
# ---------------------------------------------------------------------------

if process:
    if not txt_file or not xls_file or not applider_file:
        st.warning(
            "⚠️ Envie **os três arquivos** (TXT comercial + saldo GVBUS + "
            "planilha AppLider) para continuar."
        )
        st.stop()

    if periodo_inicio > periodo_fim:
        st.error("⚠️ Período inválido: a data inicial deve ser menor ou igual à final.")
        st.stop()

    # TXT
    try:
        txt_rows = parse_txt(txt_file.getvalue())
    except Exception as e:
        st.error(f"Erro ao ler o TXT: {e}")
        st.stop()
    if not txt_rows:
        st.error("O TXT está vazio ou em formato inválido.")
        st.stop()

    # Saldo
    try:
        saldo_table: SaldoTable = parse_saldo(xls_file.getvalue(), xls_file.name)
    except FramesetXlsError:
        st.error("⚠️ O arquivo `.xls` enviado está **vazio por dentro**.")
        st.markdown("""
        <div class="conflict-card" style="border-left-color:#b32a2a;">
        <b>Por que isso acontece?</b><br>
        O sistema do cartão GVBUS exporta um <code>.xls</code> no formato antigo
        <i>"Excel — Página da Web"</i>. Esse formato guarda os dados em uma pasta
        auxiliar que precisa vir junto.<br><br>
        <b>Como resolver em 10 segundos:</b><br>
        1. Abra o <code>.xls</code> no Excel<br>
        2. <b>Arquivo → Salvar como</b><br>
        3. Escolha <b>Pasta de Trabalho do Excel (*.xlsx)</b><br>
        4. Envie o <code>.xlsx</code> aqui, ou melhor: envie o <b>PDF</b> direto
        do sistema GVBUS (é o formato preferido do app)
        </div>
        """, unsafe_allow_html=True)
        st.stop()
    except Exception as e:
        st.error(f"Erro ao ler o saldo GVBUS: {e}")
        st.stop()

    # AppLider
    try:
        applider_table: AppLiderTable = parse_applider(
            applider_file.getvalue(), applider_file.name
        )
    except Exception as e:
        st.error(f"Erro ao ler a planilha do AppLider: {e}")
        st.stop()

    st.session_state.txt_rows = txt_rows
    st.session_state.saldo_table = saldo_table
    st.session_state.applider_table = applider_table
    st.session_state.conflicts = find_matricula_conflicts(
        txt_rows, applider_table.df,
        min_similarity=st.session_state.min_similarity,
    )
    st.session_state.confirmed_overrides = {}
    st.session_state.validated_pairs = {}
    st.session_state.sem_escala_decisions = {}
    st.rerun()


# ---------------------------------------------------------------------------
# A partir daqui: os três arquivos foram carregados
# ---------------------------------------------------------------------------

if (st.session_state.txt_rows is None
        or st.session_state.saldo_table is None
        or st.session_state.applider_table is None):
    st.info("👆 Configure o período, envie os **três arquivos** e clique em "
            "**Processar arquivos**.")
    st.stop()


txt_rows: list[TxtRow] = st.session_state.txt_rows
saldo_table: SaldoTable = st.session_state.saldo_table
applider_table: AppLiderTable = st.session_state.applider_table
conflicts: list[MatriculaConflict] = st.session_state.conflicts or []


# ---------------------------------------------------------------------------
# Visão geral dos inputs
# ---------------------------------------------------------------------------

st.markdown('<div class="section-title"><span class="dot"></span>Visão geral dos dados carregados</div>',
            unsafe_allow_html=True)

m1, m2, m3, m4, m5 = st.columns(5)
m1.metric("📄 Linhas no TXT", f"{len(txt_rows)}")
m2.metric("📊 Cartões (GVBUS)", f"{saldo_table.n_rows}",
          help=f"{saldo_table.n_ignored} linhas auxiliares ignoradas.")
m3.metric("🗂 AppLider", f"{applider_table.n_rows}",
          help=f"{applider_table.n_ignored} linhas ignoradas (duplicatas/sem matrícula).")
total_txt = sum(r.valor for r in txt_rows)
m4.metric("💰 Total do TXT", f"R$ {_format_brl(total_txt)}")
m5.metric("📅 Dias no período",
          f"{(st.session_state.periodo_fim - st.session_state.periodo_inicio).days + 1}",
          help=f"De {st.session_state.periodo_inicio.strftime('%d/%m/%Y')} "
               f"a {st.session_state.periodo_fim.strftime('%d/%m/%Y')}")

with st.expander("🔍 Detalhes técnicos dos arquivos"):
    d1, d2, d3 = st.columns(3)
    with d1:
        st.markdown("**TXT comercial**")
        obs_count = sum(1 for r in txt_rows if r.obs)
        st.write(f"- Linhas: {len(txt_rows)}")
        st.write(f"- Com OBS (ATS/FLT/FÉRIAS): {obs_count}")
        st.write(f"- Já zerados: {sum(1 for r in txt_rows if r.valor == 0)}")
    with d2:
        st.markdown("**Saldo GVBUS**")
        cols_str = " · ".join([c for c in saldo_table.raw_columns if c])
        st.write(f"- Cabeçalho na linha: {saldo_table.header_row_index + 1}")
        st.write(f"- Colunas: {cols_str}")
        st.write(f"- Colaboradores: {saldo_table.n_rows}")
    with d3:
        st.markdown("**Planilha AppLider**")
        st.write(f"- Cabeçalho na linha: {applider_table.header_row_index + 1}")
        st.write(f"- Colaboradores: {applider_table.n_rows}")
        # distribuição escalas
        dist = applider_table.df["escala"].value_counts()
        st.write("- Escalas:")
        for e, n in dist.items():
            st.write(f"   • {e}: {n}")


# ---------------------------------------------------------------------------
# VALIDAÇÃO 1: conflitos de matrícula
# ---------------------------------------------------------------------------

render_steps(3)

if conflicts:
    st.markdown('<div class="section-title"><span class="dot"></span>⚠️ Vínculo AppLider (matrículas divergentes)</div>',
                unsafe_allow_html=True)
    st.markdown(
        f"Encontramos **{len(conflicts)} colaboradores** com nome idêntico ao "
        "do AppLider mas com matrícula diferente. Como as matrículas do TXT e "
        "do AppLider vêm de sistemas diferentes, precisamos criar o vínculo "
        "entre eles para saber a **escala** de cada um. "
        "A matrícula do TXT no arquivo final **não muda** — só usamos a "
        "matrícula do AppLider para lookup interno."
    )

    # ações rápidas
    ac1, ac2, ac3 = st.columns([1, 1, 2])
    if ac1.button("✅ Vincular todos (100% de similaridade)",
                  use_container_width=True, key="link_100"):
        st.session_state.validated_pairs = {
            c.key: (c.similarity >= 99.5) for c in conflicts
        }
        st.session_state.confirmed_overrides = {
            c.txt_index: c.correct_matricula
            for c in conflicts if c.similarity >= 99.5
        }
        st.toast(f"{len(st.session_state.confirmed_overrides)} vínculos criados.", icon="✅")
        st.rerun()
    if ac2.button("✅ Vincular todos os sugeridos",
                  use_container_width=True, key="link_all"):
        st.session_state.validated_pairs = {c.key: True for c in conflicts}
        st.session_state.confirmed_overrides = {
            c.txt_index: c.correct_matricula for c in conflicts
        }
        st.toast(f"{len(conflicts)} vínculos criados.", icon="✅")
        st.rerun()

    n_confirmed = len(st.session_state.confirmed_overrides)
    ac3.markdown(f"""
    <div class="tile" style="margin-top:0;">
        <div class="label">Status</div>
        <div class="value">{n_confirmed} / {len(conflicts)}</div>
        <div class="subv">vínculos confirmados</div>
    </div>
    """, unsafe_allow_html=True)

    with st.expander(f"Ver os {len(conflicts)} casos individualmente",
                     expanded=len(conflicts) <= 10):
        with st.form("conflicts_form", clear_on_submit=False):
            decisions: dict[str, bool] = {}
            for c in conflicts:
                cols = st.columns([0.06, 0.4, 0.13, 0.4])
                with cols[0]:
                    checked = st.checkbox(
                        " ", value=st.session_state.validated_pairs.get(c.key, False),
                        key=f"chk_{c.key}", label_visibility="collapsed",
                    )
                with cols[1]:
                    escala_lbl = LABELS_ESCALA.get(c.correct_escala,
                                                     LABELS_ESCALA["DESCONHECIDA"])
                    st.markdown(
                        f"**TXT**<br><span style='font-size:1.02em;'>{c.txt_nome}</span><br>"
                        f"matrícula <code>{c.txt_matricula}</code> · "
                        f"valor R$ {_format_brl(c.txt_valor)}",
                        unsafe_allow_html=True,
                    )
                with cols[2]:
                    sim_cls = "ok" if c.similarity >= 99 else "info" if c.similarity >= 90 else "warn"
                    st.markdown(
                        f"<div style='text-align:center;font-size:1.4em;color:#FF6B1A;'>↔</div>"
                        f"<div style='text-align:center;'>"
                        f"<span class='badge {sim_cls}'>{c.similarity:.0f}%</span></div>",
                        unsafe_allow_html=True,
                    )
                with cols[3]:
                    st.markdown(
                        f"**AppLider**<br><span style='font-size:1.02em;'>{c.correct_nome}</span><br>"
                        f"matrícula <code>{c.correct_matricula}</code> · "
                        f"<span class='badge purple'>{escala_lbl.label}</span>",
                        unsafe_allow_html=True,
                    )
                st.markdown("<hr style='margin:2px 0 6px;border:0;border-top:1px solid #eef0f5;'/>",
                            unsafe_allow_html=True)
                decisions[c.key] = checked

            if st.form_submit_button("💾 Confirmar seleção acima", use_container_width=True):
                st.session_state.validated_pairs = decisions
                overrides: dict[int, str] = {}
                for c in conflicts:
                    if decisions.get(c.key):
                        overrides[c.txt_index] = c.correct_matricula
                st.session_state.confirmed_overrides = overrides
                st.toast(f"{len(overrides)} vínculo(s) confirmado(s).", icon="✅")
                st.rerun()
else:
    st.success("🎉 Todas as matrículas do TXT já estão vinculadas ao AppLider — nenhuma correção necessária.")


# ---------------------------------------------------------------------------
# VALIDAÇÃO 2: sem escala (bloqueia geração)
# ---------------------------------------------------------------------------

sem_casos = find_sem_escala(
    txt_rows, applider_table.df,
    applider_overrides=st.session_state.confirmed_overrides,
)

if sem_casos:
    st.markdown('<div class="section-title"><span class="dot"></span>🚧 Colaboradores sem escala (precisam ser resolvidos)</div>',
                unsafe_allow_html=True)
    st.markdown(f"""
    <div class="blocker">
    <b>{len(sem_casos)} colaboradores do TXT não têm ficha correspondente no AppLider.</b><br>
    Sem escala, não conseguimos calcular quantos dias eles ainda vão trabalhar no
    período — e portanto não podemos gerar um TXT confiável para eles.<br><br>
    <b>Como resolver rápido:</b><br>
    • Se o app achou o nome no AppLider com outra matrícula, clique em
    <b>Vincular</b> na sugestão.<br>
    • Se realmente não existe no AppLider (ex.: colaborador novo/demitido),
    marque <b>Ignorar</b> — o valor do TXT sai como está e a linha ganha a marca
    <code>[SEM ESCALA]</code> na OBS.
    </div>
    """, unsafe_allow_html=True)

    ba1, ba2, ba3 = st.columns([1, 1, 2])
    if ba1.button("⚡ Vincular todas as sugestões ≥ 90%",
                  use_container_width=True, key="quicklink90"):
        added = 0
        for c in sem_casos:
            if c.sugestoes and c.sugestoes[0][3] >= 90:
                st.session_state.confirmed_overrides[c.txt_index] = c.sugestoes[0][0]
                added += 1
        st.toast(f"{added} vínculos criados automaticamente.", icon="⚡")
        st.rerun()
    if ba2.button("🚫 Marcar todos os restantes como 'Ignorar'",
                  use_container_width=True, key="ignore_all"):
        for c in sem_casos:
            st.session_state.sem_escala_decisions[c.txt_index] = "IGNORE"
        st.toast(f"{len(sem_casos)} marcados como Ignorar.", icon="🚫")
        st.rerun()

    n_resolved = sum(
        1 for c in sem_casos
        if (c.txt_index in st.session_state.confirmed_overrides
            or st.session_state.sem_escala_decisions.get(c.txt_index) == "IGNORE")
    )
    ba3.markdown(f"""
    <div class="tile" style="margin-top:0;">
        <div class="label">Status</div>
        <div class="value">{n_resolved} / {len(sem_casos)}</div>
        <div class="subv">resolvidos</div>
    </div>
    """, unsafe_allow_html=True)

    unresolved = [
        c for c in sem_casos
        if (c.txt_index not in st.session_state.confirmed_overrides
            and st.session_state.sem_escala_decisions.get(c.txt_index) != "IGNORE")
    ]

    if unresolved:
        with st.expander(f"Resolver os {len(unresolved)} casos pendentes",
                         expanded=len(unresolved) <= 8):
            for c in unresolved:
                st.markdown(f"""
                <div class="noesc-card">
                    <b>{c.nome}</b> ·
                    matrícula TXT <code>{c.matricula}</code> ·
                    valor R$ {_format_brl(c.valor)}
                </div>
                """, unsafe_allow_html=True)

                if c.sugestoes:
                    opts = ["— escolher —"]
                    values = [None]
                    for m, n, e, sim in c.sugestoes:
                        el = LABELS_ESCALA.get(e, LABELS_ESCALA["DESCONHECIDA"])
                        opts.append(f"✅ Vincular a {m} · {n} · {el.label} · {sim:.0f}%")
                        values.append(m)
                    opts.append("🚫 Ignorar (mantém valor do TXT + marca [SEM ESCALA])")
                    values.append("IGNORE")

                    key = f"resolve_{c.txt_index}"
                    idx = st.selectbox(
                        "Ação:", opts, key=key, label_visibility="collapsed",
                    )
                    chosen = values[opts.index(idx)]
                    if chosen == "IGNORE":
                        st.session_state.sem_escala_decisions[c.txt_index] = "IGNORE"
                    elif chosen is not None:
                        st.session_state.confirmed_overrides[c.txt_index] = chosen
                else:
                    st.warning("Nenhuma sugestão encontrada no AppLider.")
                    if st.button("🚫 Ignorar (deixar como está)",
                                 key=f"ignore_{c.txt_index}"):
                        st.session_state.sem_escala_decisions[c.txt_index] = "IGNORE"
                        st.rerun()
    # botão pra "recalcular"
    if st.button("🔄 Aplicar decisões e recalcular", type="primary",
                 use_container_width=True):
        st.rerun()


# ---------------------------------------------------------------------------
# CÁLCULO E RESULTADO
# ---------------------------------------------------------------------------

render_steps(4)

# recalcula sem escala após aplicar decisões
sem_casos_now = find_sem_escala(
    txt_rows, applider_table.df,
    applider_overrides=st.session_state.confirmed_overrides,
)
unresolved_now = [
    c for c in sem_casos_now
    if st.session_state.sem_escala_decisions.get(c.txt_index) != "IGNORE"
]

# Se ainda há pendências não resolvidas, alerta o usuário — mas gera assim mesmo
# (marcando "sem_escala" para essas linhas)
if unresolved_now:
    st.warning(
        f"⚠️ {len(unresolved_now)} colaborador(es) ainda estão sem vínculo com o "
        "AppLider. O resultado abaixo já está calculado, mas essas linhas vão sair "
        "com o valor do TXT original e a marca `[SEM ESCALA]` na OBS."
    )

# roda o compare
result: ComparisonResult = compare(
    txt_rows, saldo_table.df, applider_table.df,
    periodo_inicio=st.session_state.periodo_inicio,
    periodo_fim=st.session_state.periodo_fim,
    valor_vale=st.session_state.valor_vale,
    applider_overrides=st.session_state.confirmed_overrides,
    feriados_customizados=st.session_state.feriados,
)


# ---------------------------------------------------------------------------
# Dashboard rico
# ---------------------------------------------------------------------------

st.markdown('<div class="section-title"><span class="dot"></span>💰 Resultado do cálculo</div>',
            unsafe_allow_html=True)

# tiles principais
r1, r2, r3, r4 = st.columns(4)
economia = result.total_txt - result.total_depositar
economia_pct = (economia / result.total_txt * 100) if result.total_txt else 0

r1.markdown(f"""
<div class="tile">
    <div class="label">💵 Total a depositar</div>
    <div class="value">R$ {_format_brl(result.total_depositar)}</div>
    <div class="subv">de R$ {_format_brl(result.total_txt)} no TXT</div>
</div>
""", unsafe_allow_html=True)

r2.markdown(f"""
<div class="tile">
    <div class="label">💸 Economia total</div>
    <div class="value" style="color:#1a8a4a;">−R$ {_format_brl(economia)}</div>
    <div class="subv">{economia_pct:.1f}% de redução</div>
</div>
""", unsafe_allow_html=True)

r3.markdown(f"""
<div class="tile">
    <div class="label">🏦 Saldo bruto (PDF)</div>
    <div class="value">R$ {_format_brl(result.total_saldo_bruto)}</div>
    <div class="subv">antes do consumo do período</div>
</div>
""", unsafe_allow_html=True)

r4.markdown(f"""
<div class="tile">
    <div class="label">🚌 Consumo do mês atual (desconta PDF)</div>
    <div class="value">R$ {_format_brl(result.total_consumo_mes_atual)}</div>
    <div class="subv">saldo ajustado: R$ {_format_brl(result.total_saldo_ajustado)}<br>
    mês seguinte (informativo): R$ {_format_brl(result.total_consumo_mes_seg)}</div>
</div>
""", unsafe_allow_html=True)

# breakdown por status
st.markdown('<div class="section-title"><span class="dot"></span>Distribuição por status</div>',
            unsafe_allow_html=True)
c1, c2, c3, c4, c5, c6, c7 = st.columns(7)
c1.metric("✅ Zerados", result.qtd_zerados_completo,
          help="Saldo do cartão já cobre 100% do valor do TXT — não precisa depositar.")
c2.metric("💰 Complementos", result.qtd_complemento,
          help="Deposita só a diferença (valor TXT − saldo ajustado).")
c3.metric("— Já 0 no TXT", result.qtd_mantidos,
          help="Ex.: férias, afastamento. Sai como 0,00.")
c4.metric("🆕 Sem saldo", result.qtd_sem_saldo,
          help="Colaborador não achado no cartão GVBUS. Deposita valor cheio.")
c5.metric("🖐 2x2 manual", result.qtd_manuais_2x2,
          help="Escala 2x2A/2x2B: não calculamos automaticamente.")
c6.metric("⚠️ Sem escala", result.qtd_sem_escala,
          help="Ainda não vinculados ao AppLider.")
c7.metric("🔹 CETURB (1 vale/dia)", result.qtd_posto_1_vale,
          delta=f"−R$ {_format_brl(result.economia_1_vale)}",
          delta_color="normal",
          help=("Colaboradores em posto CETURB usam apenas 1 passagem por "
                "dia (R$ 5,10) em vez de 2 (R$ 10,20). O delta mostra o "
                "quanto menos se consome de saldo por causa dessa regra."))

# distribuição por escala (bar chart embutido)
st.markdown('<div class="section-title"><span class="dot"></span>Distribuição por escala</div>',
            unsafe_allow_html=True)
esc_stats = {}
for r in result.rows:
    lbl = r.escala_label
    esc_stats.setdefault(lbl, {
        "qtd": 0, "depositar": 0.0, "txt": 0.0,
        "consumo_atual": 0.0, "consumo_seg": 0.0,
    })
    esc_stats[lbl]["qtd"] += 1
    esc_stats[lbl]["depositar"] += r.valor_final
    esc_stats[lbl]["txt"] += r.valor_txt
    esc_stats[lbl]["consumo_atual"] += r.consumo_mes_atual
    esc_stats[lbl]["consumo_seg"] += r.consumo_mes_seg

esc_df = pd.DataFrame([
    {
        "Escala": lbl,
        "Colaboradores": v["qtd"],
        "Total TXT (R$)": v["txt"],
        "Consumo mês atual (R$)": v["consumo_atual"],
        "Consumo mês seg. (R$)": v["consumo_seg"],
        "A depositar (R$)": v["depositar"],
    }
    for lbl, v in esc_stats.items()
]).sort_values("Colaboradores", ascending=False)
esc_df_display = esc_df.copy()
for col in ["Total TXT (R$)", "Consumo mês atual (R$)",
            "Consumo mês seg. (R$)", "A depositar (R$)"]:
    esc_df_display[col] = esc_df_display[col].map(_format_brl)
st.dataframe(esc_df_display, use_container_width=True, hide_index=True)

# ---------------------------------------------------------------------------
# Tabela detalhada + filtros
# ---------------------------------------------------------------------------

st.markdown('<div class="section-title"><span class="dot"></span>Tabela detalhada</div>',
            unsafe_allow_html=True)

df = result_to_dataframe(result)
df_display = df.copy()
# formata todas as colunas monetárias existentes (nomes de colunas de dias
# variam com o período, então fazemos por prefixo)
for col in df_display.columns:
    if col.endswith("(R$)"):
        df_display[col] = df_display[col].map(_format_brl)

with st.expander("🔍 Filtros", expanded=False):
    fc1, fc2, fc3 = st.columns([1, 1, 2])
    status_filter = fc1.multiselect(
        "Status",
        options=df["Status"].unique().tolist(),
        default=df["Status"].unique().tolist(),
    )
    escala_filter = fc2.multiselect(
        "Escala",
        options=df["Escala"].unique().tolist(),
        default=df["Escala"].unique().tolist(),
    )
    text_filter = fc3.text_input("Buscar por nome ou matrícula", "")

filtered = df_display[df_display["Status"].isin(status_filter)]
filtered = filtered[filtered["Escala"].isin(escala_filter)]
if text_filter:
    t = text_filter.strip().lower()
    mask = (
        filtered["Nome"].str.lower().str.contains(t, na=False)
        | filtered["Matrícula"].astype(str).str.contains(t, na=False)
    )
    filtered = filtered[mask]

st.dataframe(
    filtered,
    use_container_width=True, hide_index=True,
    column_config={
        "Matrícula": st.column_config.TextColumn(width="small"),
        "Empresa": st.column_config.TextColumn(width="medium"),
        "Posto": st.column_config.TextColumn(width="medium"),
        "OBS": st.column_config.TextColumn(width="small"),
    },
    height=520,
)


# ---------------------------------------------------------------------------
# Downloads
# ---------------------------------------------------------------------------

st.markdown('<div class="section-title"><span class="dot"></span>Downloads</div>',
            unsafe_allow_html=True)

txt_out = result_to_txt(result)
ts = datetime.now().strftime("%Y%m%d_%H%M")

dl1, dl2, dl3 = st.columns(3)
with dl1:
    st.download_button(
        "⬇️ TXT de depósito",
        data=txt_out.encode("utf-8"),
        file_name=f"deposito_gvbus_{ts}.txt",
        mime="text/plain",
        use_container_width=True,
    )
with dl2:
    csv_buf = io.StringIO()
    df.to_csv(csv_buf, index=False, sep=";", decimal=",")
    st.download_button(
        "⬇️ Relatório detalhado (CSV)",
        data=csv_buf.getvalue().encode("utf-8-sig"),
        file_name=f"relatorio_gvbus_{ts}.csv",
        mime="text/csv",
        use_container_width=True,
    )
with dl3:
    df_2x2 = result_to_2x2_dataframe(result)
    if not df_2x2.empty:
        buf22 = io.StringIO()
        df_2x2.to_csv(buf22, index=False, sep=";", decimal=",")
        st.download_button(
            f"⬇️ Revisão manual 2x2 ({len(df_2x2)})",
            data=buf22.getvalue().encode("utf-8-sig"),
            file_name=f"revisao_2x2_{ts}.csv",
            mime="text/csv",
            use_container_width=True,
        )
    else:
        st.markdown("""
        <div class="tile" style="text-align:center;">
            <div class="label">Revisão manual 2x2</div>
            <div class="value" style="font-size:1rem;">Nenhum caso</div>
        </div>
        """, unsafe_allow_html=True)

with st.expander("👁️ Pré-visualização do TXT final"):
    st.code(txt_out[:5000] + ("\n... (truncado)" if len(txt_out) > 5000 else ""),
            language="text")


# ---------------------------------------------------------------------------
# Footer
# ---------------------------------------------------------------------------

st.markdown(f"""
<div class="footer">
    Líder Limpe · Limpeza e Conservação · GVBUS Comparator v4 ·
    período {st.session_state.periodo_inicio.strftime('%d/%m/%Y')} →
    {st.session_state.periodo_fim.strftime('%d/%m/%Y')} ·
    gerado em {datetime.now().strftime("%d/%m/%Y %H:%M")}
</div>
""", unsafe_allow_html=True)
