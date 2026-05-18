"""
Líder Limpe — GVBUS Comparator
==============================

App Streamlit que compara o arquivo TXT comercial (valores a depositar no
cartão GVBUS) com a planilha de Saldo Estimado dos Cartões e gera um TXT
final com os valores ajustados (complementos / zerados).

Execute:
    streamlit run app.py
"""

from __future__ import annotations

import base64
import io
import os
from datetime import datetime
from pathlib import Path

import pandas as pd
import streamlit as st

from utils.parser import (
    FramesetXlsError,
    SaldoTable,
    TxtRow,
    parse_saldo,
    parse_txt,
    _format_brl,
)
from utils.comparator import (
    ComparisonResult,
    MatriculaConflict,
    compare,
    find_matricula_conflicts,
    result_to_dataframe,
    result_to_txt,
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
/* base */
.stApp {
    background: linear-gradient(180deg, #f8fafc 0%, #eef2f9 100%);
}

/* hero header */
.hero {
    background: linear-gradient(120deg, #0F2A5C 0%, #1E3F8A 55%, #FF6B1A 130%);
    color: #fff;
    padding: 28px 32px;
    border-radius: 18px;
    margin-bottom: 24px;
    box-shadow: 0 12px 30px -10px rgba(15,42,92,.35);
    display: flex;
    align-items: center;
    gap: 22px;
}
.hero img { height: 64px; border-radius: 12px; background: #fff; padding: 4px; }
.hero h1 { margin: 0; font-size: 1.55rem; font-weight: 700; letter-spacing: -0.01em; }
.hero p  { margin: 4px 0 0; opacity: 0.86; font-size: 0.95rem; }

/* cards de métrica */
[data-testid="stMetric"] {
    background: #ffffff;
    border-radius: 14px;
    padding: 14px 18px;
    border: 1px solid #e6ebf3;
    box-shadow: 0 4px 14px -8px rgba(15,42,92,.15);
}
[data-testid="stMetricLabel"] { color: #4b5e80 !important; }
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

/* botões */
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
.stButton > button:active { transform: translateY(0); }

/* secondary button */
.stButton.secondary > button {
    background: #ffffff;
    color: #0F2A5C;
    border: 1px solid #c7d2e6;
    box-shadow: none;
}

/* section titles */
.section-title {
    font-size: 1.05rem;
    font-weight: 700;
    color: #0F2A5C;
    margin: 18px 0 8px;
    display:flex; align-items:center; gap:8px;
}
.section-title .dot {
    width:10px; height:10px; border-radius:50%; background:#FF6B1A;
}

/* badges status */
.badge { padding: 2px 10px; border-radius: 999px; font-size: .8rem; font-weight: 600; }
.badge.ok    { background: #e6f7ee; color: #1a8a4a; }
.badge.warn  { background: #fff5e0; color: #b56a00; }
.badge.alert { background: #fde7e7; color: #b32a2a; }
.badge.info  { background: #e7eefb; color: #1e3f8a; }

/* conflito card */
.conflict-card {
    background: #fff;
    border-left: 4px solid #FF6B1A;
    border-radius: 12px;
    padding: 14px 16px;
    margin-bottom: 10px;
    box-shadow: 0 4px 14px -10px rgba(15,42,92,.25);
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
}
.step.active { background:#0F2A5C; color:#fff; border-color:#0F2A5C; }
.step.done   { background:#e6f7ee; color:#1a8a4a; border-color:#bfe7d0; }
</style>
"""

st.markdown(CUSTOM_CSS, unsafe_allow_html=True)


# ---------------------------------------------------------------------------
# Header
# ---------------------------------------------------------------------------

def _logo_data_uri() -> str:
    """Tenta carregar o logo local; se não tiver, devolve string vazia."""
    candidates = [
        Path(__file__).parent / "assets" / "logo.png",
        Path(__file__).parent / "assets" / "logo.jpg",
    ]
    for p in candidates:
        if p.exists():
            data = p.read_bytes()
            mime = "image/png" if p.suffix == ".png" else "image/jpeg"
            return f"data:{mime};base64,{base64.b64encode(data).decode()}"
    return ""


logo_uri = _logo_data_uri()
logo_html = f'<img src="{logo_uri}" alt="Líder Limpe" />' if logo_uri else ""
st.markdown(
    f"""
    <div class="hero">
        {logo_html}
        <div>
            <h1>GVBUS Comparator · Líder Limpe</h1>
            <p>Compare o TXT comercial com o saldo estimado dos cartões e gere
            automaticamente o arquivo final de depósito.</p>
        </div>
    </div>
    """,
    unsafe_allow_html=True,
)


# ---------------------------------------------------------------------------
# Estado da sessão
# ---------------------------------------------------------------------------

DEFAULTS = {
    "txt_rows": None,
    "saldo_table": None,
    "conflicts": None,
    "confirmed_overrides": {},   # dict[int, str]
    "validated_pairs": {},       # dict[key, bool] (decisões do usuário)
    "min_similarity": 92.0,
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
    labels = ["1 · Upload", "2 · Validação", "3 · Resultado"]
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
# Step 1: upload
# ---------------------------------------------------------------------------

render_steps(1 if st.session_state.txt_rows is None else 2)

st.markdown('<div class="section-title"><span class="dot"></span>1. Enviar arquivos</div>', unsafe_allow_html=True)

col_up1, col_up2 = st.columns(2, gap="large")
with col_up1:
    txt_file = st.file_uploader(
        "📄 TXT comercial (matrícula;nome;valor;obs)",
        type=["txt"],
        key="txt_uploader",
    )
with col_up2:
    xls_file = st.file_uploader(
        "📊 Saldo estimado (.pdf, .xlsx, .xls ou .csv)",
        type=["pdf", "xlsx", "xls", "csv"],
        key="xls_uploader",
        help=(
            "O PDF do relatório direto do sistema GVBUS é o formato preferido. "
            "O app também aceita XLSX, XLS binário e CSV."
        ),
    )

with st.expander("⚙️ Configurações avançadas"):
    st.session_state.min_similarity = st.slider(
        "Similaridade mínima para detectar matrículas divergentes (%)",
        min_value=70,
        max_value=100,
        value=int(st.session_state.min_similarity),
        step=1,
        help=(
            "Quando o nome do TXT for muito parecido com o nome da planilha, "
            "mas a matrícula for diferente, o app vai sugerir uma correção. "
            "Diminua o valor para sugerir mais casos; aumente para ser mais "
            "rigoroso."
        ),
    )

go_col, reset_col = st.columns([1, 1])
with go_col:
    process = st.button("🚀 Processar arquivos", use_container_width=True, type="primary")
with reset_col:
    if st.button("🔄 Limpar e começar de novo", use_container_width=True):
        _reset()
        st.rerun()


# ---------------------------------------------------------------------------
# Processamento dos arquivos
# ---------------------------------------------------------------------------

if process:
    if not txt_file or not xls_file:
        st.warning("⚠️ Envie tanto o TXT quanto a planilha de saldo para continuar.")
        st.stop()

    try:
        txt_rows = parse_txt(txt_file.getvalue())
    except Exception as e:
        st.error(f"Erro ao ler o TXT: {e}")
        st.stop()

    try:
        saldo_table: SaldoTable = parse_saldo(xls_file.getvalue(), xls_file.name)
    except FramesetXlsError:
        st.error("⚠️ O arquivo `.xls` enviado está **vazio por dentro**.")
        with st.container():
            st.markdown(
                """
                <div class="conflict-card" style="border-left-color:#b32a2a;">
                <b>Por que isso acontece?</b><br>
                O sistema do cartão GVBUS exporta um <code>.xls</code> no
                formato antigo <i>“Excel — Página da Web”</i>. Esse formato
                guarda os dados em uma <b>pasta auxiliar</b>
                (<code>Saldo_Estimado_Cartao_arquivos/</code>) que precisa
                vir junto com o arquivo. Quando você só envia o
                <code>.xls</code> sozinho, não há dados para ler.
                <br><br>
                <b>Como resolver em 10 segundos:</b><br>
                1. Abra o <code>Saldo_Estimado_Cartao.xls</code> no Excel<br>
                2. <b>Arquivo → Salvar como</b><br>
                3. Escolha <b>Pasta de Trabalho do Excel (*.xlsx)</b><br>
                4. Envie o <code>.xlsx</code> aqui no app — funciona
                perfeitamente 👍
                </div>
                """,
                unsafe_allow_html=True,
            )
        st.stop()
    except Exception as e:
        st.error(f"Erro ao ler a planilha: {e}")
        st.stop()

    if not txt_rows:
        st.error("O TXT está vazio ou em formato inválido.")
        st.stop()
    if saldo_table.n_rows == 0:
        st.error("A planilha está vazia após interpretação.")
        st.stop()

    st.session_state.txt_rows = txt_rows
    st.session_state.saldo_table = saldo_table
    st.session_state.conflicts = find_matricula_conflicts(
        txt_rows,
        saldo_table.df,
        min_similarity=st.session_state.min_similarity,
    )
    st.session_state.confirmed_overrides = {}
    st.session_state.validated_pairs = {}
    st.rerun()


# ---------------------------------------------------------------------------
# Após carga: visão geral + validação + resultado
# ---------------------------------------------------------------------------

if st.session_state.txt_rows is None or st.session_state.saldo_table is None:
    st.info("👆 Faça o upload dos dois arquivos e clique em **Processar arquivos**.")
    st.stop()

txt_rows: list[TxtRow] = st.session_state.txt_rows
saldo_table: SaldoTable = st.session_state.saldo_table
conflicts: list[MatriculaConflict] = st.session_state.conflicts or []

# --- Métricas iniciais ----------------------------------------------------
st.markdown('<div class="section-title"><span class="dot"></span>2. Visão geral</div>', unsafe_allow_html=True)

m1, m2, m3, m4 = st.columns(4)
m1.metric("Linhas no TXT", f"{len(txt_rows)}")
m2.metric(
    "Colaboradores na planilha",
    f"{saldo_table.n_rows}",
    help=(
        f"Cabeçalho detectado na linha {saldo_table.header_row_index + 1}. "
        f"{saldo_table.n_ignored} linhas auxiliares (pré-cabeçalho, cartões "
        f"sem matrícula, rodapé de totais) foram ignoradas automaticamente."
    ),
)

total_txt = sum(r.valor for r in txt_rows)
m3.metric("Total no TXT", f"R$ {_format_brl(total_txt)}")
m4.metric(
    "Possíveis correções de matrícula",
    f"{len(conflicts)}",
    help="Casos em que o nome bate, mas a matrícula é diferente.",
)

# detalhe da detecção
with st.expander("🔍 Detalhes da leitura do relatório de saldo"):
    cols_str = " · ".join([c for c in saldo_table.raw_columns if c])
    st.markdown(
        f"- **Linha/local do cabeçalho**: {saldo_table.header_row_index + 1}\n"
        f"- **Colunas detectadas**: {cols_str}\n"
        f"- **Linhas válidas extraídas**: {saldo_table.n_rows}\n"
        f"- **Linhas auxiliares ignoradas**: {saldo_table.n_ignored} "
        "(cabeçalho, cartões bloqueados sem matrícula, rodapé de totais)"
    )


# --- Conflicts UI ---------------------------------------------------------
if conflicts:
    st.markdown('<div class="section-title"><span class="dot"></span>⚠️ Validação de matrículas divergentes</div>',
                unsafe_allow_html=True)
    st.markdown(
        "Encontramos colaboradores no TXT com **matrícula que não existe na "
        "planilha**, mas há na planilha um nome muito parecido com matrícula "
        "diferente. Marque os pares que se referem ao mesmo colaborador e "
        "clique em **Confirmar correções**."
    )

    with st.form("conflicts_form", clear_on_submit=False):
        decisions: dict[str, bool] = {}
        for c in conflicts:
            cols = st.columns([0.07, 0.4, 0.13, 0.4])
            with cols[0]:
                checked = st.checkbox(
                    " ",
                    value=st.session_state.validated_pairs.get(c.key, False),
                    key=f"chk_{c.key}",
                    label_visibility="collapsed",
                )
            with cols[1]:
                st.markdown(
                    f"**TXT**<br><span style='font-size:1.05em;'>{c.txt_nome}</span><br>"
                    f"matrícula <code>{c.txt_matricula}</code> · "
                    f"valor R$ {_format_brl(c.txt_valor)}",
                    unsafe_allow_html=True,
                )
            with cols[2]:
                st.markdown(
                    f"<div style='text-align:center;font-size:1.4em;color:#FF6B1A;'>↔</div>"
                    f"<div style='text-align:center;'><span class='badge info'>{c.similarity:.0f}%</span></div>",
                    unsafe_allow_html=True,
                )
            with cols[3]:
                st.markdown(
                    f"**Planilha**<br><span style='font-size:1.05em;'>{c.sheet_nome}</span><br>"
                    f"matrícula <code>{c.sheet_matricula}</code> · "
                    f"saldo R$ {_format_brl(c.sheet_saldo)}",
                    unsafe_allow_html=True,
                )
            st.markdown("<hr style='margin:4px 0 12px;border:0;border-top:1px solid #eef0f5;'/>",
                        unsafe_allow_html=True)
            decisions[c.key] = checked

        col_a, col_b, col_c = st.columns([1, 1, 2])
        confirm = col_a.form_submit_button("✅ Confirmar correções", use_container_width=True)
        marcar_todas = col_b.form_submit_button("Marcar todas", use_container_width=True)

        if marcar_todas:
            st.session_state.validated_pairs = {c.key: True for c in conflicts}
            st.rerun()

        if confirm:
            st.session_state.validated_pairs = decisions
            overrides: dict[int, str] = {}
            for c in conflicts:
                if decisions.get(c.key):
                    overrides[c.txt_index] = c.sheet_matricula
            st.session_state.confirmed_overrides = overrides
            st.toast(f"{len(overrides)} matrícula(s) corrigida(s).", icon="✅")
            st.rerun()

    # se ainda não confirmou nada, mostra um aviso visual mas deixa seguir
    if not st.session_state.confirmed_overrides and any(
        st.session_state.validated_pairs.values()
    ):
        st.info(
            "Você marcou correções mas ainda não clicou em **Confirmar correções**. "
            "O resultado abaixo está sendo gerado **sem as correções**."
        )
else:
    st.success("🎉 Nenhuma divergência de matrícula encontrada — todas as matrículas do TXT estão consistentes com a planilha.")


# ---------------------------------------------------------------------------
# Step 3: resultado
# ---------------------------------------------------------------------------

render_steps(3)

result: ComparisonResult = compare(
    txt_rows,
    saldo_table.df,
    matricula_overrides=st.session_state.confirmed_overrides,
)

st.markdown('<div class="section-title"><span class="dot"></span>3. Resultado do cálculo</div>',
            unsafe_allow_html=True)

c1, c2, c3, c4, c5 = st.columns(5)
c1.metric("💵 Total a depositar", f"R$ {_format_brl(result.total_depositar)}",
          delta=f"−R$ {_format_brl(result.total_txt - result.total_depositar)}",
          delta_color="inverse")
c2.metric("✅ Zerados (saldo OK)", result.qtd_zerados_completo)
c3.metric("💰 Complementos", result.qtd_complemento)
c4.metric("🆕 Sem saldo", result.qtd_sem_saldo)
c5.metric("✏️ Matrículas corrigidas", result.qtd_corrigidos)

# preview da tabela
df = result_to_dataframe(result)
df_display = df.copy()
for col in ["Valor TXT (R$)", "Saldo cartão (R$)", "A depositar (R$)"]:
    df_display[col] = df_display[col].map(_format_brl)

# filtros
with st.expander("🔍 Filtrar", expanded=False):
    fcol1, fcol2 = st.columns([1, 2])
    status_filter = fcol1.multiselect(
        "Status",
        options=df["Status"].unique().tolist(),
        default=df["Status"].unique().tolist(),
    )
    text_filter = fcol2.text_input("Buscar por nome ou matrícula", "")

filtered = df_display[df_display["Status"].isin(status_filter)]
if text_filter:
    t = text_filter.strip().lower()
    mask = (
        filtered["Nome"].str.lower().str.contains(t, na=False)
        | filtered["Matrícula"].astype(str).str.contains(t, na=False)
    )
    filtered = filtered[mask]

st.dataframe(
    filtered,
    use_container_width=True,
    hide_index=True,
    column_config={
        "Matrícula": st.column_config.TextColumn(width="small"),
        "Status": st.column_config.TextColumn(width="medium"),
        "OBS": st.column_config.TextColumn(width="small"),
        "Corrigida?": st.column_config.TextColumn(width="small"),
    },
    height=440,
)

# ---------------------------------------------------------------------------
# Downloads
# ---------------------------------------------------------------------------

txt_out = result_to_txt(result)
ts = datetime.now().strftime("%Y%m%d_%H%M")

dl1, dl2 = st.columns(2)
with dl1:
    st.download_button(
        "⬇️ Baixar TXT de depósito",
        data=txt_out.encode("utf-8"),
        file_name=f"deposito_gvbus_{ts}.txt",
        mime="text/plain",
        use_container_width=True,
    )
with dl2:
    # csv com o detalhe completo (com saldos), para conferência
    csv_buf = io.StringIO()
    df.to_csv(csv_buf, index=False, sep=";", decimal=",")
    st.download_button(
        "⬇️ Baixar relatório detalhado (CSV)",
        data=csv_buf.getvalue().encode("utf-8-sig"),
        file_name=f"relatorio_gvbus_{ts}.csv",
        mime="text/csv",
        use_container_width=True,
    )

with st.expander("👁️ Pré-visualização do TXT final"):
    st.code(txt_out[:4000] + ("\n... (truncado)" if len(txt_out) > 4000 else ""),
            language="text")


# ---------------------------------------------------------------------------
# Footer
# ---------------------------------------------------------------------------

st.markdown(
    f"""
    <div class="footer">
        Líder Limpe · Limpeza e Conservação · GVBUS Comparator ·
        gerado em {datetime.now().strftime("%d/%m/%Y %H:%M")}
    </div>
    """,
    unsafe_allow_html=True,
)
