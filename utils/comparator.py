"""
utils/comparator.py
===================

Coração da regra de negócio. Compara o TXT comercial (valores a depositar no
próximo mês) contra:

1. **Saldo do cartão GVBUS** (PDF/XLSX) — o que já está no cartão
2. **Planilha do AppLider** — quem cada matrícula é (nome oficial) e qual a
   escala de trabalho

E, considerando a **janela de apuração** (data_inicio → data_fim), calcula
o valor final a depositar por colaborador seguindo esta lógica de precisão:

    saldo_disponivel = saldo_pdf − (dias_a_trabalhar_no_periodo × vales_por_dia × valor_vale)
    a_depositar      = max(valor_txt − max(saldo_disponivel, 0), 0)

Ou seja: o saldo mostrado no PDF ainda vai ser CONSUMIDO pelos dias que
faltam no mês atual dentro do período. Só o que sobrar de fato é abatido do
próximo depósito.

Casos especiais:
  - 2x2A / 2x2B          → não calculamos; mantém valor do TXT + sinaliza
  - Sem escala (não achou no AppLider) → seção de atenção, bloqueia geração
  - Já 0 no TXT (férias, afastamento) → mantém 0
  - Complemento sai 0 se o saldo cobrir → nunca negativo
"""

from __future__ import annotations

import re
import unicodedata
from dataclasses import dataclass, field
from datetime import date
from typing import Dict, List, Optional, Tuple

import pandas as pd

try:
    from rapidfuzz import fuzz, process as fuzz_process
    _HAS_RAPIDFUZZ = True
except ImportError:  # fallback
    _HAS_RAPIDFUZZ = False
    fuzz_process = None
    import difflib

from .escala import (
    LABELS_ESCALA,
    dias_trabalhados,
    feriados_no_periodo,
)
from .parser import TxtRow, _format_brl


# ---------------------------------------------------------------------------
# Config
# ---------------------------------------------------------------------------

VALOR_VALE_PADRAO = 5.10
VALES_POR_DIA = 2       # ida + volta


# ---------------------------------------------------------------------------
# Normalização de nomes (para fuzzy match)
# ---------------------------------------------------------------------------

_PUNCT_RE = re.compile(r"[^a-z0-9\s]")
_SPACE_RE = re.compile(r"\s+")


def normalize_name(name: str) -> str:
    if not name:
        return ""
    s = unicodedata.normalize("NFD", str(name))
    s = "".join(c for c in s if unicodedata.category(c) != "Mn")
    s = s.lower()
    s = _PUNCT_RE.sub(" ", s)
    s = _SPACE_RE.sub(" ", s).strip()
    return s


def name_similarity(a: str, b: str) -> float:
    a, b = normalize_name(a), normalize_name(b)
    if not a or not b:
        return 0.0
    if _HAS_RAPIDFUZZ:
        return fuzz.token_set_ratio(a, b)
    return difflib.SequenceMatcher(None, a, b).ratio() * 100


# ---------------------------------------------------------------------------
# Conflitos de matrícula (agora usando AppLider como fonte de verdade)
# ---------------------------------------------------------------------------

@dataclass
class MatriculaConflict:
    """
    Caso onde nome do TXT bate com um nome do AppLider, mas a matrícula
    difere. Como o AppLider é a verdade, a matrícula correta é a do AppLider.
    """
    key: str
    txt_index: int
    txt_matricula: str
    txt_nome: str
    txt_valor: float
    correct_matricula: str
    correct_nome: str
    correct_escala: str
    similarity: float


def _build_applider_index(applider_df: pd.DataFrame):
    """Constrói índices reutilizáveis do AppLider (nome normalizado -> registros).
    Retorna: (set matriculas, dict nome->linhas, lista flat pra fuzzy).
    """
    applider_mats = set(applider_df["matricula"].astype(str))
    by_nome: Dict[str, List[Tuple[str, str, str]]] = {}
    flat_names: List[str] = []
    flat_meta: List[Tuple[str, str, str]] = []  # (matricula, nome, escala)
    for _, r in applider_df.iterrows():
        nm_norm = normalize_name(r.get("nome", ""))
        if not nm_norm:
            continue
        tup = (str(r["matricula"]), str(r.get("nome", "")), str(r.get("escala", "")))
        by_nome.setdefault(nm_norm, []).append(tup)
        flat_names.append(nm_norm)
        flat_meta.append(tup)
    return applider_mats, by_nome, flat_names, flat_meta


def find_matricula_conflicts(
    txt_rows: List[TxtRow],
    applider_df: pd.DataFrame,
    *,
    min_similarity: float = 92.0,
) -> List[MatriculaConflict]:
    """
    Detecta linhas do TXT cuja matrícula NÃO está no AppLider, mas o nome
    é muito parecido com algum nome do AppLider (com matrícula diferente).

    Usa rapidfuzz.process.extractOne (C-otimizado) para escalar bem em bases
    grandes (10k+ colaboradores).
    """
    if applider_df is None or applider_df.empty:
        return []

    applider_mats, by_nome, flat_names, flat_meta = _build_applider_index(applider_df)

    conflicts: List[MatriculaConflict] = []
    seen: set[str] = set()

    for i, row in enumerate(txt_rows):
        if row.matricula in applider_mats:
            continue
        nm_txt = normalize_name(row.nome)
        if not nm_txt:
            continue

        # match exato por nome normalizado
        if nm_txt in by_nome:
            for mat, nome_orig, escala in by_nome[nm_txt]:
                if mat != row.matricula:
                    key = f"{i}::{mat}"
                    if key in seen:
                        continue
                    seen.add(key)
                    conflicts.append(MatriculaConflict(
                        key=key, txt_index=i,
                        txt_matricula=row.matricula, txt_nome=row.nome,
                        txt_valor=row.valor,
                        correct_matricula=mat, correct_nome=nome_orig,
                        correct_escala=escala,
                        similarity=100.0,
                    ))
            continue

        # fuzzy match otimizado — extractOne varre a lista inteira em C
        if _HAS_RAPIDFUZZ:
            best = fuzz_process.extractOne(
                nm_txt, flat_names,
                scorer=fuzz.token_set_ratio,
                score_cutoff=min_similarity,
            )
            if best is not None:
                _, sim, idx = best
                mat, nome_orig, escala = flat_meta[idx]
                if mat != row.matricula:
                    key = f"{i}::{mat}"
                    if key not in seen:
                        seen.add(key)
                        conflicts.append(MatriculaConflict(
                            key=key, txt_index=i,
                            txt_matricula=row.matricula, txt_nome=row.nome,
                            txt_valor=row.valor,
                            correct_matricula=mat, correct_nome=nome_orig,
                            correct_escala=escala,
                            similarity=sim,
                        ))
        else:
            # fallback puro Python — só se rapidfuzz não instalado
            best = None
            for j, nm_ap in enumerate(flat_names):
                sim = name_similarity(nm_txt, nm_ap)
                if sim >= min_similarity and (best is None or sim > best[0]):
                    mat, nome_orig, escala = flat_meta[j]
                    if mat != row.matricula:
                        best = (sim, mat, nome_orig, escala)
            if best:
                sim, mat, nome_orig, escala = best
                key = f"{i}::{mat}"
                if key not in seen:
                    seen.add(key)
                    conflicts.append(MatriculaConflict(
                        key=key, txt_index=i,
                        txt_matricula=row.matricula, txt_nome=row.nome,
                        txt_valor=row.valor,
                        correct_matricula=mat, correct_nome=nome_orig,
                        correct_escala=escala,
                        similarity=sim,
                    ))
    return conflicts


# ---------------------------------------------------------------------------
# Sem escala: bloqueadores de resultado
# ---------------------------------------------------------------------------

@dataclass
class SemEscalaCase:
    """Colaborador do TXT que não foi encontrado no AppLider (não tem escala
    associada). Precisa ser resolvido antes de gerar o resultado."""
    key: str
    txt_index: int
    matricula: str
    nome: str
    valor: float
    # sugestões (matrículas do AppLider com nome parecido, ordenadas por sim)
    sugestoes: List[Tuple[str, str, str, float]] = field(default_factory=list)
    # (matricula, nome, escala, similarity)


def find_sem_escala(
    txt_rows: List[TxtRow],
    applider_df: pd.DataFrame,
    *,
    applider_overrides: Dict[int, str] | None = None,
    min_sim: float = 70.0,
) -> List[SemEscalaCase]:
    """
    Retorna as linhas do TXT cuja matrícula-espelho do AppLider (após aplicar
    correções de conflito) NÃO está no AppLider. Cada caso vem com
    sugestões de matrículas parecidas do AppLider (top 3).
    """
    applider_overrides = applider_overrides or {}
    if applider_df is None or applider_df.empty:
        return []

    applider_mats, _, flat_names, flat_meta = _build_applider_index(applider_df)

    casos: List[SemEscalaCase] = []
    for i, row in enumerate(txt_rows):
        mat_ap = applider_overrides.get(i, row.matricula)
        if mat_ap in applider_mats:
            continue
        nm_txt = normalize_name(row.nome)
        sugestoes: List[Tuple[str, str, str, float]] = []
        if nm_txt:
            if _HAS_RAPIDFUZZ:
                # extract top-3 já vem ordenado
                results = fuzz_process.extract(
                    nm_txt, flat_names,
                    scorer=fuzz.token_set_ratio,
                    limit=3, score_cutoff=min_sim,
                )
                for _, sim, idx in results:
                    mat, nome, escala = flat_meta[idx]
                    sugestoes.append((mat, nome, escala, sim))
            else:
                scored = []
                for j, nm_ap in enumerate(flat_names):
                    sim = name_similarity(nm_txt, nm_ap)
                    if sim >= min_sim:
                        mat, nome, escala = flat_meta[j]
                        scored.append((sim, mat, nome, escala))
                scored.sort(reverse=True)
                sugestoes = [(m, n, e, s) for s, m, n, e in scored[:3]]

        casos.append(SemEscalaCase(
            key=f"noesc::{i}",
            txt_index=i,
            matricula=row.matricula,
            nome=row.nome,
            valor=row.valor,
            sugestoes=sugestoes,
        ))
    return casos


# ---------------------------------------------------------------------------
# Cálculo final
# ---------------------------------------------------------------------------

@dataclass
class ResultRow:
    matricula: str         # matrícula do TXT/cartão GVBUS (sempre preservada)
    nome: str              # nome oficial do AppLider (fallback: do TXT)
    escala: str            # canônica (5x2, 6x1, 12x36P, ...)
    escala_label: str      # amigável
    valor_txt: float
    saldo_pdf: float       # saldo bruto do PDF/planilha GVBUS
    dias_periodo: int      # dias trabalhados no período de apuração
    consumo_periodo: float # dias × 2 × valor_vale
    saldo_ajustado: float  # saldo_pdf − consumo_periodo (não negativo)
    valor_final: float     # a depositar
    obs: str
    status: str
    original_matricula: str = ""   # matrícula-espelho no AppLider
    foi_corrigida: bool = False    # True se foi feito vínculo TXT↔AppLider
    empresa: str = ""
    posto: str = ""
    manual: bool = False        # True → 2x2 (revisão manual)
    sem_escala: bool = False    # True → não achou no AppLider


@dataclass
class ComparisonResult:
    rows: List[ResultRow] = field(default_factory=list)
    total_txt: float = 0.0
    total_depositar: float = 0.0
    total_saldo_bruto: float = 0.0
    total_consumo_periodo: float = 0.0
    total_saldo_ajustado: float = 0.0
    qtd_zerados_completo: int = 0
    qtd_complemento: int = 0
    qtd_mantidos: int = 0
    qtd_sem_saldo: int = 0
    qtd_corrigidos: int = 0
    qtd_manuais_2x2: int = 0
    qtd_sem_escala: int = 0
    # metadados
    periodo_inicio: Optional[date] = None
    periodo_fim: Optional[date] = None
    valor_vale: float = VALOR_VALE_PADRAO
    n_feriados: int = 0


# status válidos:
STATUS_LABELS = {
    "zerado_completo": "✅ Saldo cobre 100%",
    "complemento":     "💰 Complemento",
    "mantido":         "—  Já 0 no TXT",
    "sem_saldo":       "🆕 Sem saldo no cartão",
    "manual_2x2":      "🖐 Revisão manual (2x2)",
    "sem_escala":      "⚠️ Sem escala (revisar)",
}


def compare(
    txt_rows: List[TxtRow],
    saldo_df: pd.DataFrame,               # GVBUS: matricula → saldo, nome
    applider_df: Optional[pd.DataFrame],  # AppLider: matricula → nome, escala, ...
    *,
    periodo_inicio: date,
    periodo_fim: date,
    valor_vale: float = VALOR_VALE_PADRAO,
    applider_overrides: Optional[Dict[int, str]] = None,
    feriados_customizados: Optional[set[date]] = None,
) -> ComparisonResult:
    """
    Executa a comparação principal.

    A matrícula do TXT é a mesma do PDF do GVBUS (sistema de cartão físico) —
    ela nunca muda no arquivo final. Já o AppLider tem uma numeração PRÓPRIA
    e diferente; para associar cada linha do TXT à sua ficha do AppLider
    (nome/escala), usamos o `applider_overrides` que mapeia
    `txt_index → matricula_do_applider`. Esse mapeamento é gerado pela UI a
    partir do casamento por nome ou de correções manuais do usuário.

    Parâmetros
    ----------
    txt_rows : linhas do TXT comercial (o quanto o próximo mês precisa)
    saldo_df : DataFrame do GVBUS (matricula, saldo, nome) — busca é feita
        pela matrícula original do TXT (mesmo sistema).
    applider_df : DataFrame do AppLider (matricula, nome, escala, empresa,
        ativo, posto). Se None, todo mundo cai em "sem_escala".
    periodo_inicio, periodo_fim : janela de apuração (inclusiva). Usada para
        calcular quantos dias o colaborador AINDA vai trabalhar (e consumir
        do saldo atual) antes do próximo depósito.
    valor_vale : valor unitário do vale-transporte
    applider_overrides : {txt_index: matricula_applider} — vínculo entre a
        linha do TXT e a ficha correta do AppLider (não altera a matrícula
        do TXT no arquivo final)
    feriados_customizados : override total dos feriados (se None, usa
        Grande Vitória automaticamente)
    """
    applider_overrides = applider_overrides or {}
    if feriados_customizados is None:
        feriados = feriados_no_periodo(periodo_inicio, periodo_fim)
    else:
        feriados = feriados_customizados

    # índice do saldo GVBUS: matricula (TXT/PDF) → (saldo, nome)
    sheet_map: Dict[str, Tuple[float, str]] = {
        str(r["matricula"]): (float(r["saldo"]), str(r.get("nome", "")))
        for _, r in saldo_df.iterrows()
    } if saldo_df is not None and not saldo_df.empty else {}

    # índice do AppLider: matricula_applider → dados completos
    applider_map: Dict[str, dict] = {}
    if applider_df is not None and not applider_df.empty:
        for _, r in applider_df.iterrows():
            applider_map[str(r["matricula"])] = {
                "nome": str(r.get("nome", "")),
                "escala": str(r.get("escala", "DESCONHECIDA")),
                "empresa": str(r.get("empresa", "")),
                "posto": str(r.get("posto", "")),
            }

    result = ComparisonResult(
        periodo_inicio=periodo_inicio,
        periodo_fim=periodo_fim,
        valor_vale=valor_vale,
        n_feriados=len(feriados),
    )

    for i, row in enumerate(txt_rows):
        # matrícula do TXT/PDF NUNCA muda
        mat_final = row.matricula
        # matrícula-espelho no AppLider (definida pelo mapping da UI)
        mat_ap = applider_overrides.get(i, row.matricula)
        foi_corrigida = mat_ap != row.matricula
        original_mat = mat_ap if foi_corrigida else ""

        # dados do AppLider (via matrícula-espelho)
        ap = applider_map.get(mat_ap, {})
        nome_oficial = ap.get("nome", "") or row.nome
        escala_can = ap.get("escala", "DESCONHECIDA")
        empresa = ap.get("empresa", "")
        posto = ap.get("posto", "")
        label = LABELS_ESCALA.get(escala_can, LABELS_ESCALA["DESCONHECIDA"])

        # dados do GVBUS (via matrícula do TXT — que é a mesma do cartão)
        saldo_pdf, _ = sheet_map.get(mat_final, (0.0, ""))

        valor_txt = row.valor

        sem_escala = escala_can == "DESCONHECIDA"
        manual = escala_can in ("2x2A", "2x2B")

        # --- lógica principal -----------------------------------------------
        if valor_txt <= 0:
            # já zerado no TXT (férias/afastamento) — mantém
            dias_periodo = 0
            consumo = 0.0
            saldo_ajustado = saldo_pdf
            valor_final = 0.0
            status = "mantido"
            result.qtd_mantidos += 1

        elif manual:
            # 2x2 — mantém valor cheio e sinaliza
            dias_periodo = 0
            consumo = 0.0
            saldo_ajustado = saldo_pdf
            valor_final = valor_txt   # mantém original
            status = "manual_2x2"
            result.qtd_manuais_2x2 += 1

        elif sem_escala:
            # sem escala — não gera valor final confiável, sinaliza
            dias_periodo = 0
            consumo = 0.0
            saldo_ajustado = saldo_pdf
            valor_final = valor_txt   # placeholder — usuário vai resolver
            status = "sem_escala"
            result.qtd_sem_escala += 1

        else:
            # 5x2 / 6x1 / 12x36
            qtd_dias, _ = dias_trabalhados(
                escala_can, periodo_inicio, periodo_fim,
                feriados=feriados,
                escala_ancora_mes=periodo_inicio,
            )
            dias_periodo = qtd_dias
            consumo = round(qtd_dias * VALES_POR_DIA * valor_vale, 2)
            saldo_ajustado = max(saldo_pdf - consumo, 0.0)

            if mat_final not in sheet_map:
                # sem saldo no cartão (matrícula não no PDF) — deposita cheio
                valor_final = valor_txt
                status = "sem_saldo"
                saldo_ajustado = 0.0
                result.qtd_sem_saldo += 1
            elif saldo_ajustado >= valor_txt:
                valor_final = 0.0
                status = "zerado_completo"
                result.qtd_zerados_completo += 1
            else:
                valor_final = round(valor_txt - saldo_ajustado, 2)
                if valor_final < 0:
                    valor_final = 0.0
                status = "complemento"
                result.qtd_complemento += 1

        if foi_corrigida:
            result.qtd_corrigidos += 1

        result.rows.append(ResultRow(
            matricula=mat_final,      # matrícula do TXT/cartão (nunca muda)
            nome=nome_oficial,
            escala=escala_can,
            escala_label=label.label,
            valor_txt=valor_txt,
            saldo_pdf=saldo_pdf,
            dias_periodo=dias_periodo,
            consumo_periodo=consumo,
            saldo_ajustado=saldo_ajustado,
            valor_final=valor_final,
            obs=row.obs,
            status=status,
            original_matricula=original_mat,   # matrícula do AppLider (referência)
            foi_corrigida=foi_corrigida,
            empresa=empresa,
            posto=posto,
            manual=manual,
            sem_escala=sem_escala,
        ))
        result.total_txt += valor_txt
        result.total_depositar += valor_final
        result.total_saldo_bruto += saldo_pdf
        result.total_consumo_periodo += consumo
        result.total_saldo_ajustado += saldo_ajustado

    return result


def result_to_txt(result: ComparisonResult) -> str:
    """
    Gera o TXT no mesmo formato do original.
    IMPORTANTE: colaboradores 2x2 e sem-escala saem com o valor original do TXT
    (marcados na OBS para não passar batido).
    """
    lines = []
    for r in result.rows:
        obs = r.obs
        # anota na obs as situações especiais que preservam o valor original
        if r.status == "manual_2x2" and "2X2" not in obs.upper():
            obs = f"{obs} [2x2 manual]".strip()
        elif r.status == "sem_escala" and "SEM ESCALA" not in obs.upper():
            obs = f"{obs} [SEM ESCALA]".strip()
        lines.append(f"{r.matricula};{r.nome};{_format_brl(r.valor_final)};{obs}")
    return "\n".join(lines) + "\n"


def result_to_dataframe(result: ComparisonResult) -> pd.DataFrame:
    """DataFrame rico para exibir no Streamlit (com todas as colunas de detalhe)."""
    data = []
    for r in result.rows:
        data.append({
            "Matrícula": r.matricula,
            "Nome": r.nome,
            "Escala": r.escala_label,
            "Empresa": r.empresa,
            "Posto": r.posto,
            "Valor TXT (R$)": r.valor_txt,
            "Saldo PDF (R$)": r.saldo_pdf,
            "Dias no período": r.dias_periodo,
            "Consumo período (R$)": r.consumo_periodo,
            "Saldo ajustado (R$)": r.saldo_ajustado,
            "A depositar (R$)": r.valor_final,
            "OBS": r.obs,
            "Status": STATUS_LABELS.get(r.status, r.status),
            "Corrigida?": "✓" if r.foi_corrigida else "",
            "Matr. original": r.original_matricula,
        })
    return pd.DataFrame(data)


def result_to_2x2_dataframe(result: ComparisonResult) -> pd.DataFrame:
    """Só as linhas 2x2 — para o download separado de revisão manual."""
    data = [
        {
            "Matrícula": r.matricula,
            "Nome": r.nome,
            "Escala": r.escala_label,
            "Empresa": r.empresa,
            "Posto": r.posto,
            "Valor TXT (R$)": r.valor_txt,
            "Saldo PDF (R$)": r.saldo_pdf,
            "OBS": r.obs,
        }
        for r in result.rows
        if r.manual
    ]
    return pd.DataFrame(data)
