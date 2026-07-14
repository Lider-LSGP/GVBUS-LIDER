"""
utils/comparator.py
===================

Coração da regra de negócio. Compara o TXT comercial (valores a depositar no
próximo mês) contra:

1. **Saldo do cartão GVBUS** (PDF/XLSX) — o que já está no cartão
2. **Planilha do AppLider** — quem cada matrícula é (nome oficial) e qual a
   escala de trabalho

E, considerando a **janela de apuração** (data_inicio → data_fim), calcula
o valor final a depositar por colaborador com a seguinte lógica de precisão:

    # separa o período em dois subperíodos:
    mês_atual    = periodo_inicio → último dia do mês do início
    mês_seguinte = 1º dia do próximo mês → periodo_fim

    consumo_mes_atual = dias_trabalhados(mês_atual)  × 2 × valor_vale
    consumo_mes_seg   = dias_trabalhados(mês_seguinte) × 2 × valor_vale  # informativo

    saldo_ajustado    = max(saldo_pdf − consumo_mes_atual, 0)
    a_depositar       = max(valor_txt − saldo_ajustado, 0)

Ou seja: o saldo do PDF só é consumido pelo que o colaborador AINDA vai
gastar no mês atual. O consumo do mês seguinte não entra na conta — o
próprio depósito do TXT é quem vai cobri-lo.

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
from datetime import date, timedelta
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
    primeiro_e_ultimo_dia_do_mes as _primeiro_ultimo_dia_do_mes,
)
from .parser import TxtRow, _format_brl


# ---------------------------------------------------------------------------
# Config
# ---------------------------------------------------------------------------

VALOR_VALE_PADRAO = 5.10
VALES_POR_DIA_PADRAO = 2       # ida + volta
VALES_POR_DIA_ESPECIAL = 1     # postos com passagem em cima (ex.: CETURB)

# Palavras-chave no campo `Posto Trabalho:` do AppLider que sinalizam "1 vale
# por dia". Comparado em uppercase e sem acento. Alterável por parâmetro no
# `compare()` caso a Lider Limpe passe a ter outros clientes com essa regra.
POSTOS_1_VALE_PADRAO: tuple[str, ...] = ("CETURB",)


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
    # --- separação MÊS ATUAL vs MÊS SEGUINTE ---
    dias_mes_atual: int    # dias trabalhados no restante do mês atual (consome PDF)
    consumo_mes_atual: float   # dias_mes_atual × vales_por_dia × valor_vale
    dias_mes_seg: int      # dias trabalhados no mês seguinte (informativo)
    consumo_mes_seg: float # dias_mes_seg × vales_por_dia × valor_vale (informativo)
    # --- agregado (retrocompat) ---
    dias_periodo: int      # = dias_mes_atual + dias_mes_seg
    consumo_periodo: float # = consumo_mes_atual + consumo_mes_seg
    saldo_ajustado: float  # = max(saldo_pdf − consumo_MES_ATUAL, 0)  ← só desconta o atual!
    valor_final: float     # a depositar
    obs: str
    status: str
    original_matricula: str = ""   # matrícula-espelho no AppLider
    foi_corrigida: bool = False    # True se foi feito vínculo TXT↔AppLider
    empresa: str = ""
    posto: str = ""
    manual: bool = False        # True → 2x2 (revisão manual)
    sem_escala: bool = False    # True → não achou no AppLider
    vales_por_dia: int = 2      # 2 = ida+volta padrão; 1 = passagem em cima
    posto_1_vale: bool = False  # True se caiu na regra CETURB / posto-especial
    regra_1_vale: str = ""      # palavra-chave que casou (ex.: 'CETURB')


@dataclass
class ComparisonResult:
    rows: List[ResultRow] = field(default_factory=list)
    total_txt: float = 0.0
    total_depositar: float = 0.0
    total_saldo_bruto: float = 0.0
    # totais separados
    total_consumo_mes_atual: float = 0.0  # o que desconta o PDF
    total_consumo_mes_seg: float = 0.0    # informativo
    total_consumo_periodo: float = 0.0    # soma dos dois (retrocompat)
    total_saldo_ajustado: float = 0.0
    qtd_zerados_completo: int = 0
    qtd_complemento: int = 0
    qtd_mantidos: int = 0
    qtd_sem_saldo: int = 0
    qtd_corrigidos: int = 0
    qtd_manuais_2x2: int = 0
    qtd_sem_escala: int = 0
    qtd_posto_1_vale: int = 0             # quantos caem na regra CETURB
    economia_1_vale: float = 0.0          # o quanto se economizou vs 2 vales/dia
    # metadados
    periodo_inicio: Optional[date] = None
    periodo_fim: Optional[date] = None
    # subperíodos calculados
    mes_atual_inicio: Optional[date] = None
    mes_atual_fim: Optional[date] = None
    mes_seg_inicio: Optional[date] = None
    mes_seg_fim: Optional[date] = None
    valor_vale: float = VALOR_VALE_PADRAO
    n_feriados: int = 0
    postos_1_vale: tuple[str, ...] = POSTOS_1_VALE_PADRAO


# status válidos:
STATUS_LABELS = {
    "zerado_completo": "✅ Saldo cobre 100%",
    "complemento":     "💰 Complemento",
    "mantido":         "—  Já 0 no TXT",
    "sem_saldo":       "🆕 Sem saldo no cartão",
    "manual_2x2":      "🖐 Revisão manual (2x2)",
    "sem_escala":      "⚠️ Sem escala (revisar)",
}


def _norm_posto(s: str) -> str:
    """Normaliza o nome do posto para comparar sem acento/caixa."""
    if not s:
        return ""
    s = unicodedata.normalize("NFD", str(s))
    return (
        "".join(c for c in s if unicodedata.category(c) != "Mn")
        .upper()
        .strip()
    )


def _match_posto_1_vale(posto: str, keys: tuple[str, ...]) -> Optional[str]:
    """Se o posto casar com alguma palavra-chave, retorna a key que casou.
    Senão retorna None. Comparado sem acento e em UPPER."""
    posto_norm = _norm_posto(posto)
    if not posto_norm:
        return None
    for k in keys:
        if _norm_posto(k) in posto_norm:
            return k
    return None


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
    postos_1_vale: Optional[tuple[str, ...]] = None,
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
    periodo_inicio, periodo_fim : janela de apuração (inclusiva). É dividida
        internamente em **mês atual** (que consome o PDF) e **mês seguinte**
        (informativo, coberto pelo próprio TXT).
    valor_vale : valor unitário do vale-transporte
    applider_overrides : {txt_index: matricula_applider} — vínculo entre a
        linha do TXT e a ficha correta do AppLider (não altera a matrícula
        do TXT no arquivo final)
    feriados_customizados : override total dos feriados (se None, usa
        Grande Vitória automaticamente)
    postos_1_vale : tupla de palavras-chave que, se aparecerem no campo
        `Posto Trabalho:` do AppLider, indicam que o colaborador usa apenas
        1 vale/dia (“passagem em cima”). Padrão: ('CETURB',). Comparação é
        case-insensitive e sem acento.
    """
    applider_overrides = applider_overrides or {}
    postos_1_vale = postos_1_vale or POSTOS_1_VALE_PADRAO
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

    # ------------------------------------------------------------------
    # Divisão do período em subperíodos: MÊS ATUAL vs MÊS SEGUINTE
    # ------------------------------------------------------------------
    # O MÊS ATUAL é o mês do `periodo_inicio` (o colaborador ainda vai gastar
    # o saldo do PDF durante esses dias). O TXT só cobre a parte do PRÓXIMO
    # MÊS que está dentro do período de apuração.
    #
    # Ex.: período 18/05 → 30/06 →
    #   MÊS ATUAL   = 18/05 → 31/05  (desconta do saldo PDF)
    #   MÊS SEGUINTE = 01/06 → 30/06  (informativo; coberto pelo TXT)
    #
    # Se o período não cruza a virada de mês (ex.: 05/07 → 20/07), tudo cai
    # no "mês atual" e o mês seguinte fica vazio.
    _, _ultimo_dia_mes_ini = _primeiro_ultimo_dia_do_mes(periodo_inicio)
    if periodo_fim <= _ultimo_dia_mes_ini:
        mes_atual_ini = periodo_inicio
        mes_atual_fim = periodo_fim
        mes_seg_ini = None
        mes_seg_fim = None
    else:
        mes_atual_ini = periodo_inicio
        mes_atual_fim = _ultimo_dia_mes_ini
        mes_seg_ini = _ultimo_dia_mes_ini + timedelta(days=1)
        mes_seg_fim = periodo_fim

    result = ComparisonResult(
        periodo_inicio=periodo_inicio,
        periodo_fim=periodo_fim,
        mes_atual_inicio=mes_atual_ini,
        mes_atual_fim=mes_atual_fim,
        mes_seg_inicio=mes_seg_ini,
        mes_seg_fim=mes_seg_fim,
        valor_vale=valor_vale,
        n_feriados=len(feriados),
        postos_1_vale=postos_1_vale,
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

        # regra CETURB (posto especial → 1 vale/dia)
        posto_match = _match_posto_1_vale(posto, postos_1_vale)
        posto_1_vale_flag = posto_match is not None
        vales_dia = VALES_POR_DIA_ESPECIAL if posto_1_vale_flag else VALES_POR_DIA_PADRAO

        # padrões (recalculados no caminho normal)
        dias_atual = 0
        dias_seg = 0
        consumo_atual = 0.0
        consumo_seg = 0.0
        saldo_ajustado = saldo_pdf

        # --- lógica principal -----------------------------------------------
        if valor_txt <= 0:
            # já zerado no TXT (férias/afastamento) — mantém
            valor_final = 0.0
            status = "mantido"
            result.qtd_mantidos += 1

        elif manual:
            # 2x2 — mantém valor cheio e sinaliza
            valor_final = valor_txt
            status = "manual_2x2"
            result.qtd_manuais_2x2 += 1

        elif sem_escala:
            # sem escala — não gera valor final confiável, sinaliza
            valor_final = valor_txt
            status = "sem_escala"
            result.qtd_sem_escala += 1

        else:
            # 5x2 / 6x1 / 12x36 — calcula por subperíodo
            #
            # MÊS ATUAL: dias que ainda serão trabalhados neste mês.
            # Esse consumo diminui o saldo do PDF.
            dias_atual, _ = dias_trabalhados(
                escala_can, mes_atual_ini, mes_atual_fim,
                feriados=feriados,
                escala_ancora_mes=periodo_inicio,
            )
            consumo_atual = round(dias_atual * vales_dia * valor_vale, 2)

            # MÊS SEGUINTE: dias trabalhados no início do próximo mês dentro do
            # período de apuração. Esse consumo é INFORMATIVO — o próprio TXT
            # é quem cobre esses dias, então não entra no cálculo do saldo
            # ajustado.
            if mes_seg_ini is not None:
                dias_seg, _ = dias_trabalhados(
                    escala_can, mes_seg_ini, mes_seg_fim,
                    feriados=feriados,
                    escala_ancora_mes=periodo_inicio,
                )
                consumo_seg = round(dias_seg * vales_dia * valor_vale, 2)

            # Saldo disponível quando começar o próximo mês:
            # só o consumo do MÊS ATUAL desconta do PDF.
            saldo_ajustado = max(saldo_pdf - consumo_atual, 0.0)

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
            matricula=mat_final,
            nome=nome_oficial,
            escala=escala_can,
            escala_label=label.label,
            valor_txt=valor_txt,
            saldo_pdf=saldo_pdf,
            dias_mes_atual=dias_atual,
            consumo_mes_atual=consumo_atual,
            dias_mes_seg=dias_seg,
            consumo_mes_seg=consumo_seg,
            dias_periodo=dias_atual + dias_seg,
            consumo_periodo=round(consumo_atual + consumo_seg, 2),
            saldo_ajustado=saldo_ajustado,
            valor_final=valor_final,
            obs=row.obs,
            status=status,
            original_matricula=original_mat,
            foi_corrigida=foi_corrigida,
            empresa=empresa,
            posto=posto,
            manual=manual,
            sem_escala=sem_escala,
            vales_por_dia=vales_dia,
            posto_1_vale=posto_1_vale_flag,
            regra_1_vale=posto_match or "",
        ))
        result.total_txt += valor_txt
        result.total_depositar += valor_final
        result.total_saldo_bruto += saldo_pdf
        result.total_consumo_mes_atual += consumo_atual
        result.total_consumo_mes_seg += consumo_seg
        result.total_consumo_periodo += (consumo_atual + consumo_seg)
        result.total_saldo_ajustado += saldo_ajustado
        if posto_1_vale_flag:
            result.qtd_posto_1_vale += 1
            # o quanto se economizou versus 2 vales/dia
            economia = (dias_atual + dias_seg) * (VALES_POR_DIA_PADRAO - vales_dia) * valor_vale
            result.economia_1_vale += round(economia, 2)

    result.total_consumo_periodo = round(result.total_consumo_periodo, 2)
    result.economia_1_vale = round(result.economia_1_vale, 2)
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
    """DataFrame rico para exibir no Streamlit (com todas as colunas de detalhe,
    incluindo a separação mês atual / mês seguinte e a regra CETURB)."""
    # rótulos amigáveis dos subperíodos
    if result.mes_atual_inicio and result.mes_atual_fim:
        rot_atual = (f"{result.mes_atual_inicio.strftime('%d/%m')}–"
                     f"{result.mes_atual_fim.strftime('%d/%m')}")
    else:
        rot_atual = "mês atual"
    if result.mes_seg_inicio and result.mes_seg_fim:
        rot_seg = (f"{result.mes_seg_inicio.strftime('%d/%m')}–"
                   f"{result.mes_seg_fim.strftime('%d/%m')}")
    else:
        rot_seg = "mês seguinte"

    data = []
    for r in result.rows:
        vale_dia = f"R$ {r.vales_por_dia * result.valor_vale:.2f}".replace(".", ",")
        marcador = f"🔹 {r.regra_1_vale}" if r.posto_1_vale else ""
        data.append({
            "Matrícula": r.matricula,
            "Nome": r.nome,
            "Escala": r.escala_label,
            "Empresa": r.empresa,
            "Posto": r.posto,
            "Regra especial": marcador,
            "Vales/dia": r.vales_por_dia,
            "Custo/dia": vale_dia,
            "Valor TXT (R$)": r.valor_txt,
            "Saldo PDF (R$)": r.saldo_pdf,
            # MÊS ATUAL — desconta do PDF
            f"Dias mês atual ({rot_atual})": r.dias_mes_atual,
            "Consumo mês atual (R$)": r.consumo_mes_atual,
            "Saldo ajustado (R$)": r.saldo_ajustado,
            # MÊS SEGUINTE — informativo
            f"Dias mês seg. ({rot_seg})": r.dias_mes_seg,
            "Consumo mês seg. (R$)": r.consumo_mes_seg,
            # totais e resultado
            "Dias no período (total)": r.dias_periodo,
            "A depositar (R$)": r.valor_final,
            "OBS": r.obs,
            "Status": STATUS_LABELS.get(r.status, r.status),
            "Corrigida?": "✓" if r.foi_corrigida else "",
            "Matr. original (AppLider)": r.original_matricula,
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
