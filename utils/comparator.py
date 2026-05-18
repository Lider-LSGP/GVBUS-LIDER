"""
Lógica de comparação entre o TXT comercial e a planilha de saldo do cartão.

Para cada colaborador no TXT:
  - se o saldo no cartão (planilha) >= valor do TXT → resultado = 0,00
  - se o saldo for menor → resultado = valor_txt - saldo
  - nunca negativo

Detecta também matrículas divergentes (mesmo nome aparece com matrículas
diferentes no TXT e na planilha) e oferece uma estrutura para o usuário
validar e corrigir a matrícula.
"""

from __future__ import annotations

import re
import unicodedata
from dataclasses import dataclass, field
from typing import Dict, List, Optional, Tuple

import pandas as pd

try:
    from rapidfuzz import fuzz
    _HAS_RAPIDFUZZ = True
except ImportError:  # fallback
    _HAS_RAPIDFUZZ = False
    import difflib

from .parser import TxtRow, _format_brl


# ---------------------------------------------------------------------------
# normalização de nomes
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
    """Retorna similaridade entre 0 e 100."""
    a, b = normalize_name(a), normalize_name(b)
    if not a or not b:
        return 0.0
    if _HAS_RAPIDFUZZ:
        # token_set_ratio é mais tolerante a ordem das palavras
        return fuzz.token_set_ratio(a, b)
    return difflib.SequenceMatcher(None, a, b).ratio() * 100


# ---------------------------------------------------------------------------
# Detecção de divergência de matrícula
# ---------------------------------------------------------------------------

@dataclass
class MatriculaConflict:
    """Um caso onde nome do TXT casa com um nome da planilha, mas a matrícula
    é diferente — provavelmente é a mesma pessoa com matrícula errada no TXT.
    """
    key: str                         # id estável (pra usar como chave no UI)
    txt_index: int
    txt_matricula: str
    txt_nome: str
    txt_valor: float
    sheet_matricula: str
    sheet_nome: str
    sheet_saldo: float
    similarity: float                # 0-100


def find_matricula_conflicts(
    txt_rows: List[TxtRow],
    saldo_df: pd.DataFrame,
    *,
    min_similarity: float = 92.0,
) -> List[MatriculaConflict]:
    """
    Detecta pares (linha do TXT) x (linha da planilha) em que:
      - a matrícula do TXT NÃO existe na planilha,
      - mas existe na planilha uma matrícula diferente com nome similar.
    """
    if saldo_df.empty:
        return []

    sheet_matriculas: set[str] = set(saldo_df["matricula"].astype(str).tolist())
    # mapa nome_normalizado -> lista de (matricula, nome_original, saldo)
    sheet_by_nome: Dict[str, List[Tuple[str, str, float]]] = {}
    for _, row in saldo_df.iterrows():
        nm = normalize_name(row.get("nome", ""))
        if not nm:
            continue
        sheet_by_nome.setdefault(nm, []).append(
            (str(row["matricula"]), str(row.get("nome", "")), float(row["saldo"]))
        )

    # também guarda flat list para fuzzy match
    sheet_flat = [
        (
            str(r["matricula"]),
            str(r.get("nome", "")),
            float(r["saldo"]),
            normalize_name(r.get("nome", "")),
        )
        for _, r in saldo_df.iterrows()
        if normalize_name(r.get("nome", ""))
    ]

    conflicts: List[MatriculaConflict] = []
    seen_keys: set[str] = set()

    for i, row in enumerate(txt_rows):
        if row.matricula in sheet_matriculas:
            continue  # matrícula já bate, sem conflito
        nm_txt = normalize_name(row.nome)
        if not nm_txt:
            continue

        # tenta primeiro um match exato pelo nome
        if nm_txt in sheet_by_nome:
            for mat, nome_orig, saldo in sheet_by_nome[nm_txt]:
                if mat != row.matricula:
                    key = f"{i}::{mat}"
                    if key in seen_keys:
                        continue
                    seen_keys.add(key)
                    conflicts.append(
                        MatriculaConflict(
                            key=key,
                            txt_index=i,
                            txt_matricula=row.matricula,
                            txt_nome=row.nome,
                            txt_valor=row.valor,
                            sheet_matricula=mat,
                            sheet_nome=nome_orig,
                            sheet_saldo=saldo,
                            similarity=100.0,
                        )
                    )
            continue  # se já achou match exato, não precisa fuzzy

        # fuzzy match
        best: Optional[Tuple[float, Tuple[str, str, float]]] = None
        for mat, nome_orig, saldo, nm_sheet in sheet_flat:
            if mat == row.matricula:
                continue
            sim = name_similarity(nm_txt, nm_sheet)
            if sim >= min_similarity and (best is None or sim > best[0]):
                best = (sim, (mat, nome_orig, saldo))

        if best is not None:
            sim, (mat, nome_orig, saldo) = best
            key = f"{i}::{mat}"
            if key not in seen_keys:
                seen_keys.add(key)
                conflicts.append(
                    MatriculaConflict(
                        key=key,
                        txt_index=i,
                        txt_matricula=row.matricula,
                        txt_nome=row.nome,
                        txt_valor=row.valor,
                        sheet_matricula=mat,
                        sheet_nome=nome_orig,
                        sheet_saldo=saldo,
                        similarity=sim,
                    )
                )
    return conflicts


# ---------------------------------------------------------------------------
# Cálculo final
# ---------------------------------------------------------------------------

@dataclass
class ResultRow:
    matricula: str           # matrícula final (já pode estar corrigida)
    nome: str
    valor_txt: float         # valor original do TXT
    saldo_cartao: float      # saldo encontrado na planilha (0 se não achou)
    valor_final: float       # valor a depositar (= max(valor_txt - saldo, 0))
    obs: str
    status: str              # 'zerado_completo', 'complemento', 'mantido', 'sem_saldo'
    original_matricula: str = ""   # se foi corrigida, guarda a original
    foi_corrigida: bool = False

    @property
    def diferenca(self) -> float:
        return self.valor_txt - self.valor_final


@dataclass
class ComparisonResult:
    rows: List[ResultRow] = field(default_factory=list)
    total_txt: float = 0.0
    total_depositar: float = 0.0
    qtd_zerados_completo: float = 0
    qtd_complemento: int = 0
    qtd_mantidos: int = 0
    qtd_sem_saldo: int = 0
    qtd_corrigidos: int = 0


def compare(
    txt_rows: List[TxtRow],
    saldo_df: pd.DataFrame,
    *,
    matricula_overrides: Optional[Dict[int, str]] = None,
) -> ComparisonResult:
    """
    `matricula_overrides`: mapa {txt_index: nova_matricula} com correções
    confirmadas pelo usuário.
    """
    matricula_overrides = matricula_overrides or {}

    # mapa matricula -> saldo
    sheet_map = {
        str(r["matricula"]): float(r["saldo"])
        for _, r in saldo_df.iterrows()
    }

    result = ComparisonResult()
    for i, row in enumerate(txt_rows):
        mat_final = matricula_overrides.get(i, row.matricula)
        foi_corrigida = mat_final != row.matricula
        original_mat = row.matricula if foi_corrigida else ""

        saldo = sheet_map.get(mat_final, None)
        valor_txt = row.valor

        if valor_txt <= 0:
            # já está zerado no TXT (ferias, afast etc) — mantém
            valor_final = 0.0
            status = "mantido"
        elif saldo is None:
            # não achou na planilha → deposita valor cheio
            valor_final = valor_txt
            saldo = 0.0
            status = "sem_saldo"
        elif saldo >= valor_txt:
            # já tem saldo suficiente → zera
            valor_final = 0.0
            status = "zerado_completo"
        else:
            # complemento
            valor_final = round(valor_txt - saldo, 2)
            if valor_final < 0:
                valor_final = 0.0
            status = "complemento"

        result.rows.append(
            ResultRow(
                matricula=mat_final,
                nome=row.nome,
                valor_txt=valor_txt,
                saldo_cartao=saldo if saldo is not None else 0.0,
                valor_final=valor_final,
                obs=row.obs,
                status=status,
                original_matricula=original_mat,
                foi_corrigida=foi_corrigida,
            )
        )

        result.total_txt += valor_txt
        result.total_depositar += valor_final
        if foi_corrigida:
            result.qtd_corrigidos += 1
        if status == "zerado_completo":
            result.qtd_zerados_completo += 1
        elif status == "complemento":
            result.qtd_complemento += 1
        elif status == "sem_saldo":
            result.qtd_sem_saldo += 1
        elif status == "mantido":
            result.qtd_mantidos += 1

    return result


def result_to_txt(result: ComparisonResult) -> str:
    """Gera o conteúdo do TXT de resultado no mesmo formato do original."""
    lines = []
    for r in result.rows:
        lines.append(
            f"{r.matricula};{r.nome};{_format_brl(r.valor_final)};{r.obs}"
        )
    return "\n".join(lines) + "\n"


def result_to_dataframe(result: ComparisonResult) -> pd.DataFrame:
    """Para exibição no Streamlit."""
    data = []
    for r in result.rows:
        data.append(
            {
                "Matrícula": r.matricula,
                "Nome": r.nome,
                "Valor TXT (R$)": r.valor_txt,
                "Saldo cartão (R$)": r.saldo_cartao,
                "A depositar (R$)": r.valor_final,
                "OBS": r.obs,
                "Status": _STATUS_LABEL[r.status],
                "Corrigida?": "✓" if r.foi_corrigida else "",
                "Matr. original": r.original_matricula,
            }
        )
    return pd.DataFrame(data)


_STATUS_LABEL = {
    "zerado_completo": "✅ Saldo suficiente (zerado)",
    "complemento": "💰 Complemento",
    "mantido": "—  Já zerado no TXT",
    "sem_saldo": "🆕 Sem saldo / matrícula não encontrada",
}
