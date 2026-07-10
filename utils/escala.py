"""
utils/escala.py
================

Regras de escala de trabalho + calendário de feriados da Grande Vitória (ES).

Este módulo é a "verdade" de quantos dias um colaborador trabalha em qualquer
janela de datas, respeitando a escala dele. A precisão aqui é crítica: cada
dia contado a mais ou a menos vira R$ 10,20 de erro no depósito do
colaborador.

Escalas suportadas
------------------
- **5x2**   — segunda a sexta, folga sáb+dom, exclui feriados
- **6x1**   — segunda a sábado, folga dom, exclui feriados
- **12x36P** — trabalha dias PARES do mês (com virada quando o mês termina em ímpar)
- **12x36I** — trabalha dias ÍMPARES do mês (com virada quando o mês termina em ímpar)
- **2x2A / 2x2B** — não calculamos automaticamente (revisão manual)

Regra da virada 12x36
---------------------
Quando um mês termina em dia ÍMPAR (31), a paridade "vira" no dia 1º do mês
seguinte: quem era P passa a ser I definitivamente e vice-versa. Isso
acontece porque a escala é "trabalha um dia, folga um dia" e o calendário
mensal quebra essa paridade quando tem 31 dias.

Meses que geram virada: janeiro (31), março (31), maio (31), julho (31),
agosto (31), outubro (31), dezembro (31). Fevereiro, abril, junho, setembro
e novembro NÃO viram.
"""

from __future__ import annotations

from dataclasses import dataclass
from datetime import date, timedelta
from typing import Iterable, Literal

# --------------------------------------------------------------------------
# Feriados fixos + calculados (Páscoa)
# --------------------------------------------------------------------------

def _easter_sunday(year: int) -> date:
    """Algoritmo de Meeus/Jones/Butcher para Páscoa gregoriana."""
    a = year % 19
    b = year // 100
    c = year % 100
    d = b // 4
    e = b % 4
    f = (b + 8) // 25
    g = (b - f + 1) // 3
    h = (19 * a + b - d - g + 15) % 30
    i = c // 4
    k = c % 4
    l = (32 + 2 * e + 2 * i - h - k) % 7
    m = (a + 11 * h + 22 * l) // 451
    month = (h + l - 7 * m + 114) // 31
    day = ((h + l - 7 * m + 114) % 31) + 1
    return date(year, month, day)


def _last_monday_of_april(year: int) -> date:
    """Nossa Senhora da Penha — feriado estadual do ES."""
    d = date(year, 4, 30)
    while d.weekday() != 0:  # 0 = segunda
        d -= timedelta(days=1)
    return d


def feriados_grande_vitoria(year: int) -> set[date]:
    """
    Retorna o conjunto de feriados do ano na Grande Vitória, cobrindo:
    - Nacionais (incluindo Consciência Negra a partir de 2024)
    - Estaduais do ES (N.S. da Penha)
    - Móveis calculados via Páscoa: Carnaval seg+ter, Sexta Santa, Corpus Christi
    """
    easter = _easter_sunday(year)
    holidays: set[date] = {
        date(year, 1, 1),   # Ano Novo
        # Carnaval segunda + terça = 48 e 47 dias antes da Páscoa
        easter - timedelta(days=48),
        easter - timedelta(days=47),
        easter - timedelta(days=2),   # Sexta-feira Santa
        date(year, 4, 21),  # Tiradentes
        date(year, 5, 1),   # Dia do Trabalho
        easter + timedelta(days=60),  # Corpus Christi
        date(year, 9, 7),   # Independência
        date(year, 10, 12), # N.S. Aparecida
        date(year, 11, 2),  # Finados
        date(year, 11, 15), # Proclamação da República
        date(year, 12, 25), # Natal
        _last_monday_of_april(year),  # N.S. da Penha (ES)
    }
    # Consciência Negra — feriado nacional a partir de 2024
    if year >= 2024:
        holidays.add(date(year, 11, 20))
    return holidays


def feriados_no_periodo(inicio: date, fim: date) -> set[date]:
    """Une os feriados de todos os anos entre início e fim."""
    holidays: set[date] = set()
    for y in range(inicio.year, fim.year + 1):
        holidays |= feriados_grande_vitoria(y)
    return {d for d in holidays if inicio <= d <= fim}


# --------------------------------------------------------------------------
# Normalização de escala vinda do AppLider
# --------------------------------------------------------------------------

EscalaCanonica = Literal[
    "5x2", "6x1", "12x36P", "12x36I", "2x2A", "2x2B", "DESCONHECIDA"
]


def normalizar_escala(raw: str | None) -> EscalaCanonica:
    """
    Aceita as variações que vêm do AppLider:
      '5x2', '5X2\\n', '5 x 2', '5X 2'
      '6x1', '6X1\\n'
      '12x36P', '12X36P\\n', '12x36 P'
      '12x36I', '12X36I\\n', '12x36 I'
      '2x2A', '2X2A\\n', '2 x 2 A'
      '2x2B', '2X2B\\n', '2 x 2 B'
    """
    if raw is None:
        return "DESCONHECIDA"
    s = str(raw).strip().upper().replace("\n", "").replace(" ", "")
    # remove sufixos/prefixos comuns
    if s in ("5X2", "5/2", "5-2"):
        return "5x2"
    if s in ("6X1", "6/1", "6-1"):
        return "6x1"
    if s in ("12X36P", "12/36P", "12-36P", "12X36-P"):
        return "12x36P"
    if s in ("12X36I", "12/36I", "12-36I", "12X36-I"):
        return "12x36I"
    if s in ("2X2A", "2/2A", "2-2A"):
        return "2x2A"
    if s in ("2X2B", "2/2B", "2-2B"):
        return "2x2B"
    return "DESCONHECIDA"


# --------------------------------------------------------------------------
# Cálculo de dias trabalhados
# --------------------------------------------------------------------------

def _iter_dates(inicio: date, fim: date) -> Iterable[date]:
    d = inicio
    while d <= fim:
        yield d
        d += timedelta(days=1)


def _mes_termina_impar(ano: int, mes: int) -> bool:
    """Retorna True se o último dia do mês/ano dado é ímpar."""
    if mes == 12:
        prox = date(ano + 1, 1, 1)
    else:
        prox = date(ano, mes + 1, 1)
    ultimo_dia = (prox - timedelta(days=1)).day
    return ultimo_dia % 2 == 1


def paridade_efetiva_em(
    dia: date,
    escala_inicio: Literal["P", "I"],
    ano_ref: int,
    mes_ref: int,
) -> Literal["P", "I"]:
    """
    Descobre a paridade EFETIVA (P/I) de um colaborador 12x36 num dia
    específico, contando todas as viradas que aconteceram desde o mês/ano
    de referência (quando conhecemos a escala dele).

    Regra: cada mês que termina em dia ÍMPAR inverte a paridade no dia 1º do
    mês seguinte.
    """
    par = escala_inicio
    y, m = ano_ref, mes_ref
    # avança mês a mês até chegar ao ano/mês do dia
    while (y, m) < (dia.year, dia.month):
        if _mes_termina_impar(y, m):
            par = "I" if par == "P" else "P"
        if m == 12:
            y, m = y + 1, 1
        else:
            m += 1
    return par


def dias_trabalhados(
    escala: EscalaCanonica,
    inicio: date,
    fim: date,
    *,
    feriados: set[date] | None = None,
    escala_ancora_mes: date | None = None,
) -> tuple[int, list[date]]:
    """
    Retorna (qtd_dias, lista_dias) que o colaborador trabalha entre `inicio`
    e `fim` (ambos inclusive), respeitando a escala e os feriados.

    Parâmetros
    ----------
    escala : escala canônica (5x2 / 6x1 / 12x36P / 12x36I / ...)
    inicio, fim : intervalo (datas inclusivas)
    feriados : conjunto opcional de datas-feriado. Se None, usa Grande Vitória.
    escala_ancora_mes : mês/ano em que sabemos que a escala vale como
        declarada (default = o mês do `inicio`). Usado para calcular
        corretamente as viradas do 12x36. Passe uma date qualquer do
        mês-âncora (dia é ignorado).

    Retorna
    -------
    (qtd, [datas trabalhadas])
    """
    if inicio > fim:
        return 0, []

    if feriados is None:
        feriados = feriados_no_periodo(inicio, fim)

    if escala in ("2x2A", "2x2B", "DESCONHECIDA"):
        return 0, []  # não calculamos automaticamente

    if escala == "5x2":
        dias = [d for d in _iter_dates(inicio, fim)
                if d.weekday() < 5 and d not in feriados]
        return len(dias), dias

    if escala == "6x1":
        dias = [d for d in _iter_dates(inicio, fim)
                if d.weekday() < 6 and d not in feriados]
        return len(dias), dias

    if escala in ("12x36P", "12x36I"):
        assert escala in ("12x36P", "12x36I")
        paridade_inicial: Literal["P", "I"] = "P" if escala == "12x36P" else "I"
        ancora = escala_ancora_mes or inicio
        dias = []
        for d in _iter_dates(inicio, fim):
            # 12x36 trabalha feriado — os feriados só afetam 5x2 e 6x1
            par_efet = paridade_efetiva_em(
                d, paridade_inicial, ancora.year, ancora.month
            )
            if par_efet == "P" and d.day % 2 == 0:
                dias.append(d)
            elif par_efet == "I" and d.day % 2 == 1:
                dias.append(d)
        return len(dias), dias

    return 0, []


# --------------------------------------------------------------------------
# Helpers de mês
# --------------------------------------------------------------------------

def primeiro_e_ultimo_dia_do_mes(d: date) -> tuple[date, date]:
    primeiro = d.replace(day=1)
    if d.month == 12:
        ultimo = date(d.year, 12, 31)
    else:
        ultimo = date(d.year, d.month + 1, 1) - timedelta(days=1)
    return primeiro, ultimo


def mes_seguinte(d: date) -> date:
    """Retorna o primeiro dia do mês seguinte a `d`."""
    if d.month == 12:
        return date(d.year + 1, 1, 1)
    return date(d.year, d.month + 1, 1)


@dataclass(frozen=True)
class LabelEscala:
    canonica: EscalaCanonica
    label: str            # ex: "12x36 (Par)"
    cor: str              # cor hex para badges
    manual: bool = False  # True → precisa revisão manual (2x2)
    desconhecida: bool = False


LABELS_ESCALA: dict[str, LabelEscala] = {
    "5x2":     LabelEscala("5x2",     "5x2 (seg–sex)",    "#1e88e5"),
    "6x1":     LabelEscala("6x1",     "6x1 (seg–sáb)",    "#43a047"),
    "12x36P":  LabelEscala("12x36P",  "12x36 (Par)",       "#8e44ad"),
    "12x36I":  LabelEscala("12x36I",  "12x36 (Ímpar)",     "#d81b60"),
    "2x2A":    LabelEscala("2x2A",    "2x2 A (manual)",    "#e67e22", manual=True),
    "2x2B":    LabelEscala("2x2B",    "2x2 B (manual)",    "#f39c12", manual=True),
    "DESCONHECIDA": LabelEscala(
        "DESCONHECIDA", "Sem escala", "#95a5a6", desconhecida=True
    ),
}
