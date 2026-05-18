"""
Parsers do arquivo TXT (folha comercial) e da planilha (saldo do cartão GVBUS).

Aceita .xls (binário), .xlsx, .xls em formato HTML (export do Excel "Salvar
como Página da Web") e .csv. Detecta automaticamente colunas de matrícula,
nome e saldo na planilha, mesmo com pré-cabeçalho e rodapé de totais.
"""

from __future__ import annotations

import io
import re
import unicodedata
from dataclasses import dataclass
from typing import List, Optional, Tuple

import pandas as pd

try:
    import pdfplumber
    _HAS_PDFPLUMBER = True
except ImportError:  # pragma: no cover
    _HAS_PDFPLUMBER = False


# ---------------------------------------------------------------------------
# TXT
# ---------------------------------------------------------------------------

@dataclass
class TxtRow:
    """Uma linha do arquivo TXT comercial."""
    matricula: str
    nome: str
    valor: float        # valor em reais (já convertido de "x,xx" para float)
    obs: str            # ATS, FLT, FERIAS, AFAST, etc. (vazio se não houver)
    raw: str            # linha original (para debug)


_VALOR_RE = re.compile(r"^-?\d+,\d{1,2}$")


def _to_float_br(valor_str: str) -> float:
    """Converte '163,20' -> 163.20."""
    valor_str = (valor_str or "").strip().replace(".", "").replace(",", ".")
    if not valor_str:
        return 0.0
    try:
        return float(valor_str)
    except ValueError:
        return 0.0


def _format_brl(v: float) -> str:
    """Formata 163.2 -> '163,20'."""
    if v is None:
        v = 0.0
    s = f"{abs(v):.2f}"
    s = s.replace(".", ",")
    return ("-" + s) if v < 0 else s


def parse_txt(content: bytes | str) -> List[TxtRow]:
    """
    Lê o conteúdo do TXT (matricula;nome;valor;obs).
    Aceita bytes (com encoding auto) ou string.
    """
    if isinstance(content, bytes):
        text = None
        for enc in ("utf-8", "latin-1", "windows-1252", "cp850"):
            try:
                text = content.decode(enc)
                break
            except UnicodeDecodeError:
                continue
        if text is None:
            text = content.decode("utf-8", errors="replace")
    else:
        text = content

    rows: List[TxtRow] = []
    for raw in text.splitlines():
        line = raw.rstrip("\r\n")
        if not line.strip():
            continue
        parts = line.split(";")
        if len(parts) < 3:
            continue
        matricula = parts[0].strip()
        nome = parts[1].strip()
        valor_str = parts[2].strip()
        obs = ";".join(p for p in parts[3:]).strip() if len(parts) > 3 else ""

        if not matricula or not nome:
            continue

        rows.append(
            TxtRow(
                matricula=_clean_matricula(matricula),
                nome=nome,
                valor=_to_float_br(valor_str),
                obs=obs,
                raw=line,
            )
        )
    return rows


def format_txt(rows: List[TxtRow]) -> str:
    out = []
    for r in rows:
        out.append(f"{r.matricula};{r.nome};{_format_brl(r.valor)};{r.obs}")
    return "\n".join(out) + "\n"


# ---------------------------------------------------------------------------
# Planilha de saldo
# ---------------------------------------------------------------------------


def _strip_accents(s: str) -> str:
    if s is None:
        return ""
    s = unicodedata.normalize("NFD", str(s))
    return (
        "".join(c for c in s if unicodedata.category(c) != "Mn")
        .lower()
        .strip()
        .replace("\xa0", " ")
        .replace("\u200b", "")
    )


def _is_html_disguised_xls(raw: bytes) -> bool:
    head = raw[:512].lower()
    return b"<html" in head or b"<!doctype html" in head or b"<table" in head


def _is_frameset_xls(raw: bytes) -> bool:
    """Detecta o .xls 'frameset' do Excel 97-2003 (HTML que aponta para uma
    pasta auxiliar sem incluir os dados)."""
    head = raw[:4096].lower()
    return (
        b"excel workbook frameset" in head
        or (b"<frameset" in raw[:8192].lower() and b"<table" not in raw[:8192].lower())
    )


def _read_html_xls(raw: bytes) -> List[pd.DataFrame]:
    """Lê um .xls que na verdade é HTML (Excel "Salvar como Página da Web")."""
    text = None
    for enc in ("windows-1252", "latin-1", "utf-8"):
        try:
            text = raw.decode(enc)
            break
        except UnicodeDecodeError:
            continue
    if text is None:
        text = raw.decode("utf-8", errors="replace")
    text = re.sub(r"<script[\s\S]*?</script>", "", text, flags=re.IGNORECASE)
    try:
        tables = pd.read_html(io.StringIO(text), decimal=",", thousands=".")
    except ValueError:
        tables = []
    return tables


class FramesetXlsError(ValueError):
    """Erro específico: o usuário enviou um .xls 'frameset' sem a pasta
    auxiliar. Mensagem amigável no Streamlit."""


# ---------------------------------------------------------------------------
# Parser do PDF (relatório do sistema GVBUS)
# ---------------------------------------------------------------------------

# Linha de dados do PDF tem o formato:
#   06850000773592 ADELSON COELHO PINHEIRO 46149 VT Funcionário Ativo 29,60
#   06853943283923 VT Funcionário Bloqueado 0,00          ← sem nome/matrícula
#
# Estratégia: regex "de trás pra frente" — isolar saldo (final), status,
# tipo, matrícula (penult. número antes de VT) e nome (o resto).

_PDF_LINE_FULL = re.compile(
    r"^\s*(?P<cartao>\d{14,20})\s+"               # cartão (14-20 dígitos)
    r"(?P<nome>.+?)\s+"                            # nome (greedy reverso)
    r"(?P<matricula>\d{1,7})\s+"                   # matrícula
    r"(?P<tipo>VT|VR|VA|VTP)\s+"                   # tipo de utilização
    r"\S+\s+"                                      # "Funcionário" (ou similar)
    r"(?P<status>Ativo|Bloqueado|Cancelado|Inativo|Suspenso)\s+"
    r"(?P<saldo>-?[\d.]+,\d{2})\s*$",  # saldo BR (com ou sem milhar)
    re.IGNORECASE,
)

# fallback: cartão sem nome/matrícula (bloqueado etc.)
_PDF_LINE_NONAME = re.compile(
    r"^\s*(?P<cartao>\d{14,20})\s+"
    r"(?P<tipo>VT|VR|VA|VTP)\s+"
    r"\S+\s+"
    r"(?P<status>Ativo|Bloqueado|Cancelado|Inativo|Suspenso)\s+"
    r"(?P<saldo>-?[\d.]+,\d{2})\s*$",
    re.IGNORECASE,
)

_PDF_SKIP_PREFIXES = (
    "consulta", "hora:", "ordenado", "página", "pagina", "titular:",
    "cnpj:", "cartão", "cartao", "total",
)


def _read_pdf_saldo(raw: bytes) -> Optional[pd.DataFrame]:
    """Lê o PDF do relatório GVBUS e devolve um DataFrame com colunas
    [Cartão, Funcionário, Matrícula, Saldo, Status, Tipo].
    """
    records: list[dict] = []
    with pdfplumber.open(io.BytesIO(raw)) as pdf:
        for page in pdf.pages:
            text = page.extract_text() or ""
            for line in text.split("\n"):
                line = line.strip()
                if not line:
                    continue
                low = line.lower()
                if any(low.startswith(p) for p in _PDF_SKIP_PREFIXES):
                    continue

                m = _PDF_LINE_FULL.match(line)
                if m:
                    records.append(
                        {
                            "Cartão": m.group("cartao"),
                            "Funcionário": m.group("nome").strip(),
                            "Matrícula": m.group("matricula"),
                            "Tipo": m.group("tipo"),
                            "Status": m.group("status"),
                            "Saldo": m.group("saldo"),
                        }
                    )
                    continue

                m = _PDF_LINE_NONAME.match(line)
                if m:
                    records.append(
                        {
                            "Cartão": m.group("cartao"),
                            "Funcionário": "",
                            "Matrícula": "",
                            "Tipo": m.group("tipo"),
                            "Status": m.group("status"),
                            "Saldo": m.group("saldo"),
                        }
                    )
                    continue
                # se não casa, ignora (pode ser quebra de página ou rodapé)

    if not records:
        return None

    df = pd.DataFrame(records)
    # Acrescenta uma linha de cabeçalho no topo (para o detector de header
    # reusar a mesma pipeline da planilha)
    header = pd.DataFrame(
        [["Cartão:", "Funcionário:", "Matrícula:", "Tipo Utilização:", "Status:", "Saldo:"]],
        columns=df.columns,
    )
    return pd.concat([header, df], ignore_index=True)


def _try_read_any(raw: bytes, filename: str) -> List[pd.DataFrame]:
    """Tenta ler o arquivo em todos os formatos possíveis (sempre header=None,
    pois o cabeçalho real será detectado depois)."""
    name = (filename or "").lower()
    errors: list[str] = []

    # PDF: relatório direto do sistema GVBUS
    if name.endswith(".pdf") or raw[:4] == b"%PDF":
        if not _HAS_PDFPLUMBER:
            raise ValueError(
                "Para ler PDFs, instale a dependência 'pdfplumber' "
                "(pip install pdfplumber)."
            )
        df = _read_pdf_saldo(raw)
        if df is None or df.empty:
            raise ValueError(
                "PDF lido, mas não foi possível extrair nenhum colaborador. "
                "O formato do PDF mudou? Esperado: 'cartão nome matrícula "
                "VT Funcionário Status saldo'."
            )
        return [df]

    # detecta o frameset xls ANTES de tentar outros formatos (mensagem clara)
    if _is_frameset_xls(raw):
        raise FramesetXlsError(
            "O arquivo enviado é um '.xls' do tipo 'Página da Web' (frameset) "
            "do Excel 97-2003 e está vazio — os dados ficam em uma pasta "
            "auxiliar que não veio junto. Abra o arquivo no Excel e use "
            "'Salvar como → Pasta de Trabalho do Excel (.xlsx)'."
        )

    # CSV
    if name.endswith(".csv"):
        for sep in (";", ",", "\t"):
            for enc in ("utf-8", "latin-1", "windows-1252"):
                try:
                    df = pd.read_csv(
                        io.BytesIO(raw),
                        sep=sep,
                        encoding=enc,
                        header=None,
                        dtype=str,
                    )
                    if df.shape[1] >= 2 and len(df) > 1:
                        return [df]
                except Exception as e:  # noqa: BLE001
                    errors.append(f"csv {sep}/{enc}: {e}")

    # XLSX moderno
    try:
        df_dict = pd.read_excel(
            io.BytesIO(raw),
            engine="openpyxl",
            sheet_name=None,
            header=None,
            dtype=object,
        )
        return list(df_dict.values())
    except Exception as e:  # noqa: BLE001
        errors.append(f"openpyxl: {e}")

    # XLS antigo (binário)
    try:
        df_dict = pd.read_excel(
            io.BytesIO(raw),
            engine="xlrd",
            sheet_name=None,
            header=None,
            dtype=object,
        )
        return list(df_dict.values())
    except Exception as e:  # noqa: BLE001
        errors.append(f"xlrd: {e}")

    # HTML disfarçado (mas não frameset)
    if _is_html_disguised_xls(raw):
        tables = _read_html_xls(raw)
        if tables:
            return tables
        errors.append("html: nenhuma tabela encontrada")

    raise ValueError(
        "Não foi possível ler a planilha. Detalhes técnicos:\n - "
        + "\n - ".join(errors)
    )


# ---------------------------------------------------------------------------
# Detecção do cabeçalho real (pode estar em qualquer linha)
# ---------------------------------------------------------------------------

# o app aceita várias palavras-chave para o mesmo conceito
_MATRICULA_KEYS = (
    "matricul", "matrícul", "chapa", "registro", "cracha", "crachá",
)
_SALDO_KEYS = (
    "saldo", "estimado", "credito", "crédito", "valor disponivel",
    "valor disponível", "disponivel", "disponível", "atual",
)
_NOME_KEYS = (
    "nome", "funcionario", "funcionário", "colaborador", "empregado",
    "titular",
)


def _cell_matches(cell, keys: Tuple[str, ...]) -> bool:
    if cell is None:
        return False
    v = _strip_accents(str(cell))
    if not v:
        return False
    return any(k in v for k in keys)


def _find_header_row(df: pd.DataFrame, max_scan: int = 30) -> Optional[int]:
    """
    Procura a linha onde estão os cabeçalhos. Critério: linha que contém PELO
    MENOS duas das três palavras-chave (matrícula / saldo / nome).
    """
    limit = min(max_scan, len(df))
    best = None
    best_score = 0
    for i in range(limit):
        row = df.iloc[i].tolist()
        has_mat = any(_cell_matches(c, _MATRICULA_KEYS) for c in row)
        has_sal = any(_cell_matches(c, _SALDO_KEYS) for c in row)
        has_nom = any(_cell_matches(c, _NOME_KEYS) for c in row)
        score = int(has_mat) + int(has_sal) + int(has_nom)
        if has_mat and has_sal and score > best_score:
            best_score = score
            best = i
    return best


def _find_col_index(row: list, keys: Tuple[str, ...]) -> Optional[int]:
    for j, v in enumerate(row):
        if _cell_matches(v, keys):
            return j
    return None


# ---------------------------------------------------------------------------
# Limpeza de matrícula e saldo
# ---------------------------------------------------------------------------

def _clean_matricula(s) -> str:
    if s is None:
        return ""
    # pandas pode entregar float (matricula 119 vira 119.0); cobre todos os casos
    if isinstance(s, float):
        if pd.isna(s):
            return ""
        # ints disfarçados de float
        if float(s).is_integer():
            return str(int(s))
        return str(s)
    if isinstance(s, int):
        return str(s)
    txt = str(s)
    # remove invisíveis e nbsp
    txt = txt.replace("\xa0", "").replace("\u200b", "").strip()
    if not txt or txt.lower() in ("nan", "none"):
        return ""
    # planilhas frequentemente vêm como "119.0" (float convertido p/ str)
    if re.match(r"^\d+\.0+$", txt):
        txt = txt.split(".")[0]
    # remove zeros à esquerda (só se for numérico)
    if txt.isdigit():
        txt = str(int(txt))
    return txt


_MONEY_CLEAN_RE = re.compile(r"[^\d,.\-]")


def _clean_name(s) -> str:
    if s is None:
        return ""
    if isinstance(s, float) and pd.isna(s):
        return ""
    return (
        str(s)
        .replace("\xa0", " ")
        .replace("\u200b", "")
        .strip()
    )


def _parse_money(v) -> float:
    """
    Aceita valores em ponto (US) ou vírgula (BR). Heurística:
      - se tiver vírgula E ponto: assume BR (ponto = milhar)
      - se tiver só vírgula: BR
      - se tiver só ponto: assume US/decimal (96.9 = 96,90)
      - se for inteiro: literal
    """
    if v is None:
        return 0.0
    if isinstance(v, (int, float)):
        if pd.isna(v):
            return 0.0
        return float(v)
    s = str(v).strip().replace("\xa0", "").replace("\u200b", "")
    if not s or s.lower() in ("nan", "none"):
        return 0.0
    s = _MONEY_CLEAN_RE.sub("", s)
    if not s or s in ("-", ",", "."):
        return 0.0

    if "," in s and "." in s:
        # 1.234,56 → BR
        s = s.replace(".", "").replace(",", ".")
    elif "," in s:
        s = s.replace(",", ".")
    # se só tem ponto, mantém como está (é decimal US)
    try:
        return float(s)
    except ValueError:
        return 0.0


# ---------------------------------------------------------------------------
# API principal
# ---------------------------------------------------------------------------

@dataclass
class SaldoTable:
    df: pd.DataFrame
    raw_columns: list[str]
    n_rows: int
    n_ignored: int = 0            # linhas ignoradas (sem matrícula etc.)
    header_row_index: int = -1    # linha onde foi achado o cabeçalho
    sheet_used: int = 0


def parse_saldo(content: bytes, filename: str) -> SaldoTable:
    tables = _try_read_any(content, filename)

    best: Optional[Tuple[pd.DataFrame, int, dict, int]] = None
    best_score = -1

    for sheet_idx, raw_df in enumerate(tables):
        if raw_df is None or raw_df.empty:
            continue
        # achata MultiIndex de colunas se vier de pd.read_html
        if isinstance(raw_df.columns, pd.MultiIndex):
            raw_df.columns = [
                " ".join([str(x) for x in tup if str(x) != "nan"]).strip()
                for tup in raw_df.columns
            ]
        # sempre resetamos para usar índice numérico
        df = raw_df.reset_index(drop=True)

        header_idx = _find_header_row(df)
        if header_idx is None:
            continue

        header_row = df.iloc[header_idx].tolist()
        col_mat = _find_col_index(header_row, _MATRICULA_KEYS)
        col_sal = _find_col_index(header_row, _SALDO_KEYS)
        col_nom = _find_col_index(header_row, _NOME_KEYS)

        if col_mat is None or col_sal is None:
            continue

        n_rows = len(df) - header_idx - 1
        score = 10 + n_rows / 100
        if col_nom is not None:
            score += 2

        if score > best_score:
            best_score = score
            best = (df, header_idx, {"mat": col_mat, "sal": col_sal, "nom": col_nom}, sheet_idx)

    if best is None:
        raise ValueError(
            "Não foi possível localizar as colunas de matrícula e saldo na "
            "planilha. Verifique se há linhas com 'Matrícula:' e 'Saldo:' "
            "no cabeçalho (esperado em qualquer das primeiras 30 linhas)."
        )

    df, header_idx, cols, sheet_idx = best
    header_row = df.iloc[header_idx].tolist()

    raw_columns = []
    for key in ("nom", "mat", "sal"):
        j = cols.get(key)
        raw_columns.append(str(header_row[j]).strip() if j is not None else "")

    # extrai os dados após o header
    data = df.iloc[header_idx + 1 :].copy().reset_index(drop=True)

    mat_series = data.iloc[:, cols["mat"]].map(_clean_matricula)
    sal_series = data.iloc[:, cols["sal"]].map(_parse_money)
    if cols.get("nom") is not None:
        nom_series = data.iloc[:, cols["nom"]].map(_clean_name)
    else:
        nom_series = pd.Series([""] * len(data))

    norm = pd.DataFrame({"matricula": mat_series, "nome": nom_series, "saldo": sal_series})

    # ---- filtros: remove lixo ----
    total_lines = len(norm)

    # 1) remove linhas sem matrícula válida
    norm = norm[norm["matricula"] != ""]
    # 2) remove se a matrícula for "matricula" (a própria palavra) ou texto
    norm = norm[~norm["matricula"].str.lower().isin({"matricula", "matrícula", "total"})]
    # 3) remove rodapés tipo "Total de Cartões"
    mask_total = norm["nome"].astype(str).str.lower().str.contains(
        r"total\s*(?:de)?\s*cart", regex=True, na=False
    )
    norm = norm[~mask_total]
    # 4) só mantém matrículas que pareçam código (mais comum: só dígitos)
    norm = norm[norm["matricula"].str.match(r"^[A-Za-z0-9]+$", na=False)]

    # 5) duplicatas — soma os saldos da MESMA matrícula
    norm = (
        norm.groupby("matricula", as_index=False)
        .agg({"saldo": "sum", "nome": "first"})
    )

    n_ignored = total_lines - len(norm)

    if norm.empty:
        raise ValueError(
            "A planilha foi lida, mas após filtrar pré-cabeçalho/rodapé não "
            "sobrou nenhum colaborador válido. Confira se as colunas de "
            "matrícula e saldo estão preenchidas."
        )

    return SaldoTable(
        df=norm,
        raw_columns=raw_columns,
        n_rows=len(norm),
        n_ignored=n_ignored,
        header_row_index=header_idx,
        sheet_used=sheet_idx,
    )
