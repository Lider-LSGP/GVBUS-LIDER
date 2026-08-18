"""
Parsers do arquivo TXT (folha comercial) e da planilha (saldo do cartão GVBUS).

Aceita .xls (binário), .xlsx, .xls em formato HTML (export do Excel "Salvar
como Página da Web") e .csv. Detecta automaticamente colunas de matrícula,
nome e saldo na planilha, mesmo com pré-cabeçalho e rodapé de totais.

OTIMIZAÇÕES DE MEMÓRIA (v5):
- PDF lido com `pypdfium2` (motor C++ do Chromium) em vez de `pdfplumber`.
  Consome ~90% menos RAM e é 5-10x mais rápido em PDFs grandes (~500 pgs).
- Cada página do PDF é fechada explicitamente após ler o texto (close()).
- gc.collect() forçado a cada 50 páginas para liberar RAM da C-extension.
- .xls antigo lido com xlrd streaming; fallback para HTML só se realmente
  não for um binário .xls, evitando dobrar a memória por engano.
"""

from __future__ import annotations

import gc
import io
import re
import unicodedata
from dataclasses import dataclass
from typing import List, Optional, Tuple

import pandas as pd

# --- motor de PDF: pypdfium2 (leve) com fallback para pdfplumber ------------
try:
    import pypdfium2 as pdfium
    _HAS_PDFIUM = True
except ImportError:  # pragma: no cover
    _HAS_PDFIUM = False

try:
    import pdfplumber
    _HAS_PDFPLUMBER = True
except ImportError:  # pragma: no cover
    _HAS_PDFPLUMBER = False

# se pelo menos um dos dois motores estiver instalado, sabemos ler PDF
_HAS_PDF = _HAS_PDFIUM or _HAS_PDFPLUMBER


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


def _is_real_binary_xls(raw: bytes) -> bool:
    """.xls binário verdadeiro (OLE Compound File) começa com D0 CF 11 E0."""
    return raw[:4] == b"\xd0\xcf\x11\xe0"


# -- regex para o layout "uma <table> por colaborador" do GVBUS -------------
#
# O sistema do GVBUS exporta o Saldo Estimado como um .xls que é na verdade
# um HTML onde CADA COLABORADOR está em uma <table> separada, tipo:
#   <Table><tr>
#     <td colspan='2'>&#8203;&nbsp;06852890995381&nbsp;&nbsp;</td>
#     <td colspan='3' align='left'>MARIA LIZA VIANA</td>
#     <td align='left'>652</td>
#     <td align='left'>VT Funcionário</td>
#     <td align='left'>Ativo</td>
#     <td align='left'>0</td>
#   </tr></Table>
#
# Um arquivo com 3.000 colaboradores gera 3.000+ <table> — se jogássemos
# tudo em pandas.read_html() cria 3.000 DataFrames pequenos, o que é lento
# e come muita RAM. O regex abaixo extrai direto do HTML: rápido e leve.
#
_GVBUS_ROW_RE = re.compile(
    r"<table[^>]*>\s*<tr[^>]*>"
    r"\s*<td[^>]*>(?:&#8203;)?(?:&nbsp;|\s)*(?P<cartao>\d{10,25})(?:&nbsp;|\s)*</td>"
    r"\s*<td[^>]*>(?P<nome>[^<]*)</td>"
    r"\s*<td[^>]*>(?P<matricula>[^<]*)</td>"
    r"\s*<td[^>]*>(?P<tipo>[^<]*)</td>"
    r"\s*<td[^>]*>(?P<status>[^<]*)</td>"
    r"\s*<td[^>]*>(?P<saldo>[^<]*)</td>",
    re.IGNORECASE,
)


def _clean_html_cell(s: str) -> str:
    """Limpa uma célula do HTML do GVBUS (remove &nbsp;, &#8203;, tags, acentos HTML)."""
    if not s:
        return ""
    s = s.replace("&nbsp;", " ").replace("&#8203;", "").replace("\xa0", " ")
    # remove tags residuais
    s = re.sub(r"<[^>]+>", "", s)
    # decodifica entidades HTML básicas comuns em pt-BR
    ents = {
        "&aacute;": "á", "&eacute;": "é", "&iacute;": "í",
        "&oacute;": "ó", "&uacute;": "ú", "&atilde;": "ã",
        "&otilde;": "õ", "&ccedil;": "ç", "&Aacute;": "Á",
        "&Eacute;": "É", "&Iacute;": "Í", "&Oacute;": "Ó",
        "&Uacute;": "Ú", "&Atilde;": "Ã", "&Otilde;": "Õ",
        "&Ccedil;": "Ç", "&amp;": "&", "&lt;": "<", "&gt;": ">",
        "&quot;": '"', "&apos;": "'",
    }
    for k, v in ents.items():
        s = s.replace(k, v)
    return re.sub(r"\s+", " ", s).strip()


def _read_gvbus_html_saldo(text: str) -> Optional[pd.DataFrame]:
    """Extrai o relatório de Saldo Estimado do GVBUS direto por regex.
    Retorna um DataFrame no MESMO formato do parser de PDF (para o resto
    do código não precisar saber a origem)."""
    records: list[dict] = []
    for m in _GVBUS_ROW_RE.finditer(text):
        cartao = _clean_html_cell(m.group("cartao"))
        nome = _clean_html_cell(m.group("nome"))
        matricula = _clean_html_cell(m.group("matricula"))
        tipo = _clean_html_cell(m.group("tipo"))
        status = _clean_html_cell(m.group("status"))
        saldo = _clean_html_cell(m.group("saldo"))

        # ignora linhas de cabeçalho (com títulos em vez de dados)
        if not cartao.isdigit():
            continue
        if "funcionário" in nome.lower() and "matrícula" in matricula.lower():
            continue

        records.append({
            "Cartão": cartao,
            "Funcionário": nome,
            "Matrícula": matricula,
            "Tipo": tipo,
            "Status": status,
            "Saldo": saldo,
        })

    if not records:
        return None

    df = pd.DataFrame(records)
    # linha 0 = cabeçalho "fake" (o _find_header_row do detector procura por
    # 'Matrícula:', 'Saldo:', etc. na primeira linha; alimentamos isso).
    header = pd.DataFrame(
        [["Cartão:", "Funcionário:", "Matrícula:", "Tipo Utilização:", "Status:", "Saldo:"]],
        columns=df.columns,
    )
    return pd.concat([header, df], ignore_index=True)


def _read_html_xls(raw: bytes) -> List[pd.DataFrame]:
    """Lê um .xls que na verdade é HTML (Excel "Salvar como Página da Web"
    OU export do sistema GVBUS).

    Estratégia (ordem):
      1) Tenta o regex do formato GVBUS (uma <table> por colaborador).
         Se casar pelo menos 1 linha útil, usa esse resultado.
      2) Cai para pandas.read_html() para HTMLs comuns.
    """
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

    # 1) TENTATIVA GVBUS: extrai por regex (LEVE, muito rápido)
    df_gvbus = _read_gvbus_html_saldo(text)
    if df_gvbus is not None and len(df_gvbus) > 1:
        gc.collect()
        return [df_gvbus]

    # 2) FALLBACK: pandas.read_html para HTMLs "normais"
    try:
        tables = pd.read_html(io.StringIO(text), decimal=",", thousands=".")
    except ValueError:
        tables = []

    # se voltou muitas tabelas pequenininhas (>50), concatena tudo em uma só
    # para o detector achar o cabeçalho
    if len(tables) > 50:
        try:
            big = pd.concat(tables, ignore_index=True)
            tables = [big]
        except Exception:
            pass

    gc.collect()
    return tables


class FramesetXlsError(ValueError):
    """Erro específico: o usuário enviou um .xls 'frameset' sem a pasta
    auxiliar. Mensagem amigável no Streamlit."""


# ---------------------------------------------------------------------------
# Parser do PDF (relatório do sistema GVBUS) — VERSÃO LEVE COM PYPDFIUM2
# ---------------------------------------------------------------------------

_PDF_LINE_FULL = re.compile(
    r"^\s*(?P<cartao>\d{14,20})\s+"
    r"(?P<nome>.+?)\s+"
    r"(?P<matricula>\d{1,7})\s+"
    r"(?P<tipo>VT|VR|VA|VTP)\s+"
    r"\S+\s+"
    r"(?P<status>Ativo|Bloqueado|Cancelado|Inativo|Suspenso)\s+"
    r"(?P<saldo>-?[\d.]+,\d{2})\s*$",
    re.IGNORECASE,
)

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


def _parse_pdf_lines(text: str, records: list[dict]) -> None:
    """Aplica os regexes numa string de texto e acumula em `records`.
    Isolado para poder ser reutilizado pelos dois motores (pdfium/plumber)."""
    for line in text.splitlines():
        line = line.strip()
        if not line:
            continue
        low = line.lower()
        if any(low.startswith(p) for p in _PDF_SKIP_PREFIXES):
            continue

        m = _PDF_LINE_FULL.match(line)
        if m:
            records.append({
                "Cartão": m.group("cartao"),
                "Funcionário": m.group("nome").strip(),
                "Matrícula": m.group("matricula"),
                "Tipo": m.group("tipo"),
                "Status": m.group("status"),
                "Saldo": m.group("saldo"),
            })
            continue

        m = _PDF_LINE_NONAME.match(line)
        if m:
            records.append({
                "Cartão": m.group("cartao"),
                "Funcionário": "",
                "Matrícula": "",
                "Tipo": m.group("tipo"),
                "Status": m.group("status"),
                "Saldo": m.group("saldo"),
            })


def _read_pdf_saldo_pdfium(raw: bytes) -> Optional[pd.DataFrame]:
    """
    Motor LEVE — pypdfium2. Consome ~90% menos RAM que pdfplumber e roda
    5-10x mais rápido, o que permite processar PDFs de 500+ páginas dentro
    do limite de 1 GB do Streamlit Cloud gratuito.
    """
    records: list[dict] = []
    pdf = pdfium.PdfDocument(raw)
    try:
        n_pages = len(pdf)
        for i in range(n_pages):
            page = pdf[i]
            textpage = page.get_textpage()
            try:
                text = textpage.get_text_range() or ""
            finally:
                # SEMPRE fecha o textpage e a página antes de ir pra próxima
                textpage.close()
                page.close()
            _parse_pdf_lines(text, records)

            # limpa a RAM da C-extension a cada 50 páginas
            if (i + 1) % 50 == 0:
                gc.collect()
    finally:
        pdf.close()
        del pdf
        gc.collect()

    if not records:
        return None

    df = pd.DataFrame(records)
    header = pd.DataFrame(
        [["Cartão:", "Funcionário:", "Matrícula:", "Tipo Utilização:", "Status:", "Saldo:"]],
        columns=df.columns,
    )
    return pd.concat([header, df], ignore_index=True)


def _read_pdf_saldo_pdfplumber(raw: bytes) -> Optional[pd.DataFrame]:
    """Fallback: pdfplumber com flush_cache() e gc.collect() por página.
    Só é chamado se pypdfium2 não estiver disponível."""
    records: list[dict] = []
    with pdfplumber.open(io.BytesIO(raw)) as pdf:
        n_pages = len(pdf.pages)
        for i, page in enumerate(pdf.pages):
            text = page.extract_text() or ""
            _parse_pdf_lines(text, records)
            # limpa o cache interno de cada página lida
            try:
                page.flush_cache()
            except Exception:
                pass
            if (i + 1) % 20 == 0:
                gc.collect()

    gc.collect()
    if not records:
        return None
    df = pd.DataFrame(records)
    header = pd.DataFrame(
        [["Cartão:", "Funcionário:", "Matrícula:", "Tipo Utilização:", "Status:", "Saldo:"]],
        columns=df.columns,
    )
    return pd.concat([header, df], ignore_index=True)


def _read_pdf_saldo(raw: bytes) -> Optional[pd.DataFrame]:
    """
    Lê o PDF do relatório GVBUS. Usa pypdfium2 se disponível (recomendado,
    muito mais leve em RAM); cai para pdfplumber se não estiver.
    """
    if _HAS_PDFIUM:
        return _read_pdf_saldo_pdfium(raw)
    if _HAS_PDFPLUMBER:
        return _read_pdf_saldo_pdfplumber(raw)
    raise ValueError(
        "Nenhum motor de PDF instalado. Adicione 'pypdfium2' ao requirements.txt."
    )


def _try_read_any(raw: bytes, filename: str) -> List[pd.DataFrame]:
    """Tenta ler o arquivo em todos os formatos possíveis (sempre header=None,
    pois o cabeçalho real será detectado depois).

    ORDEM DE TENTATIVA (importante p/ memória):
      1) PDF   → pypdfium2 (streaming, leve)
      2) frameset xls → erro amigável
      3) CSV   → pandas
      4) XLS binário real (OLE) → xlrd  (checa o magic byte antes!)
      5) XLSX  → openpyxl
      6) HTML disfarçado (só se NÃO for binário xls)
    """
    name = (filename or "").lower()
    errors: list[str] = []

    # 1) PDF
    if name.endswith(".pdf") or raw[:4] == b"%PDF":
        if not _HAS_PDF:
            raise ValueError(
                "Para ler PDFs, instale a dependência 'pypdfium2' "
                "(pip install pypdfium2)."
            )
        df = _read_pdf_saldo(raw)
        if df is None or df.empty:
            raise ValueError(
                "PDF lido, mas não foi possível extrair nenhum colaborador. "
                "O formato do PDF mudou? Esperado: 'cartão nome matrícula "
                "VT Funcionário Status saldo'."
            )
        return [df]

    # 2) detecta o frameset xls ANTES de tentar outros formatos
    if _is_frameset_xls(raw):
        raise FramesetXlsError(
            "O arquivo enviado é um '.xls' do tipo 'Página da Web' (frameset) "
            "do Excel 97-2003 e está vazio — os dados ficam em uma pasta "
            "auxiliar que não veio junto. Abra o arquivo no Excel e use "
            "'Salvar como → Pasta de Trabalho do Excel (.xlsx)'."
        )

    # 3) CSV
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

    # 4) XLS binário REAL (OLE) — só tenta xlrd se o header for OLE
    if _is_real_binary_xls(raw):
        try:
            df_dict = pd.read_excel(
                io.BytesIO(raw),
                engine="xlrd",
                sheet_name=None,
                header=None,
                dtype=object,
            )
            result = list(df_dict.values())
            gc.collect()
            return result
        except Exception as e:  # noqa: BLE001
            errors.append(f"xlrd: {e}")

    # 5) XLSX moderno
    try:
        df_dict = pd.read_excel(
            io.BytesIO(raw),
            engine="openpyxl",
            sheet_name=None,
            header=None,
            dtype=object,
        )
        result = list(df_dict.values())
        gc.collect()
        return result
    except Exception as e:  # noqa: BLE001
        errors.append(f"openpyxl: {e}")

    # 6) HTML disfarçado (mas não frameset, e não é xls binário)
    if _is_html_disguised_xls(raw) and not _is_real_binary_xls(raw):
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
    if isinstance(s, float):
        if pd.isna(s):
            return ""
        if float(s).is_integer():
            return str(int(s))
        return str(s)
    if isinstance(s, int):
        return str(s)
    txt = str(s)
    txt = txt.replace("\xa0", "").replace("\u200b", "").strip()
    if not txt or txt.lower() in ("nan", "none"):
        return ""
    if re.match(r"^\d+\.0+$", txt):
        txt = txt.split(".")[0]
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
        s = s.replace(".", "").replace(",", ".")
    elif "," in s:
        s = s.replace(",", ".")
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
    n_ignored: int = 0
    header_row_index: int = -1
    sheet_used: int = 0


def parse_saldo(content: bytes, filename: str) -> SaldoTable:
    tables = _try_read_any(content, filename)

    best: Optional[Tuple[pd.DataFrame, int, dict, int]] = None
    best_score = -1

    for sheet_idx, raw_df in enumerate(tables):
        if raw_df is None or raw_df.empty:
            continue
        if isinstance(raw_df.columns, pd.MultiIndex):
            raw_df.columns = [
                " ".join([str(x) for x in tup if str(x) != "nan"]).strip()
                for tup in raw_df.columns
            ]
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

    norm = norm[norm["matricula"] != ""]
    norm = norm[~norm["matricula"].str.lower().isin({"matricula", "matrícula", "total"})]
    mask_total = norm["nome"].astype(str).str.lower().str.contains(
        r"total\s*(?:de)?\s*cart", regex=True, na=False
    )
    norm = norm[~mask_total]
    norm = norm[norm["matricula"].str.match(r"^[A-Za-z0-9]+$", na=False)]

    norm = (
        norm.groupby("matricula", as_index=False)
        .agg({"saldo": "sum", "nome": "first"})
    )

    n_ignored = total_lines - len(norm)

    # libera memória dos DataFrames intermediários grandes
    del tables, df, data, mat_series, sal_series, nom_series
    gc.collect()

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


# ===========================================================================
# Parser da planilha do AppLider (Matrícula + Nome + Tipo Escala)
# ===========================================================================

@dataclass
class AppLiderRow:
    matricula: str
    nome: str
    escala_raw: str
    escala: str
    empresa: str = ""
    ativo: str = ""
    posto: str = ""


@dataclass
class AppLiderTable:
    df: pd.DataFrame
    n_rows: int
    n_ignored: int
    header_row_index: int
    sheet_used: int
    raw_columns: dict


_APPLIDER_MAT_KEYS = ("matricula/modal", "matrícula/modal",
                       "matriculamodal", "matrículamodal", "modal")
_APPLIDER_NOME_KEYS = ("nome",)
_APPLIDER_ESCALA_KEYS = ("tipo escala", "tipo/escala", "escala")
_APPLIDER_EMPRESA_KEYS = ("empresa",)
_APPLIDER_ATIVO_KEYS = ("ativo",)
_APPLIDER_POSTO_KEYS = ("posto trabalho", "posto de trabalho", "posto")


def _normalize_applider_cell(v) -> str:
    if v is None:
        return ""
    if isinstance(v, float) and pd.isna(v):
        return ""
    s = str(v)
    s = s.replace("\xa0", " ").replace("\u200b", "").replace("\n", " ")
    s = (
        s.replace("&aacute;", "á").replace("&eacute;", "é")
         .replace("&iacute;", "í").replace("&oacute;", "ó")
         .replace("&uacute;", "ú").replace("&atilde;", "ã")
         .replace("&otilde;", "õ").replace("&ccedil;", "ç")
         .replace("&Aacute;", "Á").replace("&Eacute;", "É")
         .replace("&Iacute;", "Í").replace("&Oacute;", "Ó")
         .replace("&Uacute;", "Ú").replace("&Atilde;", "Ã")
         .replace("&Otilde;", "Õ").replace("&Ccedil;", "Ç")
    )
    return re.sub(r"\s+", " ", s).strip()


def _find_applider_column(header_row: list, keys: tuple[str, ...]) -> Optional[int]:
    candidates: list[tuple[int, int]] = []
    for j, v in enumerate(header_row):
        norm = _strip_accents(str(v)).replace(" ", "")
        for k in keys:
            kn = k.replace(" ", "")
            if kn in norm:
                candidates.append((j, len(kn)))
                break
    if not candidates:
        return None
    candidates.sort(key=lambda t: (-t[1], t[0]))
    return candidates[0][0]


def parse_applider(content: bytes, filename: str) -> AppLiderTable:
    from .escala import normalizar_escala

    tables = _try_read_any(content, filename)
    if not tables:
        raise ValueError("Não foi possível ler a planilha do AppLider.")

    best = None
    best_score = -1
    for sheet_idx, raw_df in enumerate(tables):
        if raw_df is None or raw_df.empty:
            continue
        if isinstance(raw_df.columns, pd.MultiIndex):
            raw_df.columns = [
                " ".join([str(x) for x in tup if str(x) != "nan"]).strip()
                for tup in raw_df.columns
            ]
        df = raw_df.reset_index(drop=True)

        header_idx = None
        for i in range(min(15, len(df))):
            row = df.iloc[i].tolist()
            has_mat = any(
                ("modal" in _strip_accents(str(c)).replace(" ", ""))
                for c in row
            )
            has_nom = any(_cell_matches(c, _APPLIDER_NOME_KEYS) for c in row)
            has_esc = any(_cell_matches(c, _APPLIDER_ESCALA_KEYS) for c in row)
            if has_mat and has_nom and has_esc:
                header_idx = i
                break

        if header_idx is None:
            continue

        header_row = df.iloc[header_idx].tolist()
        col_mat = _find_applider_column(header_row, _APPLIDER_MAT_KEYS)
        col_nom = _find_applider_column(header_row, _APPLIDER_NOME_KEYS)
        col_esc = _find_applider_column(header_row, _APPLIDER_ESCALA_KEYS)
        col_emp = _find_applider_column(header_row, _APPLIDER_EMPRESA_KEYS)
        col_ativo = _find_applider_column(header_row, _APPLIDER_ATIVO_KEYS)
        col_posto = _find_applider_column(header_row, _APPLIDER_POSTO_KEYS)

        n_rows = len(df) - header_idx - 1
        score = n_rows

        if score > best_score:
            best_score = score
            best = (df, header_idx, sheet_idx, {
                "mat": col_mat, "nom": col_nom, "esc": col_esc,
                "emp": col_emp, "ativo": col_ativo, "posto": col_posto,
            })

    if best is None:
        raise ValueError(
            "Não foi possível localizar as colunas necessárias na planilha do "
            "AppLider. O cabeçalho precisa conter, nas primeiras 15 linhas, "
            "as colunas **Matrícula/Modal** (número do cartão GVBUS), "
            "**Nome** e **Tipo Escala**."
        )

    df, header_idx, sheet_idx, cols = best
    header_row = df.iloc[header_idx].tolist()
    raw_columns = {
        k: (_normalize_applider_cell(header_row[cols[k]]) if cols[k] is not None else "")
        for k in cols
    }

    data = df.iloc[header_idx + 1:].copy().reset_index(drop=True)
    total_lines = len(data)

    matricula = data.iloc[:, cols["mat"]].map(_clean_matricula)
    nome = data.iloc[:, cols["nom"]].map(_normalize_applider_cell)
    escala_raw = data.iloc[:, cols["esc"]].map(_normalize_applider_cell)
    escala = escala_raw.map(normalizar_escala)

    def _series_or_blank(idx):
        if idx is None:
            return pd.Series([""] * len(data))
        return data.iloc[:, idx].map(_normalize_applider_cell)

    empresa = _series_or_blank(cols["emp"])
    ativo = _series_or_blank(cols["ativo"])
    posto = _series_or_blank(cols["posto"])

    norm = pd.DataFrame({
        "matricula": matricula,
        "nome": nome,
        "escala_raw": escala_raw,
        "escala": escala,
        "empresa": empresa,
        "ativo": ativo,
        "posto": posto,
    })

    norm = norm[norm["matricula"] != ""]
    norm = norm[norm["matricula"].str.match(r"^[A-Za-z0-9]+$", na=False)]
    norm = norm.sort_values(
        by=["matricula", "escala"],
        key=lambda s: s.map(lambda x: 1 if x == "DESCONHECIDA" else 0)
        if s.name == "escala" else s
    )
    norm = norm.drop_duplicates(subset=["matricula"], keep="first").reset_index(drop=True)

    n_ignored = total_lines - len(norm)

    # libera memória dos DataFrames intermediários
    del tables, df, data, matricula, nome, escala_raw, escala, empresa, ativo, posto
    gc.collect()

    if norm.empty:
        raise ValueError("A planilha do AppLider está sem colaboradores válidos.")

    return AppLiderTable(
        df=norm,
        n_rows=len(norm),
        n_ignored=n_ignored,
        header_row_index=header_idx,
        sheet_used=sheet_idx,
        raw_columns=raw_columns,
    )
