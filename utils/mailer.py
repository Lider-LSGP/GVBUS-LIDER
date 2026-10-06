"""
utils/mailer.py — Envio automático dos relatórios por e-mail.

Configuração (NÃO comitar senha no git):
  • Local: criar o arquivo `.streamlit/secrets.toml` (está no .gitignore)
  • Streamlit Cloud: Settings → Secrets → colar o conteúdo

  [smtp]
  host = "smtp.gmail.com"
  port = 465
  user = "liderlsgp@gmail.com"
  password = "SENHA_DE_APP_DO_GMAIL"   # myaccount.google.com → Segurança → Senhas de app
  from_name = "GVBUS Comparator — Líder Limpe"
  default_to = "liderlsgp@gmail.com"

Se a seção [smtp] não existir, o app simplesmente mostra os botões de
download e informa que o envio automático está desativado — nada quebra.
"""

from __future__ import annotations

import smtplib
import ssl
from email.message import EmailMessage
from typing import List, Tuple

import streamlit as st

Attachment = Tuple[str, bytes, str]  # (filename, content_bytes, mime_type)


def _cfg():
    try:
        return st.secrets.get("smtp", None)
    except Exception:
        return None


def mailer_status() -> tuple[bool, str]:
    """Retorna (disponível, mensagem)."""
    c = _cfg()
    if not c:
        return False, "seção [smtp] não encontrada nos secrets"
    missing = [k for k in ("host", "user", "password") if not c.get(k)]
    if missing:
        return False, f"campos ausentes no [smtp]: {', '.join(missing)}"
    return True, "ok"


def send_report_email(to: str, subject: str, body_text: str,
                      attachments: List[Attachment]) -> tuple[bool, str]:
    ok, msg = mailer_status()
    if not ok:
        return False, msg
    if not to or "@" not in to:
        return False, f"destinatário inválido: {to!r}"
    c = _cfg()
    m = EmailMessage()
    m["From"] = f"{c.get('from_name', 'GVBUS Comparator')} <{c['user']}>"
    m["To"] = to
    m["Subject"] = subject
    m.set_content(body_text)
    for fname, data, mime in attachments:
        maintype, subtype = mime.split("/", 1)
        m.add_attachment(data, maintype=maintype, subtype=subtype, filename=fname)
    ctx = ssl.create_default_context()
    try:
        with smtplib.SMTP_SSL(c["host"], int(c.get("port", 465)),
                              context=ctx, timeout=90) as s:
            s.login(c["user"], c["password"])
            s.send_message(m)
        return True, f"E-mail enviado para {to}"
    except smtplib.SMTPAuthenticationError:
        return False, ("Falha de autenticação SMTP. Se for Gmail, use uma "
                       "**Senha de app** (não a senha normal da conta).")
    except Exception as e:  # noqa: BLE001
        return False, f"Erro SMTP: {e}"
