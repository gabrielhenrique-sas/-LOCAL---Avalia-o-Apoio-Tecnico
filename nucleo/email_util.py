"""Montagem, rascunho (.eml) e envio por SMTP."""
import os
import re
import smtplib
import ssl
from email.message import EmailMessage
from email.utils import formataddr, formatdate, getaddresses
from pathlib import Path

_EMAIL_RE = re.compile(r"^[^@\s,;<>]+@[^@\s,;<>]+\.[^@\s,;<>]+$")


def separar_emails(texto):
    """Aceita separação por vírgula, ponto e vírgula ou quebra de linha. Devolve (válidos, inválidos)."""
    bruto = (texto or "").replace(";", ",").replace("\n", ",")
    validos, invalidos, vistos = [], [], set()
    for _, end in getaddresses([bruto]):
        end = end.strip()
        if not end:
            continue
        if _EMAIL_RE.match(end):
            if end.lower() not in vistos:
                vistos.add(end.lower())
                validos.append(end)
        else:
            invalidos.append(end)
    return validos, invalidos


def montar_mensagem(remetente, nome_remetente, para, cc, cco, assunto, corpo, pdf=None, rascunho=False):
    msg = EmailMessage()
    if remetente:
        msg["From"] = formataddr((nome_remetente or "", remetente))
    if para:
        msg["To"] = ", ".join(para)
    if cc:
        msg["Cc"] = ", ".join(cc)
    if rascunho and cco:
        msg["Bcc"] = ", ".join(cco)  # só no rascunho; no envio SMTP os CCO ficam apenas no envelope
    if rascunho:
        msg["X-Unsent"] = "1"  # o Outlook abre como rascunho
    msg["Subject"] = assunto
    msg["Date"] = formatdate(localtime=True)
    msg.set_content(corpo or "")
    if pdf:
        dados = Path(pdf).read_bytes()
        msg.add_attachment(dados, maintype="application", subtype="pdf", filename=Path(pdf).name)
    return msg


def _senha(smtp):
    return os.environ.get(smtp.get("senha_variavel_ambiente") or "SMTP_SENHA", "")


def smtp_configurado(cfg):
    smtp = cfg["email"]["smtp"]
    return bool(smtp.get("servidor") and smtp.get("usuario") and _senha(smtp) and cfg["email"].get("remetente"))


def enviar_smtp(cfg, msg, destinatarios):
    smtp = cfg["email"]["smtp"]
    if not smtp_configurado(cfg):
        raise RuntimeError("SMTP não configurado (servidor, usuário e remetente no config.json; senha na variável de ambiente SMTP_SENHA).")
    porta = int(smtp.get("porta") or 587)
    ctx = ssl.create_default_context()
    if porta == 465:
        servidor = smtplib.SMTP_SSL(smtp["servidor"], porta, context=ctx, timeout=60)
    else:
        servidor = smtplib.SMTP(smtp["servidor"], porta, timeout=60)
        if smtp.get("usar_starttls", True):
            servidor.starttls(context=ctx)
    try:
        servidor.login(smtp["usuario"], _senha(smtp))
        servidor.send_message(msg, from_addr=cfg["email"]["remetente"], to_addrs=list(destinatarios))
    finally:
        try:
            servidor.quit()
        except Exception:
            pass
