"""Leitura PRIVADA da planilha de respostas via API do Google Planilhas (somente leitura).

Usa login OAuth com a conta corporativa. A primeira autorização abre o navegador; depois
o token fica salvo em token.json e é renovado sozinho.
"""
import json
from pathlib import Path

ESCOPO = ["https://www.googleapis.com/auth/spreadsheets.readonly"]
API = "https://sheets.googleapis.com/v4/spreadsheets"


class ErroGoogle(Exception):
    pass


def _arquivos(cfg):
    raiz = Path(cfg["_raiz"])
    g = cfg["respostas"]
    cred = Path(g.get("google_credenciais") or "credentials.json")
    tok = Path(g.get("google_token") or "token.json")
    return (cred if cred.is_absolute() else raiz / cred), (tok if tok.is_absolute() else raiz / tok)


def disponivel(cfg):
    """True se existe credentials.json (ou token.json) para usar a API privada."""
    cred, tok = _arquivos(cfg)
    return cred.exists() or tok.exists()


def credenciais(cfg, interativo=False):
    """Devolve credenciais válidas. Com interativo=True abre o navegador se precisar de login."""
    try:
        from google.auth.transport.requests import Request
        from google.oauth2.credentials import Credentials
        from google_auth_oauthlib.flow import InstalledAppFlow
    except ImportError:
        raise ErroGoogle("Bibliotecas do Google não instaladas. Rode: pip install -r requirements.txt")
    cred_arq, tok_arq = _arquivos(cfg)
    cred = None
    if tok_arq.exists():
        try:
            cred = Credentials.from_authorized_user_file(str(tok_arq), ESCOPO)
        except Exception:
            cred = None
    if cred and cred.valid:
        return cred
    if cred and cred.expired and cred.refresh_token:
        try:
            cred.refresh(Request())
            tok_arq.write_text(cred.to_json(), encoding="utf-8")
            return cred
        except Exception:
            cred = None  # token revogado/expirado: precisa autorizar de novo
    if not interativo:
        raise ErroGoogle("Acesso ao Google não autorizado ou expirado. Rode 'python conectar_google.py' uma vez.")
    if not cred_arq.exists():
        raise ErroGoogle(f"Arquivo {cred_arq.name} não encontrado na pasta do programa (veja o README, item 3).")
    fluxo = InstalledAppFlow.from_client_secrets_file(str(cred_arq), ESCOPO)
    cred = fluxo.run_local_server(port=0, prompt="consent")
    tok_arq.write_text(cred.to_json(), encoding="utf-8")
    return cred


def _get(sessao, url, params=None):
    r = sessao.get(url, params=params, timeout=30)
    if r.status_code != 200:
        try:
            msg = r.json().get("error", {}).get("message", r.text)
        except Exception:
            msg = r.text[:200]
        dica = ""
        if r.status_code == 403:
            dica = " (sua conta não tem acesso à planilha, ou a Google Sheets API não está ativada no projeto)"
        elif r.status_code == 404:
            dica = " (confira o google_sheet_id)"
        raise ErroGoogle(f"Google respondeu {r.status_code}: {msg}{dica}")
    return r.json()


def ler_respostas(cfg, interativo=False):
    """Devolve (linhas, nome_da_aba). Primeira linha = cabeçalho. Valores como texto formatado (igual ao CSV)."""
    from google.auth.transport.requests import AuthorizedSession
    g = cfg["respostas"]
    if not g.get("google_sheet_id"):
        raise ErroGoogle("google_sheet_id não preenchido no config.json.")
    sessao = AuthorizedSession(credenciais(cfg, interativo))
    sid = g["google_sheet_id"]
    meta = _get(sessao, f"{API}/{sid}", {"fields": "properties(title,locale),sheets.properties(sheetId,title)"})
    loc = (meta.get("properties", {}).get("locale") or "").lower()
    if loc and not loc.startswith("pt"):
        raise ErroGoogle(
            f"A planilha está com a região '{loc}', e as datas viriam em formato diferente de dd/mm/aaaa. "
            "No Google Planilhas: Arquivo > Configurações > Região = Brasil."
        )
    abas = [s["properties"] for s in meta.get("sheets", [])]
    aba = None
    if g.get("aba"):  # nome da aba tem prioridade
        aba = next((a["title"] for a in abas if a["title"].strip().lower() == g["aba"].strip().lower()), None)
        if aba is None:
            raise ErroGoogle(f"Aba '{g['aba']}' não existe. Abas encontradas: {', '.join(a['title'] for a in abas)}")
    else:
        gid = str(g.get("google_gid") or "0")
        aba = next((a["title"] for a in abas if str(a["sheetId"]) == gid), None)
        if aba is None:
            raise ErroGoogle(f"Nenhuma aba com gid={gid}. Abas: " + ", ".join(f"{a['title']} (gid {a['sheetId']})" for a in abas))
    aba_a1 = "'" + aba.replace("'", "''") + "'"
    url = f"{API}/{sid}/values/{aba_a1}"
    dados = _get(sessao, url, {"valueRenderOption": "FORMATTED_VALUE", "dateTimeRenderOption": "FORMATTED_STRING"})
    return dados.get("values", []), aba, meta.get("properties", {}).get("title", "")
