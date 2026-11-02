"""Leitura do config.json e resolução de caminhos."""
import json
import os
from pathlib import Path

RAIZ = Path(__file__).resolve().parent.parent

PADRAO = {
    "calendario": {"fonte": ""},                    # link do Google Planilhas, ID ou caminho de um .xlsx
    "respostas": {"arquivo": "", "aba": "", "google_sheet_id": "", "google_gid": "0",
                  "google_credenciais": "credentials.json", "google_token": "token.json"},
    "controle": {"arquivo": "CONTROLE.xlsx", "modelo": ""},   # modelo: planilha antiga p/ copiar Gerências e Configurações
    "pasta_saida": "saida",
    "incluir_dia_seguinte": True,
    "campo_data_busca": "data_apoio",               # "data_apoio" (pergunta 2) ou "carimbo" (data/hora da resposta)
    "colunas": {},
    "email": {
        "remetente": "", "nome_remetente": "", "para_padrao": "",
        # a senha NÃO fica aqui: use a variável de ambiente indicada em senha_variavel_ambiente
        "smtp": {"servidor": "", "porta": 587, "usuario": "",
                 "senha_variavel_ambiente": "SMTP_SENHA", "usar_starttls": True},
    },
}


def _mesclar(base, extra):
    saida = dict(base)
    for k, v in (extra or {}).items():
        saida[k] = _mesclar(saida[k], v) if isinstance(v, dict) and isinstance(saida.get(k), dict) else v
    return saida


def carregar(arquivo=None):
    arq = Path(arquivo or os.environ.get("RELATORIOS_CONFIG") or (RAIZ / "config.json"))
    with open(arq, encoding="utf-8") as f:
        cfg = _mesclar(PADRAO, json.load(f))
    cfg["_raiz"] = str(arq.resolve().parent)
    return cfg


def caminho(cfg, relativo):
    p = Path(relativo)
    return p if p.is_absolute() else Path(cfg["_raiz"]) / p


def pasta_saida(cfg, sub=""):
    p = caminho(cfg, cfg["pasta_saida"]) / sub
    p.mkdir(parents=True, exist_ok=True)
    return p
