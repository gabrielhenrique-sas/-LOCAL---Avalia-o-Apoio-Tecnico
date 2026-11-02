"""Leitura das respostas do formulário (Google privado, link público ou arquivo local) e seleção por município/data."""
import csv
import io
import urllib.parse
import urllib.request
from datetime import timedelta

from openpyxl import load_workbook

from . import google_planilhas as google
from .config import caminho
from .util import norm, parse_carimbo, parse_data


class ErroDados(Exception):
    """Erro com mensagem pronta para mostrar ao usuário."""


# ---------------------------------------------------------------- respostas
def _csv_para_linhas(texto):
    texto = texto.lstrip("﻿")
    amostra = texto[:3000]
    delim = ";" if amostra.count(";") > amostra.count(",") else ","
    return list(csv.reader(io.StringIO(texto), delimiter=delim))


def url_google_publico(r):
    """URL CSV da planilha compartilhada por link. Usa o nome da aba (config 'aba') se houver; senão o gid."""
    sid = r["google_sheet_id"]
    if r.get("aba"):
        q = urllib.parse.urlencode({"tqx": "out:csv", "sheet": r["aba"]})
        return f"https://docs.google.com/spreadsheets/d/{sid}/gviz/tq?{q}"
    gid = str(r.get("google_gid") or "0")
    return f"https://docs.google.com/spreadsheets/d/{sid}/export?format=csv&gid={gid}"


def ler_google_publico(r):
    """Lê a planilha pelo link público. Levanta ErroDados se o Google pedir login ou der erro."""
    url = url_google_publico(r)
    req = urllib.request.Request(url, headers={"User-Agent": "Mozilla/5.0"})
    with urllib.request.urlopen(req, timeout=30) as resp:
        texto = resp.read().decode("utf-8", errors="replace")
    if texto.lstrip().lower().startswith(("<!doctype", "<html")):
        raise ErroDados(
            "o Google pediu login. Compartilhe a planilha como 'qualquer pessoa com o link pode ver' "
            "ou use o arquivo baixado (.xlsx)"
        )
    return _csv_para_linhas(texto)


def ler_tabela_respostas(cfg):
    """Devolve (linhas, avisos). A primeira linha é o cabeçalho."""
    r = cfg["respostas"]
    avisos = []
    arq = caminho(cfg, r["arquivo"]) if r.get("arquivo") else None
    if r.get("google_sheet_id") and google.disponivel(cfg):  # conexão privada (conta corporativa)
        try:
            linhas, _, _ = google.ler_respostas(cfg)
            return linhas, avisos
        except Exception as e:
            avisos.append(f"Falha na conexão privada com o Google ({e}). Tentando o link público/arquivo local.")
    if r.get("google_sheet_id"):
        try:
            return ler_google_publico(r), avisos
        except Exception as e:  # rede, login, etc.
            if arq and arq.exists():
                avisos.append(f"Não foi possível ler direto do Google ({e}). Usando o arquivo local.")
            else:
                raise ErroDados(f"Não foi possível ler a planilha de respostas no Google: {e}")
    if not arq or not arq.exists():
        raise ErroDados(f"Arquivo de respostas não encontrado: {arq}. Ajuste 'respostas' no config.json.")
    if arq.suffix.lower() in (".csv", ".txt"):
        return _csv_para_linhas(arq.read_text(encoding="utf-8-sig", errors="replace")), avisos
    wb = load_workbook(arq, read_only=True, data_only=True)
    try:
        ws = wb[r["aba"]] if r.get("aba") in wb.sheetnames else wb.worksheets[0]
        linhas = [list(row) for row in ws.iter_rows(values_only=True)]
    finally:
        wb.close()
    while linhas and all(c in (None, "") for c in linhas[-1]):
        linhas.pop()
    return linhas, avisos


_REGRAS = [
    ("carimbo", ["carimbo"]),
    ("carimbo", ["hora de inicio"]),  # Microsoft Forms (a "Hora de conclusão" vem em formato m/d/aa)
    ("municipio", ["nome do municipio"]),
    ("data_apoio", ["data do apoio"]),
    ("setor", ["unidade", "setor"]),
    ("esclareceu", ["esclareceu"]),
    ("esclareceu", ["esclarecer"]),
    ("formato", ["formato do apoio"]),
    ("tempo", ["tempo de duracao"]),
    ("explicar", ["forma de explicar"]),
    ("geral", ["no geral"]),
    ("comentario_geral", ["comentarios, elogios"]),
]
_OBRIGATORIOS = ["municipio", "data_apoio", "esclareceu", "formato", "tempo", "explicar", "geral"]
_ABERTAS = ["comentario_esclareceu", "comentario_formato", "comentario_explicar"]


def mapear_colunas(cabecalho, sobrescritas=None):
    """Localiza cada pergunta pelo título. Devolve (mapa campo->índice, avisos)."""
    mapa, avisos, abertas = {}, [], []
    for j, h in enumerate(cabecalho):
        t = norm(h)
        if not t:
            continue
        if "caso tenha respondido" in t or "por gentileza" in t:
            abertas.append(j)
            continue
        for campo, palavras in _REGRAS:
            if campo not in mapa and all(p in t for p in palavras):
                mapa[campo] = j
                break
    if len(abertas) >= 4:
        # ignora o comentário da pergunta que não entra no relatório ("atenção para ouvir as necessidades");
        # se não achar pelo título, ignora a 3ª (comportamento antigo)
        extra = next((j for j in abertas if "necessidades" in norm(cabecalho[j])), abertas[2])
        abertas = [j for j in abertas if j != extra][:3]
        #avisos.append(f"Há {len(abertas) + 1}+ colunas de comentário; ignorei a da coluna {extra + 1}.")
    for campo, j in zip(_ABERTAS, abertas):
        mapa[campo] = j
    if len(abertas) < 3:
        avisos.append("Menos de 3 colunas de comentário 'Caso tenha respondido...' foram encontradas.")
    for campo, alvo in (sobrescritas or {}).items():
        if isinstance(alvo, int):
            mapa[campo] = alvo
        else:
            for j, h in enumerate(cabecalho):
                if norm(h) == norm(alvo):
                    mapa[campo] = j
                    break
    faltando = [c for c in _OBRIGATORIOS if c not in mapa]
    if faltando:
        raise ErroDados(
            "Não encontrei no formulário as colunas: " + ", ".join(faltando)
            + ". Confira os títulos das perguntas (ou use 'colunas' no config.json)."
        )
    for opcional in ("setor", "carimbo"):
        if opcional not in mapa:
            avisos.append(f"Coluna opcional não encontrada: {opcional}.")
    return mapa, avisos


def carregar_respostas(cfg):
    linhas, avisos = ler_tabela_respostas(cfg)
    if not linhas:
        return [], avisos
    mapa, av = mapear_colunas(linhas[0], cfg.get("colunas"))
    avisos += av

    def cel(row, campo):
        j = mapa.get(campo)
        if j is None or j >= len(row):
            return None
        return row[j]

    def txt(row, campo):
        v = cel(row, campo)
        return "" if v is None else str(v).strip()

    respostas = []
    for row in linhas[1:]:
        mun = txt(row, "municipio")
        if not mun:
            continue
        resp = {
            "municipio": mun,
            "data_apoio": parse_data(cel(row, "data_apoio")),
            "carimbo": parse_carimbo(cel(row, "carimbo")),
            "setor": txt(row, "setor"),
        }
        for campo in ("esclareceu", "formato", "tempo", "explicar", "geral",
                      "comentario_esclareceu", "comentario_formato", "comentario_explicar", "comentario_geral"):
            resp[campo] = txt(row, campo)
        respostas.append(resp)
    return respostas, avisos


def selecionar_respostas(respostas, municipio, data, datas_cadastradas=(), incluir_dia_seguinte=True,
                         campo="data_apoio"):
    """Respostas do município na data informada e, opcionalmente, no dia seguinte.

    O dia seguinte só entra se não houver outra atividade cadastrada naquele dia para o mesmo município.
    Devolve (exatas, do_dia_seguinte).
    """
    m = norm(municipio)
    seguinte = data + timedelta(days=1)
    exatas, segs = [], []
    for r in respostas:
        if norm(r["municipio"]) != m:
            continue
        d = r["data_apoio"] if campo == "data_apoio" else (r["carimbo"].date() if r["carimbo"] else None)
        if d is None:
            continue
        if d == data:
            exatas.append(r)
        elif incluir_dia_seguinte and d == seguinte and seguinte not in set(datas_cadastradas):
            segs.append(r)
    return exatas, segs
