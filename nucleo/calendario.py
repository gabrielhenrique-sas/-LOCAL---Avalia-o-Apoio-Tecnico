"""
Calendário online (Google Planilhas) -> lista de ATMs de uma semana.

  baixar(fonte)  baixa a planilha inteira para a memória (nada é gravado em disco; a online nunca é alterada)
  extrair(...)   interpreta as abas mensais e devolve uma linha por ATM/município

Layout esperado (igual ao da planilha SEASS 2026):
  - uma aba por mês (JANEIRO ... DEZEMBRO)
  - cada dia ocupa um bloco de colunas: Dom=C, Seg=E, Ter=H, Qua=K, Qui=N, Sex=Q, Sáb=T
  - o número do dia fica na linha de cabeçalho da semana
  - em cada dia há duas faixas:
        "ATMs PRESENCIAIS" ... (linha antes de) "ATMs REMOTOS"  -> Presencial
        "ATMs REMOTOS"     ... (linha antes de) "AÇÕES"         -> Remoto
  - a atividade fica na 2ª coluna do bloco; o nº do processo na 1ª coluna
  - a linha "Informes" (coluna D) marca o fim da área com datas
"""
import calendar
import datetime as dt
import difflib
import io
import re
import time
import urllib.error
import urllib.request
from pathlib import Path

import openpyxl
from openpyxl.utils import get_column_letter

from .util import norm

ARQ_MUNICIPIOS = Path(__file__).resolve().parent.parent / "municipios_pe.txt"   # lista oficial (IBGE)
URL_EXPORT = "https://docs.google.com/spreadsheets/d/{id}/export?format=xlsx"
MODALIDADES = ["Presencial", "Remoto"]

# Gerências: só estes nomes entram (em qualquer ordem/posição no texto da célula).
GERENCIAS_VALIDAS = ["GEBEN", "GEPSB", "GEPMC", "GEFEAS", "VSA", "CAD", "GETEP", "GEPAC"]
APELIDOS = {"GFEAS": "GEFEAS", "GEPBS": "GEPSB"}            # erros de digitação conhecidos
GVINF_EXPANDE_PARA = ["CAD", "VSA"]                          # GVINF vira CAD e VSA (exceto GVINF/CAD, GVINF/VSA)
SEPARADOR_GERENCIAS = ", "
LIMITE_SEMELHANCA = 0.84                                     # corrige grafias: "Garanhus" -> "Garanhuns"
APELIDOS_MUNICIPIOS = {"Jaboatão": "Jaboatão dos Guararapes"}

MESES = ["JANEIRO", "FEVEREIRO", "MARÇO", "ABRIL", "MAIO", "JUNHO",
         "JULHO", "AGOSTO", "SETEMBRO", "OUTUBRO", "NOVEMBRO", "DEZEMBRO"]
COLS_DIA = [2, 4, 7, 10, 13, 16, 19]   # índices 0-based de C, E, H, K, N, Q, T
COL_INFORMES = 3                        # coluna D


# ------------------------------------------------------------------ download
class ErroFonte(Exception):
    """Falha ao obter o calendário (rede, permissão, link errado...). Nada é alterado."""


def url_de_exportacao(fonte):
    """Link/ID do Google Planilhas -> URL de download .xlsx."""
    s = str(fonte).strip()
    if "/export" in s:
        return s
    m = re.search(r"/spreadsheets/d/([A-Za-z0-9_-]+)", s)
    if m:
        return URL_EXPORT.format(id=m.group(1))
    if re.fullmatch(r"[A-Za-z0-9_-]{25,}", s):
        return URL_EXPORT.format(id=s)
    return s


def baixar(fonte, tentativas=2, timeout=40):
    """Devolve os bytes do .xlsx. 'fonte' pode ser link/ID do Google ou caminho de um arquivo local."""
    s = str(fonte or "").strip()
    if not s:
        raise ErroFonte("link do calendário não informado (config.json > calendario > fonte).")
    if not s.lower().startswith(("http://", "https://")) and Path(s).exists():
        return Path(s).read_bytes()
    url = url_de_exportacao(s)
    if not url.lower().startswith(("http://", "https://")):
        raise ErroFonte(f"arquivo não encontrado e não parece um link: {s}")
    ultimo = None
    for i in range(1, tentativas + 1):
        try:
            req = urllib.request.Request(url, headers={"User-Agent": "Mozilla/5.0 (relatorios_atm)"})
            with urllib.request.urlopen(req, timeout=timeout) as r:
                dados = r.read()
            break
        except urllib.error.HTTPError as e:
            if e.code in (401, 403):
                raise ErroFonte(f"sem permissão (HTTP {e.code}). Compartilhe como 'Qualquer pessoa com o link' (leitor).")
            if e.code == 404:
                raise ErroFonte("planilha não encontrada (HTTP 404). Confira o link em config.json.")
            ultimo = e
        except (urllib.error.URLError, TimeoutError, ConnectionError, OSError) as e:
            ultimo = e
        if i < tentativas:
            time.sleep(1.5)
    else:
        raise ErroFonte(f"falha de rede: {ultimo}")
    if not dados.startswith(b"PK"):          # .xlsx é um zip; HTML = página de login/erro
        raise ErroFonte("a resposta não é uma planilha .xlsx (provavelmente página de login do Google). "
                        "Verifique o compartilhamento 'Qualquer pessoa com o link'.")
    return dados


# ------------------------------------------------------------------ utilitários
def num_dia(v):
    s = str(v if v is not None else "").strip()
    if not re.fullmatch(r"\d+(\.0+)?", s):
        return None
    n = int(float(s))
    return n if 1 <= n <= 31 else None


# ------------------------------------------------------------------ municípios
def _carregar_municipios():
    nomes = [l.strip() for l in ARQ_MUNICIPIOS.read_text(encoding="utf-8").splitlines() if l.strip()]
    por_norm = {norm(n): n for n in nomes}
    ordenados = sorted(por_norm.items(), key=lambda kv: -len(kv[0]))   # nomes longos primeiro
    return por_norm, ordenados


MUN_POR_NORM, MUN_ORDENADOS = _carregar_municipios()
_STOP = {"ao", "a", "o", "de", "do", "da", "dos", "das", "no", "na", "em", "municipio", "municipios"}


def resolver_municipios(texto):
    """
    Identifica os municípios de um trecho de texto usando a lista oficial.
    Devolve (lista_de_municipios_oficiais, observacoes, pendencias, ok).
      - procura nomes oficiais inteiros no texto (ex.: 'Cachoeirinha e Agrestina' -> 2)
      - o que sobrar é comparado por semelhança (ex.: 'Garanhus' -> 'Garanhuns')
      - o que continuar sem identificação vai em 'pendencias' (e ok = False)
    """
    t = re.sub(r"\([^)]*\)", " ", str(texto or ""))                       # remove (parênteses)
    t = re.sub(r"\b\d{1,2}\s*(?:h|hs|hrs?)\s*\d{0,2}\b|\b\d{1,2}:\d{2}\b", " ", t, flags=re.I)  # horários
    resto = norm(t)
    achados = []
    for nn, oficial in MUN_ORDENADOS:
        padrao = r"(?<![a-z0-9])" + re.escape(nn) + r"(?![a-z0-9])"
        while True:
            m = re.search(padrao, resto)
            if not m:
                break
            achados.append((m.start(), oficial))
            resto = resto[:m.start()] + " " * (m.end() - m.start()) + resto[m.end():]
    achados.sort()
    muns = []
    for _, o in achados:
        if o not in muns:
            muns.append(o)

    obs, pend = [], []
    for peca in re.split(r",|;|/|\s-\s|\be\b", resto):
        toks = peca.split()
        while toks and toks[0] in _STOP:
            toks.pop(0)
        while toks and toks[-1] in _STOP:
            toks.pop()
        peca = " ".join(toks)
        if len(peca) < 3 or not re.search(r"[a-z]", peca):
            continue
        cand = difflib.get_close_matches(peca, list(MUN_POR_NORM), n=1, cutoff=LIMITE_SEMELHANCA)
        apelido = {norm(k): v for k, v in APELIDOS_MUNICIPIOS.items()}.get(peca)
        if apelido or cand:
            oficial = apelido or MUN_POR_NORM[cand[0]]
            if oficial not in muns:
                muns.append(oficial)
            obs.append(f"grafia corrigida: '{peca}' → '{oficial}'")
        else:
            pend.append(peca)
    return muns, obs, pend, bool(muns) and not pend


# ------------------------------------------------------------------ gerências
_TODOS_TOKENS = list(GERENCIAS_VALIDAS) + list(APELIDOS) + ["GVINF"]
RE_TOKEN = re.compile(r"(?<![A-Za-zÀ-ÿ])(" + "|".join(_TODOS_TOKENS) + r")(?![A-Za-zÀ-ÿ])", re.I)


def extrair_gerencias(texto):
    """
    Pega só as gerências válidas do texto (qualquer ordem/posição), ignorando o resto.
    Aplica apelidos (GFEAS→GEFEAS, GEPBS→GEPSB) e a regra do GVINF:
      GVINF sozinho -> CAD e VSA ;  GVINF/CAD ou GVINF/VSA -> só a escrita.
    """
    texto = str(texto or "")
    out = []
    for m in RE_TOKEN.finditer(texto):
        tok = m.group(1).upper()
        tok = APELIDOS.get(tok, tok)
        if tok == "GVINF":
            if re.match(r"\s*[/(]\s*(?:CAD|VSA)(?![A-Za-zÀ-ÿ])", texto[m.end():], re.I):
                continue                      # GVINF/CAD ou GVINF(VSA): vale só o que foi escrito
            novos = list(GVINF_EXPANDE_PARA)
        else:
            novos = [tok]
        for n in novos:
            if n not in out:
                out.append(n)
    return SEPARADOR_GERENCIAS.join(out)


# ------------------------------------------------------------------ texto da célula
def interpretar(texto):
    """Quebra o texto da célula em (texto_do_município, texto_das_gerências, observações)."""
    linhas = [l.strip() for l in str(texto).replace("\r", "").split("\n") if l.strip()]
    obs = []
    i0 = next((i for i, l in enumerate(linhas) if re.match(r"ATM", l, re.I)), 0)
    obs.extend(linhas[:i0])                 # ex.: "(Será Remarcado)" antes do ATM
    linhas = linhas[i0:]

    mun, resto = "", []
    primeira = linhas[0] if linhas else ""
    m = (re.search(r"munic[ií]pio\s+de\b\s*(.*)$", primeira, re.I)
         or re.search(r"munic[ií]pio\s+(?!de\b)(.+)$", primeira, re.I)
         or re.search(r"ATM\s+em\s+(.+)$", primeira, re.I))
    if m and m.group(1).strip():            # "ATM Remoto ao município de Olinda"
        mun, resto = m.group(1).strip(), linhas[1:]
    elif len(linhas) > 1:                   # "ATM no Município de" / "Lagoa do Carro"
        mun, resto = linhas[1], linhas[2:]

    ger_txt = []
    tk = RE_TOKEN.search(mun)               # gerência escrita na mesma linha do município
    if tk and tk.start() > 0:
        ger_txt.append(mun[tk.start():])
        mun = mun[:tk.start()]
    mun = mun.strip(" -–/,;")

    for l in resto:
        if re.fullmatch(r"CONFIRMA[ÇC][ÃA]O", l, re.I):
            obs.append("CONFIRMAÇÃO")
        elif re.fullmatch(r"\(.*\)", l) or re.search(r"remarcad", l, re.I):
            obs.append(l.strip("()"))
        else:
            ger_txt.append(l)
    return mun, " ".join(ger_txt), "; ".join(obs)


# ------------------------------------------------------------------ extração
def extrair_de_matriz(nome_aba, mes, vals, inicio=None, fim=None, ano=2026):
    """Recebe a matriz de valores de uma aba e devolve (lista de ATMs, avisos)."""
    linhas, avisos = [], []
    n_l = len(vals)

    def cel(r, c):
        return vals[r][c] if r < n_l and c < len(vals[r]) and vals[r][c] is not None else ""

    # 1) cabeçalhos de semana (linhas com o número dos dias)
    cab = {}
    for r in range(n_l):
        dias = {c: num_dia(cel(r, c)) for c in COLS_DIA if num_dia(cel(r, c))}
        if not dias:
            continue
        prox = vals[r + 1] if r + 1 < n_l else []
        if r <= 5 or any(re.match(r"ATMs", str(v or "").strip(), re.I) for v in prox):
            cab[r] = dias
    ordem = sorted(cab)

    # 2) tudo abaixo de "Informes" fica de fora (ATMs a remarcar, sem data)
    limite = next((r for r in range(n_l)
                   if str(cel(r, COL_INFORMES)).strip().lower() == "informes"), n_l)

    capturadas, area = set(), set()
    for i, r0 in enumerate(ordem):
        r1 = min(ordem[i + 1] if i + 1 < len(ordem) else limite, limite)
        for idx_col, c in enumerate(COLS_DIA):
            dia = cab[r0].get(c)
            if not dia:
                continue
            try:
                data = dt.date(ano, mes, dia)
            except ValueError:
                avisos.append(f"{nome_aba}: dia inválido {dia} em {get_column_letter(c + 1)}{r0 + 1}")
                continue
            if (data.weekday() + 1) % 7 != idx_col:
                avisos.append(f"{nome_aba}: {data:%d/%m} não bate com a coluna {get_column_letter(c + 1)}{r0 + 1}")
            if (inicio and data < inicio) or (fim and data > fim):
                continue                                   # fora da janela pedida
            faixa = None
            for r in range(r0 + 1, r1):
                area.add((r, c + 1))
                sec = cel(r, c)
                if isinstance(sec, str):
                    s = sec.strip().upper()
                    if s.startswith("ATMS PRESENCIAIS"):
                        faixa = "Presencial"
                    elif s.startswith("ATMS REMOTOS"):
                        faixa = "Remoto"
                    elif s in ("AÇÕES", "AÇOES"):
                        faixa = None
                v = cel(r, c + 1)
                if not (isinstance(v, str) and re.match(r"\s*ATM", v, re.I)) or faixa is None:
                    continue                               # só ATMs dentro das duas faixas
                mun_txt, ger_txt, obs = interpretar(v)
                ger = extrair_gerencias(ger_txt)
                muns, obs_m, pend, ok = resolver_municipios(mun_txt)
                proc = cel(r, c)
                if isinstance(proc, float) and proc.is_integer():
                    proc = int(proc)
                revisar = (not ok) or (not ger)
                extras = [obs] if obs else []
                extras += obs_m
                if pend:
                    extras.append("trecho não reconhecido: " + " / ".join(pend))
                if revisar:
                    extras.append("texto original: " + re.sub(r"\s+", " ", v).strip()[:90])
                base = {"data": data, "ger": ger, "modalidade": faixa, "proc": str(proc).strip(),
                        "obs": "; ".join(extras), "mes": nome_aba,
                        "celula": f"{get_column_letter(c + 2)}{r + 1}",
                        "revisar": "SIM" if revisar else ""}
                for mun in (muns or [mun_txt]):
                    linhas.append(dict(base, mun=mun))
                capturadas.add((r, c + 1))

    # 3) varredura de segurança: "ATM" na área lida que não foi capturado
    #    (ex.: ATM digitado fora das faixas). Sem janela, olha a aba inteira.
    celulas = ((r, c) for r in range(min(limite, n_l)) for c in range(len(vals[r]))) \
        if (inicio is None and fim is None) else sorted(area)
    for r, c in celulas:
        v = cel(r, c)
        if isinstance(v, str) and re.match(r"\s*ATM(?!s\s)", v, re.I) and (r, c) not in capturadas:
            avisos.append(f"{nome_aba}!{get_column_letter(c + 1)}{r + 1} tem 'ATM' mas não foi capturada "
                          f"(fora das faixas?): {v[:50]!r}")
    return linhas, avisos


def extrair(origem, inicio=None, fim=None, ano=2026):
    """Lê o calendário (bytes ou caminho) e devolve (ATMs ordenados, avisos). Com inicio/fim, só essa janela."""
    if isinstance(origem, (bytes, bytearray)):
        origem = io.BytesIO(origem)
    wb = openpyxl.load_workbook(origem, data_only=True, read_only=True)
    linhas, avisos = [], []
    for i, nome in enumerate(MESES, start=1):
        primeiro = dt.date(ano, i, 1)
        ultimo = dt.date(ano, i, calendar.monthrange(ano, i)[1])
        if (inicio and ultimo < inicio) or (fim and primeiro > fim):
            continue                                       # mês fora da janela: nem lê
        if nome not in wb.sheetnames:
            if inicio or fim:
                avisos.append(f"aba {nome} não encontrada na planilha (o nome da aba mudou?)")
            continue
        ws = wb[nome]
        if hasattr(ws, "reset_dimensions"):                # evita dimensão errada em arquivos exportados
            ws.reset_dimensions()
        vals = [list(r) for r in ws.iter_rows(values_only=True)]
        l, a = extrair_de_matriz(nome, i, vals, inicio, fim, ano)
        linhas += l
        avisos += a
    cont = {}
    for l in linhas:                       # chave única: data | processo | município
        base = f"{l['data'].isoformat()}|{l['proc']}|{norm(l['mun'])}"
        cont[base] = cont.get(base, 0) + 1
        l["chave"] = base if cont[base] == 1 else f"{base}#{cont[base]}"
    linhas.sort(key=lambda x: (x["data"], x["modalidade"] != "Presencial", x["mun"]))
    wb.close()
    return linhas, avisos
