"""Fluxo da semana: carrega calendário + respostas, agrupa atividades, prepara e registra os envios."""
import re
from datetime import datetime, timedelta

from . import calendario, controle, email_util, relatorio, respostas as resp
from .calendario import GERENCIAS_VALIDAS, MODALIDADES
from .config import RAIZ, pasta_saida
from .respostas import ErroDados
from .sync import mesclar
from .util import fmt_data, norm, slug, texto_datas

PAPEL = RAIZ / "assets" / "papel_timbrado.png"
ASSUNTO_RELATORIO = "Relatório de Avaliações do Apoio Técnico - {{MUNICIPIO}}"
CORPO_RELATORIO = ("Prezados(as),\n\nSegue em anexo o relatório de avaliações do Apoio Técnico realizado em "
                   "{{MUNICIPIO}} ({{DATA}}).\n\nAtenciosamente.")
ASSUNTO_SEM = "Apoio Técnico - {{MUNICIPIO}} - sem avaliações recebidas"
CORPO_SEM = ("Prezados(as),\n\nInformamos que não foram recebidas avaliações do Apoio Técnico realizado em "
             "{{MUNICIPIO}} ({{DATA}}).\n\nAtenciosamente.")
ENVIADOS = ("Enviado", "Sem respostas")
ULTIMO = {"ini": None, "avisos": []}      # avisos da última semana aberta (a página Revisão mostra)


def _lista(texto):
    return [g.strip() for g in re.split(r"[;,]+", texto or "") if g.strip()]


def _unicos(itens):
    saida = []
    for i in itens:
        if i not in saida:
            saida.append(i)
    return saida


def ativas(ctrl):
    return [a for a in ctrl["atividades"] if a["situacao"] == "Ativo"]


# ---------------------------------------------------------------- agrupamento
def agrupar(atividades):
    """Uma linha por município+data (mesmo município no mesmo dia = uma só atividade)."""
    grupos = {}
    for a in atividades:
        g = grupos.setdefault((norm(a["mun"]), a["data"]), {
            "municipio": a["mun"], "data": a["data"], "gerencias": [], "modalidades": [], "linhas": []})
        g["gerencias"] = _unicos(g["gerencias"] + _lista(a["ger"]))
        g["modalidades"] = _unicos(g["modalidades"] + [a["modalidade"]])
        g["linhas"].append(a)
    for g in grupos.values():
        env = [a for a in g["linhas"] if a["status"] in ENVIADOS]
        g["status"] = env[0]["status"] if env else ""
        g["enviado_em"] = max((a["enviado_em"] for a in env if a["enviado_em"]), default=None)
        g["chave"] = f"{g['data'].isoformat()}|{g['municipio']}"
    return sorted(grupos.values(), key=lambda g: (g["data"], norm(g["municipio"])))


def _datas_por_municipio(ctrl):
    d = {}
    for a in ativas(ctrl):
        d.setdefault(norm(a["mun"]), set()).add(a["data"])
    return d


# ---------------------------------------------------------------- semana
def carregar_semana(cfg, ini, aceitar=False):
    """Atualiza calendário e respostas (online) e devolve tudo que a tela da semana precisa."""
    fim = ini + timedelta(days=4)
    agora = datetime.now().replace(microsecond=0)
    ctx = {"ini": ini, "fim": fim, "erros": [], "avisos": [], "mudancas": [], "abortado": False, "respostas": None}
    ctrl = controle.ler(cfg)

    try:                                                     # 1) calendário -> atividades
        linhas, av = calendario.extrair(calendario.baixar(cfg["calendario"]["fonte"]),
                                        ini, ini + timedelta(days=6), ini.year)
        ctx["avisos"] += av
        for l in (x for x in linhas if x["data"].weekday() >= 5):
            ctx["avisos"].append(f"ATM em fim de semana ignorado: {l['mun']} em {fmt_data(l['data'])}.")
        linhas = [l for l in linhas if l["data"].weekday() < 5]
        ev = mesclar(ctrl["atividades"], linhas, ini, fim, agora, aceitar)
        if ev and ev[0][0] == "ABORTADO":
            ctx["abortado"] = True
            ctx["erros"].append(ev[0][3])
        elif ev:
            ctx["mudancas"] = ev
            try:
                controle.salvar(cfg, ctrl["atividades"], ev)
            except ErroDados as e:
                ctx["erros"].append(str(e))
    except calendario.ErroFonte as e:
        ctx["erros"].append(f"Calendário online indisponível ({e}). Mostrando os últimos dados salvos.")

    try:                                                     # 2) respostas do formulário
        ctx["respostas"], av = resp.carregar_respostas(cfg)
        ctx["avisos"] += av
    except Exception as e:
        ctx["erros"].append(f"Respostas do formulário indisponíveis: {e}")

    ativ = ativas(ctrl)
    ctx["n_revisar"] = sum(1 for a in ativ if a["revisar"])
    ctx["n_revisar_semana"] = sum(1 for a in ativ if a["revisar"] and ini <= a["data"] <= fim)
    ctx["grupos"] = agrupar([a for a in ativ if not a["revisar"] and ini <= a["data"] <= fim])
    datas_cad = _datas_por_municipio(ctrl)
    for g in ctx["grupos"]:
        g["exatas"] = g["seguinte"] = None
        if ctx["respostas"] is not None:
            ex, seg = resp.selecionar_respostas(ctx["respostas"], g["municipio"], g["data"],
                                                datas_cad[norm(g["municipio"])], cfg["incluir_dia_seguinte"],
                                                cfg["campo_data_busca"])
            g["exatas"], g["seguinte"] = len(ex), len(seg)
    ctx["itens_revisao"] = avisos_da_semana(cfg, ctrl, ctx)
    ULTIMO.update(ini=ini, avisos=ctx["itens_revisao"])
    return ctx


def avisos_da_semana(cfg, ctrl, ctx):
    """Itens só de leitura para a aba Revisão (avisos do calendário, respostas sem atividade, atrasos, e-mails)."""
    ini, fim, hoje = ctx["ini"], ctx["fim"], datetime.now().date()
    itens = [{"tipo": "Calendário", "municipio": "", "data": None, "detalhe": a} for a in ctx["avisos"]]
    inc, campo = cfg["incluir_dia_seguinte"], cfg["campo_data_busca"]
    cad = {}
    for a in ativas(ctrl):
        cad.setdefault(norm(a["mun"]), {})[a["data"]] = a
    sem, vistos_atraso = {}, set()
    for r in ctx["respostas"] or []:
        d = r["data_apoio"] if campo == "data_apoio" else (r["carimbo"].date() if r["carimbo"] else None)
        if d is None or not (ini <= d <= fim + timedelta(days=2)):
            continue
        dias = cad.get(norm(r["municipio"]), {})
        alvo = dias.get(d) or (dias.get(d - timedelta(days=1)) if inc and d not in dias else None)
        if alvo is None:
            if d < hoje:
                sem[(r["municipio"], d)] = sem.get((r["municipio"], d), 0) + 1
        elif alvo["status"] in ENVIADOS and alvo["enviado_em"] and r["carimbo"] and r["carimbo"] > alvo["enviado_em"]:
            if (alvo["chave"], r["carimbo"]) not in vistos_atraso:
                vistos_atraso.add((alvo["chave"], r["carimbo"]))
                itens.append({"tipo": "Resposta com atraso", "municipio": alvo["mun"], "data": alvo["data"],
                              "detalhe": f"Avaliação recebida em {r['carimbo']:%d/%m/%Y %H:%M}, depois do envio "
                                         f"({alvo['enviado_em']:%d/%m/%Y %H:%M})."})
    for (mun, d), n in sem.items():
        itens.append({"tipo": "Resposta sem atividade", "municipio": mun, "data": d,
                      "detalhe": f"{n} resposta(s) de {mun} em {fmt_data(d)} sem atividade correspondente no calendário."})
    for g in ctx["grupos"]:
        if g["status"]:
            continue
        for aviso in emails_das_gerencias(ctrl, g["gerencias"])[1]:
            itens.append({"tipo": "Gerência sem e-mail", "municipio": g["municipio"], "data": g["data"], "detalhe": aviso})
    return itens


# ---------------------------------------------------------------- revisão manual
def confirmar_revisao(cfg, chave, data, municipio, modalidade, gerencias):
    """Aplica a correção manual e libera a atividade para a lista normal. Levanta ErroDados se algo for inválido."""
    oficial = calendario.MUN_POR_NORM.get(norm(municipio))
    if not oficial:
        raise ErroDados(f"'{municipio}' não está na lista oficial de municípios de Pernambuco.")
    if data is None or data.weekday() >= 5:
        raise ErroDados("Informe uma data de segunda a sexta.")
    if modalidade not in MODALIDADES:
        raise ErroDados("Escolha a modalidade (Presencial ou Remoto).")
    ger = [g for g in GERENCIAS_VALIDAS if g in gerencias]
    if not ger:
        raise ErroDados("Marque ao menos uma gerência.")
    ctrl = controle.ler(cfg)
    a = next((x for x in ctrl["atividades"] if x["chave"] == chave), None)
    if a is None:
        raise ErroDados("Atividade não encontrada (a planilha mudou?). Recarregue a página.")
    a.update(data=data, mun=oficial, modalidade=modalidade, ger=", ".join(ger), revisar=False, manual=True,
             atualizado=datetime.now().replace(microsecond=0))
    controle.salvar(cfg, ctrl["atividades"], [("CONFIRMADO", oficial, data, f"{', '.join(ger)} ({modalidade})")])


# ---------------------------------------------------------------- e-mail e PDF
def emails_das_gerencias(ctrl, nomes):
    """Devolve (emails, avisos). Gerência desconhecida, inativa ou sem e-mail válido gera aviso."""
    emails, avisos = [], []
    for nome in nomes:
        g = ctrl["gerencias"].get(nome)
        if g is None:
            avisos.append(f"Gerência '{nome}' não está cadastrada na aba Gerências.")
        elif not g["ativa"]:
            avisos.append(f"Gerência '{nome}' está marcada como inativa.")
        else:
            validos, _ = email_util.separar_emails(g["email"])
            if not validos:
                avisos.append(f"Gerência '{nome}' está sem e-mail válido.")
            emails += [e for e in validos if e.lower() not in [x.lower() for x in emails]]
    return emails, avisos


def _preencher(modelo, municipio, datas):
    return modelo.replace("{{MUNICIPIO}}", municipio).replace("{{DATA}}", texto_datas(datas))


def preparar_envio(cfg, ctrl, selecionados, respostas, incluir_dia_seguinte=None):
    """selecionados: [(município, data)] de UM município. Seleciona respostas e gera o PDF. Não envia nada."""
    if incluir_dia_seguinte is None:
        incluir_dia_seguinte = cfg["incluir_dia_seguinte"]
    campo = cfg["campo_data_busca"]
    municipio = selecionados[0][0]
    datas = sorted({d for _, d in selecionados})
    grupos = agrupar([a for a in ativas(ctrl) if not a["revisar"] and norm(a["mun"]) == norm(municipio)
                      and a["data"] in datas])
    if not grupos:
        raise ErroDados("Atividade não encontrada. Recarregue a página da semana.")
    municipio = grupos[0]["municipio"]
    gerencias = _unicos([x for g in grupos for x in g["gerencias"]])
    datas_cad = _datas_por_municipio(ctrl)[norm(municipio)]

    secoes, contagens = [], []
    for g in grupos:
        ex, seg = resp.selecionar_respostas(respostas, municipio, g["data"], datas_cad, incluir_dia_seguinte, campo)
        contagens.append({"data": g["data"], "exatas": len(ex), "seguinte": len(seg)})
        secoes.append({"data": g["data"], "respostas": ex + seg, "modalidades": g["modalidades"]})
    total = sum(len(s["respostas"]) for s in secoes)
    datas = [s["data"] for s in secoes]

    emails, avisos = emails_das_gerencias(ctrl, gerencias)
    pdf = None
    if total:
        nome_pdf = f"Relatorio_{slug(municipio)}_{'_'.join(d.isoformat() for d in datas)}.pdf"
        pdf = relatorio.gerar_pdf(pasta_saida(cfg, "pdfs") / nome_pdf, municipio, secoes, gerencias,
                                  pasta_saida(cfg, "tmp_graficos"), PAPEL)

    c = ctrl["config"]
    modelos = (c.get("ASSUNTO_RELATORIO") or ASSUNTO_RELATORIO, c.get("CORPO_RELATORIO") or CORPO_RELATORIO) if total \
        else (c.get("ASSUNTO_SEM_RESPOSTAS") or ASSUNTO_SEM, c.get("CORPO_SEM_RESPOSTAS") or CORPO_SEM)
    email = cfg["email"]
    return {
        "municipio": municipio, "datas": datas, "gerencias": gerencias,
        "modalidades": _unicos([m for s in secoes for m in s["modalidades"]]), "incluir_dia_seguinte": incluir_dia_seguinte,
        "contagens": contagens, "total": total, "pdf": pdf, "avisos": avisos,
        "remetente": email.get("remetente") or "", "nome_remetente": email.get("nome_remetente") or c.get("NOME_REMETENTE") or "",
        "para": email.get("para_padrao") or c.get("EMAIL_PARA") or email.get("remetente") or "", "cc": "",
        "cco": ", ".join(emails),
        "assunto": _preencher(modelos[0], municipio, datas), "corpo": _preencher(modelos[1], municipio, datas),
    }


def registrar_envio(cfg, prep, n_destinatarios):
    """Marca as atividades como enviadas (status, nº de respostas, data/hora, PDF) e grava no Log."""
    ctrl = controle.ler(cfg)
    agora = datetime.now().replace(microsecond=0)
    status = "Enviado" if prep["total"] else "Sem respostas"
    n_por_data = {c["data"]: c["exatas"] + c["seguinte"] for c in prep["contagens"]}
    for a in ativas(ctrl):
        if norm(a["mun"]) == norm(prep["municipio"]) and a["data"] in n_por_data and not a["revisar"]:
            a.update(status=status, n=n_por_data[a["data"]], enviado_em=agora, link=prep["pdf"] or "")
    controle.salvar(cfg, ctrl["atividades"], logs=[{
        "execucao": agora, "municipio": prep["municipio"], "datas": texto_datas(prep["datas"]), "n": prep["total"],
        "destinatarios": n_destinatarios, "status": status, "link": prep["pdf"] or ""}])
