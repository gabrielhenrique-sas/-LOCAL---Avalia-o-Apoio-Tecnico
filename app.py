"""Relatórios de Avaliação do Apoio Técnico: página web local (calendário -> semana -> relatório -> e-mail).

Uso:  python app.py   (abre http://127.0.0.1:5000)
"""
import secrets
import threading
import webbrowser
from datetime import date, timedelta

from flask import Flask, Response, abort, redirect, render_template, request, send_file, url_for

from nucleo import calendario, controle, email_util, fluxo
from nucleo import config as cfgmod
from nucleo import respostas as resp
from nucleo.respostas import ErroDados
from nucleo.util import DIAS, fmt_data, fmt_hora, parse_data, segunda_de, slug

app = Flask(__name__)
app.jinja_env.filters.update(data=fmt_data, hora=fmt_hora)
app.jinja_env.globals.update(DIAS=DIAS)
PREPARADOS = {}  # token -> envio em preparação (uso de um só usuário, na própria máquina)
MUNICIPIOS = sorted(calendario.MUN_POR_NORM.values(), key=lambda m: m.lower())


def _cfg():
    return cfgmod.carregar()


def _semana():
    ref = parse_data(request.values.get("semana")) or date.today()
    return segunda_de(ref)


def _erro(msg, codigo=400):
    return render_template("erro.html", mensagem=msg, semana=_semana()), codigo


# ---------------------------------------------------------------- semana
@app.get("/")
def index():
    cfg, ini = _cfg(), _semana()
    try:
        ctx = fluxo.carregar_semana(cfg, ini, aceitar=bool(request.args.get("aceitar")))
    except ErroDados as e:
        return _erro(str(e), 500)
    return render_template("semana.html", ctx=ctx, n_revisar=ctx["n_revisar"], semana=ini,
                           anterior=ini - timedelta(days=7), proxima=ini + timedelta(days=7),
                           atual=segunda_de(date.today()), cfg=cfg, smtp_ok=email_util.smtp_configurado(cfg))


@app.post("/preparar")
def preparar():
    cfg = _cfg()
    marcados = [request.form["unico"]] if request.form.get("unico") else request.form.getlist("sel")
    if not marcados:
        return _erro("Marque ao menos uma atividade.")
    sel = []
    for m in marcados:
        d, _, mun = m.partition("|")
        sel.append((mun, parse_data(d)))
    if len({mun.lower() for mun, _ in sel}) > 1:
        return _erro("Selecione atividades de um único município por vez (o relatório é por município).")
    try:
        respostas, _ = resp.carregar_respostas(cfg)
        prep = fluxo.preparar_envio(cfg, controle.ler(cfg), sel, respostas)
    except ErroDados as e:
        return _erro(str(e))
    prep["semana"] = _semana().isoformat()
    token = secrets.token_urlsafe(8)
    PREPARADOS[token] = prep
    return redirect(url_for("envio", token=token))


# ---------------------------------------------------------------- revisão
@app.get("/revisao")
def revisao():
    cfg = _cfg()
    try:
        ctrl = controle.ler(cfg)
    except ErroDados as e:
        return _erro(str(e), 500)
    itens = sorted((a for a in fluxo.ativas(ctrl) if a["revisar"]), key=lambda a: (a["data"], a["mun"]))
    return render_template("revisao.html", itens=itens, n_revisar=len(itens), semana=_semana(),
                           municipios=MUNICIPIOS, gerencias=calendario.GERENCIAS_VALIDAS,
                           modalidades=calendario.MODALIDADES, avisos=fluxo.ULTIMO["avisos"],
                           semana_avisos=fluxo.ULTIMO["ini"], ok=request.args.get("ok"), erro=request.args.get("erro"))


@app.post("/revisao/confirmar")
def revisao_confirmar():
    f = request.form
    try:
        fluxo.confirmar_revisao(_cfg(), f.get("chave", ""), parse_data(f.get("data")), f.get("municipio", ""),
                                f.get("modalidade", ""), f.getlist("gerencias"))
    except ErroDados as e:
        return redirect(url_for("revisao", semana=f.get("semana"), erro=str(e)))
    return redirect(url_for("revisao", semana=f.get("semana"), ok="Atividade confirmada: já aparece na lista da semana."))


# ---------------------------------------------------------------- envio
def _prep(token):
    prep = PREPARADOS.get(token)
    if prep is None:
        abort(404)
    return prep


def _render_envio(token, prep, aviso=None, erro=None, sucesso=None):
    cfg = _cfg()
    return render_template("envio.html", token=token, p=prep, cfg=cfg, smtp_ok=email_util.smtp_configurado(cfg),
                           aviso=aviso, erro=erro, sucesso=sucesso, semana=parse_data(prep["semana"]))


@app.get("/envio/<token>")
def envio(token):
    return _render_envio(token, _prep(token))


@app.get("/pdf/<token>")
def pdf(token):
    prep = _prep(token)
    if not prep["pdf"]:
        abort(404)
    return send_file(prep["pdf"], mimetype="application/pdf", download_name=prep["pdf"].replace("\\", "/").split("/")[-1])


@app.post("/envio/<token>/acao")
def acao(token):
    cfg, prep, f = _cfg(), _prep(token), request.form
    for campo in ("para", "cc", "cco", "assunto", "corpo"):
        prep[campo] = f.get(campo, "")
    acao_ = f.get("acao")

    if acao_ == "regerar":
        try:
            respostas, _ = resp.carregar_respostas(cfg)
            novo = fluxo.preparar_envio(cfg, controle.ler(cfg), [(prep["municipio"], d) for d in prep["datas"]],
                                        respostas, incluir_dia_seguinte=bool(f.get("dia_seguinte")))
        except ErroDados as e:
            return _render_envio(token, prep, erro=str(e))
        for k in ("incluir_dia_seguinte", "contagens", "total", "pdf", "avisos"):
            prep[k] = novo[k]
        return _render_envio(token, prep, aviso="PDF atualizado.")

    para, inv1 = email_util.separar_emails(prep["para"])
    cc, inv2 = email_util.separar_emails(prep["cc"])
    cco, inv3 = email_util.separar_emails(prep["cco"])
    invalidos = inv1 + inv2 + inv3

    if acao_ == "eml":
        msg = email_util.montar_mensagem(prep["remetente"], prep["nome_remetente"], para, cc, cco,
                                         prep["assunto"], prep["corpo"], prep["pdf"], rascunho=True)
        return Response(bytes(msg), mimetype="message/rfc822",
                        headers={"Content-Disposition": f'attachment; filename="rascunho_{slug(prep["municipio"])}.eml"'})

    if acao_ == "enviar":
        if prep.get("enviado"):
            return _render_envio(token, prep, erro="Este envio já foi feito. Para enviar de novo, prepare-o outra vez na semana.")
        if invalidos:
            return _render_envio(token, prep, erro="E-mails inválidos: " + ", ".join(invalidos))
        if not (para or cco):
            return _render_envio(token, prep, erro="Informe ao menos um destinatário.")
        msg = email_util.montar_mensagem(prep["remetente"], prep["nome_remetente"], para, cc, cco,
                                         prep["assunto"], prep["corpo"], prep["pdf"])
        try:
            email_util.enviar_smtp(cfg, msg, para + cc + cco)
        except Exception as e:  # rede, login etc.
            return _render_envio(token, prep, erro=f"Falha no envio: {e}")
        prep["enviado"] = True
        aviso = None
        try:
            fluxo.registrar_envio(cfg, prep, len(para + cc + cco))
        except ErroDados as e:
            aviso = f"E-mail enviado, mas não foi possível atualizar o CONTROLE.xlsx: {e}"
        return _render_envio(token, prep, aviso=aviso, sucesso="E-mail enviado.")
    abort(400)


def main():
    threading.Timer(1.0, lambda: webbrowser.open("http://127.0.0.1:5000")).start()
    app.run(host="127.0.0.1", port=5000, debug=False)


if __name__ == "__main__":
    main()
