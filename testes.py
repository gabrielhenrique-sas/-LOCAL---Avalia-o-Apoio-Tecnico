"""
Testes automáticos (usam pasta temporária e os dados de exemplo; não mexem nos seus arquivos).

Uso:  python testes.py
"""
import datetime as dt
import shutil
import tempfile
from pathlib import Path

import openpyxl

import app as appmod
from nucleo import calendario, config, controle, fluxo
from nucleo.calendario import extrair_gerencias, resolver_municipios, url_de_exportacao
from nucleo.sync import mesclar
from nucleo.util import segunda_de

RAIZ = Path(__file__).resolve().parent
ok = True


def check(nome, cond, extra=""):
    global ok
    ok &= bool(cond)
    print(("PASSOU  " if cond else "FALHOU  ") + nome, extra if not cond else "")


# ---------------------------------------------------------------- 1) regras de leitura do calendário
G = {"GEPSB/SCFV E PI": "GEPSB", "GEPMC, GEPSB e GEFEAS": "GEPMC, GEPSB, GEFEAS", "GFEAS 09h00": "GEFEAS",
     "GEPBS (09h30)": "GEPSB", "GVINF": "CAD, VSA", "GVINF/CAD": "CAD", "GEPSB / GVINF(VSA)": "GEPSB, VSA",
     "GEPSB, GEPMC e GVINF": "GEPSB, GEPMC, CAD, VSA", "PI SUAS CF": ""}
for txt, esp in G.items():
    check(f"gerência {txt!r} -> {esp!r}", extrair_gerencias(txt) == esp, extrair_gerencias(txt))
M = {"Lagoa do Carro": (["Lagoa do Carro"], True), "Garanhus": (["Garanhuns"], True),
     "Jaboatão": (["Jaboatão dos Guararapes"], True), "Cachoeirinha e Agrestina": (["Cachoeirinha", "Agrestina"], True),
     "Capitais e Metropolitana 10h às 12h": ([], False)}
for txt, (esp, esp_ok) in M.items():
    muns, _, _, o = resolver_municipios(txt)
    check(f"município {txt!r} -> {esp}", muns == esp and o == esp_ok, (muns, o))
check("link do Google vira URL de exportação xlsx", url_de_exportacao(
    "https://docs.google.com/spreadsheets/d/1U1yoaMT9x51nDpjt-i1mJFOacDv7bibond7A6hB3E3E/edit?usp=sharing"
).endswith("/1U1yoaMT9x51nDpjt-i1mJFOacDv7bibond7A6hB3E3E/export?format=xlsx"))
check("semana é segunda a sexta (qua 28/10 -> seg 26/10)", segunda_de(dt.date(2026, 10, 28)) == dt.date(2026, 10, 26))
check("domingo pertence à semana que termina nele", segunda_de(dt.date(2026, 11, 1)) == dt.date(2026, 10, 26))

# ---------------------------------------------------------------- 2) extração com o calendário de exemplo
cal = (RAIZ / "dados" / "calendario_exemplo.xlsx").read_bytes()
ini, fim = dt.date(2026, 10, 26), dt.date(2026, 10, 30)
linhas, _ = calendario.extrair(cal, ini, fim + dt.timedelta(days=2), 2026)
check("extrai ATMs da semana 26-30/10 (bytes, sem arquivo em disco)", len(linhas) >= 5, len(linhas))
check("cada ATM traz 'modalidade' Presencial/Remoto", all(l["modalidade"] in ("Presencial", "Remoto") for l in linhas))

# ---------------------------------------------------------------- 3) sincronização (função pura)
agora = dt.datetime(2026, 10, 26, 8, 0)
ativs = []
ev = mesclar(ativs, linhas, ini, fim, agora)
check("1ª leitura: tudo NOVO", {e[0] for e in ev} == {"NOVO"} and len(ativs) == len(linhas))
check("2ª leitura sem mudanças: nenhum evento", mesclar(ativs, linhas, ini, fim, agora) == [])
alt = [dict(l) for l in linhas]
alt[0]["ger"] = "GEBEN"
ev = mesclar(ativs, alt, ini, fim, agora)
check("mudança no calendário vira ALTERADO (antes → depois)", [e[0] for e in ev] == ["ALTERADO"] and "→" in ev[0][3], ev)
ev = mesclar(ativs, alt[1:], ini, fim, agora)
check("ATM que sumiu fica 'Removido do calendário'", [e[0] for e in ev] == ["REMOVIDO"]
      and sum(a["situacao"] != "Ativo" for a in ativs) == 1, ev)
ev = mesclar(ativs, alt, ini, fim, agora)
check("ATM que voltou é REATIVADO", [e[0] for e in ev] == ["REATIVADO"])
ev = mesclar(ativs, [], ini, fim, agora)
check("leitura vazia aborta sem alterar nada", ev[0][0] == "ABORTADO" and all(a["situacao"] == "Ativo" for a in ativs))
fora = dict(alt[0], data=dt.date(2026, 11, 3), chave="2026-11-03|x|y")
ativs2 = []
mesclar(ativs2, [fora], dt.date(2026, 11, 2), dt.date(2026, 11, 6), agora)
mesclar(ativs2, [dict(l, data=l["data"]) for l in linhas], ini, fim, agora)
check("semana fora da janela não é tocada", ativs2[0]["situacao"] == "Ativo" and len(ativs2) == len(linhas) + 1)

# revisão manual preservada
rev = dict(linhas[0], mun="Cidade Inexistente", ger="", revisar="SIM", chave="2026-10-27|9|inexistente")
a_rev = []
mesclar(a_rev, [rev], ini, fim, agora)
check("ATM sem município/gerência entra com Revisar", a_rev[0]["revisar"])
a_rev[0].update(mun="Olinda", ger="CAD", revisar=False, manual=True)
mesclar(a_rev, [rev], ini, fim, agora)
check("ajuste manual vale enquanto o calendário não muda", a_rev[0]["mun"] == "Olinda" and not a_rev[0]["revisar"])
mesclar(a_rev, [dict(rev, obs="texto novo")], ini, fim, agora)
check("se a célula do calendário muda, volta para revisão", a_rev[0]["revisar"] and not a_rev[0]["manual"])

# ---------------------------------------------------------------- 4) fluxo completo (arquivos de exemplo)
tmp = Path(tempfile.mkdtemp())
shutil.copy(RAIZ / "dados" / "calendario_exemplo.xlsx", tmp / "cal.xlsx")
cfg = config.carregar()
cfg.update(_raiz=str(RAIZ), pasta_saida=str(tmp / "saida"))
cfg["calendario"]["fonte"] = str(tmp / "cal.xlsx")
# respostas de exemplo + avaliações de Bonito (20/10, atividade) e uma do dia seguinte (21/10, sem atividade)
wb = openpyxl.load_workbook(RAIZ / "dados" / "respostas_exemplo.xlsx")
ws = wb.worksheets[0]
modelo = [c.value for c in ws[2]]
for dia, n in (("2026-10-20", 4), ("2026-10-21", 1)):
    for i in range(n):
        linha = list(modelo)
        linha[0], linha[1], linha[2] = f"{dia} 1{i}:00:00", "Bonito", f"{dia} 00:00:00"
        ws.append(linha)
wb.save(tmp / "respostas.xlsx")
cfg["respostas"].update(google_sheet_id="", aba="", arquivo=str(tmp / "respostas.xlsx"))
cfg["controle"].update(arquivo=str(tmp / "CONTROLE.xlsx"))
cfg["email"]["smtp"]["servidor"] = ""          # sem SMTP: nada é enviado de verdade

ctx = fluxo.carregar_semana(cfg, dt.date(2026, 10, 19))
check("CONTROLE.xlsx criado na 1ª abertura, com Gerências copiadas do modelo",
      (tmp / "CONTROLE.xlsx").exists() and "GEBEN" in controle.ler(cfg)["gerencias"]
      and controle.ler(cfg)["gerencias"]["GEBEN"]["email"].endswith("@sas.pe.gov.br"))
nomes = openpyxl.load_workbook(tmp / "CONTROLE.xlsx").sheetnames
check("abas do CONTROLE: Atividades, Gerências, Configurações, Log, Mudanças",
      nomes == ["Atividades", "Gerências", "Configurações", "Log", "Mudanças do calendário"], nomes)
check("semana 19-23/10: atividades segunda a sexta com contagem de avaliações",
      ctx["grupos"] and all(g["data"].weekday() < 5 and g["exatas"] is not None for g in ctx["grupos"]), ctx["erros"])
bonito = next((g for g in ctx["grupos"] if g["municipio"] == "Bonito"), None)
check("Bonito 20/10: modalidade e contagem separada (4 do dia + 1 do dia seguinte)",
      bonito and bonito["modalidades"] == ["Presencial"] and (bonito["exatas"], bonito["seguinte"]) == (4, 1),
      bonito and (bonito["modalidades"], bonito["exatas"], bonito["seguinte"]))
ctx2 = fluxo.carregar_semana(cfg, dt.date(2026, 10, 19))
check("reabrir a semana não gera mudanças repetidas", ctx2["mudancas"] == [])

# revisão: força uma atividade a revisar e confirma
c = controle.ler(cfg)
alvo = next(a for a in c["atividades"] if a["mun"] == "Bonito")
alvo.update(revisar=True, ger="", manual=True)      # estado de uma atividade que aguarda revisão
controle.salvar(cfg, c["atividades"])
ctx = fluxo.carregar_semana(cfg, dt.date(2026, 10, 19))
check("atividade em revisão sai da lista da semana", ctx["n_revisar"] >= 1 and all(g["municipio"] != "Bonito" for g in ctx["grupos"]))
try:
    fluxo.confirmar_revisao(cfg, alvo["chave"], alvo["data"], "Xyz", "Presencial", ["CAD"])
    check("município fora da lista oficial é recusado", False)
except Exception as e:
    check("município fora da lista oficial é recusado", "lista oficial" in str(e))
fluxo.confirmar_revisao(cfg, alvo["chave"], alvo["data"], "bonito", "Presencial", ["GEBEN", "CAD"])
ctx = fluxo.carregar_semana(cfg, dt.date(2026, 10, 19))
b = next((g for g in ctx["grupos"] if g["municipio"] == "Bonito"), None)
check("após confirmar, vai para a lista normal com a correção", b and b["gerencias"] == ["GEBEN", "CAD"])

# ---------------------------------------------------------------- 5) telas e envio
cli = appmod.app.test_client()
appmod.cfgmod.carregar = lambda *a, **k: cfg
r = cli.get("/?semana=2026-10-21")
html = r.get_data(as_text=True)
check("tela da semana mostra o intervalo destacado", r.status_code == 200 and "Semana de 19/10/2026 a 23/10/2026" in html)
check("tela mostra Modalidade e o botão Atualizar", "Modalidade" in html and "Atualizar" in html and "Presencial" in html)
check("tela de revisão abre", cli.get("/revisao").status_code == 200)

r = cli.post("/preparar", data={"unico": "2026-10-20|Bonito", "semana": "2026-10-19"})
check("preparar envio redireciona para a tela do e-mail", r.status_code == 302, r.status_code)
envio_url = r.headers["Location"]
r = cli.get(envio_url)
h = r.get_data(as_text=True)
check("tela do e-mail traz modalidade, gerências e PDF", "Modalidade" in h and "Presencial" in h and "<iframe" in h)
token = envio_url.rsplit("/", 1)[1]
prep = appmod.PREPARADOS[token]
pdf_bytes = Path(prep["pdf"]).read_bytes()
check("PDF gerado", pdf_bytes.startswith(b"%PDF") and len(pdf_bytes) > 5000)
try:
    import fitz  # opcional
    texto = "".join(p.get_text() for p in fitz.open(prep["pdf"]))
    check("PDF contém 'Modalidade: Presencial'", "Modalidade:" in texto and "Presencial" in texto)
except ImportError:
    pass
r = cli.post(f"/envio/{token}/acao", data={"acao": "eml", "para": "a@b.com", "cc": "", "cco": prep["cco"],
                                           "assunto": prep["assunto"], "corpo": prep["corpo"]})
check("rascunho .eml com anexo", r.status_code == 200 and b"application/pdf" in r.data)

# envio simulado (SMTP trocado por um falso)
enviados = []
appmod.email_util.enviar_smtp = lambda cfg_, msg, dest: enviados.append(list(dest))
appmod.email_util.smtp_configurado = lambda cfg_: True
r = cli.post(f"/envio/{token}/acao", data={"acao": "enviar", "para": "a@b.com", "cc": "", "cco": prep["cco"],
                                           "assunto": prep["assunto"], "corpo": prep["corpo"]})
check("e-mail enviado (simulado)", enviados and "E-mail enviado" in r.get_data(as_text=True))
r = cli.post(f"/envio/{token}/acao", data={"acao": "enviar", "para": "a@b.com", "cc": "", "cco": "",
                                           "assunto": "x", "corpo": "x"})
check("segundo clique em Enviar é bloqueado", len(enviados) == 1 and "já foi feito" in r.get_data(as_text=True))
c = controle.ler(cfg)
b = [a for a in c["atividades"] if a["mun"] == "Bonito" and a["data"] == dt.date(2026, 10, 20)][0]
check("CONTROLE registra status, nº de respostas e data/hora do envio",
      b["status"] == "Enviado" and b["n"] > 0 and b["enviado_em"] is not None, b)
log = [r for r in openpyxl.load_workbook(tmp / "CONTROLE.xlsx")["Log"].iter_rows(min_row=2, values_only=True) if r[0]]
check("Log registra o envio", len(log) == 1 and log[0][1] == "Bonito")
html = cli.get("/?semana=2026-10-19").get_data(as_text=True)
check("semana mostra a atividade enviada em verde, com data e hora", "✔ Enviado" in html and 'class="enviado"' in html
      and "Reenviar" in html)

r = cli.post("/preparar", data={"sel": ["2026-10-20|Bonito", "2026-10-20|Araripina"]})
check("selecionar municípios diferentes é recusado", r.status_code == 400)
print("\n" + ("TUDO OK" if ok else "HÁ FALHAS"))
raise SystemExit(0 if ok else 1)
