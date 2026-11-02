"""
Mescla os ATMs lidos do calendário com as atividades já guardadas no CONTROLE.xlsx.

  ATM novo        -> entra
  ATM alterado    -> atualizado (registrado antes → depois)
  ATM que sumiu   -> Situação "Removido do calendário" (nada é apagado)
  ATM que voltou  -> Situação volta para "Ativo"

Só os ATMs da janela lida (segunda a sexta) podem ser marcados como removidos.
Ajuste manual (aba Revisão) vale enquanto a leitura do calendário daquele ATM não mudar.
"""
import datetime as dt

from .util import norm

CAMPOS = [("mun", "Município"), ("ger", "Gerências"), ("modalidade", "Modalidade"),
          ("obs", "Observação"), ("origem", "Origem")]
FRACAO_MINIMA = 0.5     # leitura com menos que isso dos ATMs ativos da janela é tratada como suspeita


def _assinatura(l):
    return "|".join([l["mun"], l["ger"], l["modalidade"], l["obs"], l["revisar"]])


def data_da_chave(a):
    """Data original lida do calendário (a data da atividade pode ter sido corrigida à mão)."""
    try:
        return dt.date.fromisoformat(a["chave"][:10])
    except ValueError:
        return a["data"]


def nova_atividade(l, agora):
    return {"data": l["data"], "mun": l["mun"], "modalidade": l["modalidade"], "ger": l["ger"],
            "proc": l["proc"], "obs": l["obs"], "revisar": l["revisar"] == "SIM", "status": "", "n": None,
            "enviado_em": None, "link": "", "situacao": "Ativo", "origem": f"{l['mes']} {l['celula']}",
            "assinatura": _assinatura(l), "manual": False, "chave": l["chave"], "atualizado": agora}


def mesclar(atividades, linhas, ini, fim, agora, forcar=False):
    """Altera 'atividades' no lugar. Devolve eventos [(evento, município, data, detalhe)]."""
    ev = []
    na_janela = lambda d: d is not None and ini <= d <= fim
    ativos = [a for a in atividades if a["situacao"] == "Ativo" and na_janela(data_da_chave(a))]
    suspeito = (ativos and not linhas) or (len(ativos) >= 4 and len(linhas) < len(ativos) * FRACAO_MINIMA)
    if suspeito and not forcar:
        return [("ABORTADO", "", None,
                 f"Leitura suspeita ({len(linhas)} ATMs lidos contra {len(ativos)} já cadastrados na semana). "
                 "Nada foi alterado.")]

    por_chave = {a["chave"]: a for a in atividades}
    vistas = set()
    for l in linhas:
        vistas.add(l["chave"])
        a = por_chave.get(l["chave"])
        if a is None:
            a = nova_atividade(l, agora)
            atividades.append(a)
            por_chave[a["chave"]] = a
            ev.append(("NOVO", a["mun"], a["data"], f"{a['ger']} ({a['modalidade']})"))
            continue
        reativado = a["situacao"] != "Ativo"
        a["situacao"] = "Ativo"
        nova = _assinatura(l)
        difs = ["voltou ao calendário"] if reativado else []
        if not (a["manual"] and a["assinatura"] == nova):          # ajuste manual preservado se nada mudou
            novos = {"mun": l["mun"], "ger": l["ger"], "modalidade": l["modalidade"], "obs": l["obs"],
                     "origem": f"{l['mes']} {l['celula']}"}
            for k, rotulo in CAMPOS:
                if str(a[k] or "") != novos[k]:
                    difs.append(f'{rotulo}: "{a[k] or ""}" → "{novos[k]}"')
                    a[k] = novos[k]
            a["revisar"] = l["revisar"] == "SIM"
            a["assinatura"], a["manual"] = nova, False
        if difs:
            a["atualizado"] = agora
            ev.append(("REATIVADO" if reativado else "ALTERADO", a["mun"], a["data"], " | ".join(difs)))

    for a in ativos:
        if a["chave"] not in vistas:
            a["situacao"], a["atualizado"] = "Removido do calendário", agora
            ev.append(("REMOVIDO", a["mun"], a["data"], "não está mais no calendário"))
    return ev


def ordenar(atividades):
    atividades.sort(key=lambda a: (a["data"] or dt.date.max, a["modalidade"] != "Presencial", norm(a["mun"])))
