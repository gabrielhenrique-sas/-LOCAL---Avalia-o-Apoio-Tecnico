"""CONTROLE.xlsx: atividades (vindas do calendário), status dos envios, gerências, configurações e logs."""
import shutil
from datetime import datetime
from pathlib import Path

from openpyxl import Workbook, load_workbook
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.utils import get_column_letter

from .calendario import GERENCIAS_VALIDAS
from .config import caminho, pasta_saida
from .respostas import ErroDados
from .sync import ordenar
from .util import norm, parse_carimbo, parse_data

# (campo, título da coluna, largura)
COLUNAS = [("data", "Data", 12), ("mun", "Município", 28), ("modalidade", "Modalidade", 12),
           ("ger", "Gerências", 30), ("proc", "Nº do processo", 14), ("obs", "Observação", 40),
           ("revisar", "Revisar", 9), ("status", "Status do envio", 15), ("n", "Nº de respostas", 11),
           ("enviado_em", "Enviado em", 17), ("link", "Link do PDF", 30),
           ("situacao", "Situação no calendário", 22), ("origem", "Origem (aba/célula)", 18),
           ("assinatura", "Leitura do calendário", 20), ("manual", "Ajuste manual", 10),
           ("chave", "Chave", 30), ("atualizado", "Atualizado em", 17)]
SIM_NAO = ("revisar", "manual")
CONFIG_PADRAO = [("EMAIL_PARA", "", "Endereço fixo do campo Para. Os gerentes vão sempre em cópia oculta."),
                 ("NOME_REMETENTE", "", "Nome que aparece como remetente."),
                 ("ASSUNTO_RELATORIO", "", "Assunto do e-mail com relatório. Marcadores: {{MUNICIPIO}}, {{DATA}}."),
                 ("CORPO_RELATORIO", "", "Texto do e-mail com relatório."),
                 ("ASSUNTO_SEM_RESPOSTAS", "", "Assunto do aviso sem respostas."),
                 ("CORPO_SEM_RESPOSTAS", "", "Texto do aviso sem respostas.")]
CABECALHOS = {"Gerências": ["Gerência", "Gerente", "E-mail corporativo", "Ativa"],
              "Log": ["Execução", "Município", "Datas dos apoios", "Nº de respostas", "Nº de destinatários",
                      "Status", "Link do PDF", "Erro"],
              "Mudanças do calendário": ["Data/hora", "Evento", "Município", "Data do ATM", "Detalhe"]}


def _cabecalho(ws, titulos, larguras=None):
    for j, t in enumerate(titulos, start=1):
        c = ws.cell(1, j, t)
        c.font = Font(name="Arial", bold=True, color="FFFFFF")
        c.fill = PatternFill("solid", fgColor="1F3864")
        c.alignment = Alignment(horizontal="center", vertical="center")
        ws.column_dimensions[get_column_letter(j)].width = (larguras or [24] * len(titulos))[j - 1]
    ws.freeze_panes = "A2"


def _aba(wb, nome):
    alvo = norm(nome)
    return next((ws for ws in wb.worksheets if norm(ws.title) == alvo), None)


def _criar(cfg, arq):
    """Cria o CONTROLE.xlsx; copia Gerências e Configurações do modelo antigo, se informado."""
    modelo = None
    if cfg["controle"].get("modelo") and caminho(cfg, cfg["controle"]["modelo"]).exists():
        modelo = load_workbook(caminho(cfg, cfg["controle"]["modelo"]), data_only=True)
    wb = Workbook()
    ws = wb.active
    ws.title = "Atividades"
    _cabecalho(ws, [c[1] for c in COLUNAS], [c[2] for c in COLUNAS])
    wg = wb.create_sheet("Gerências")
    _cabecalho(wg, CABECALHOS["Gerências"], [16, 22, 40, 8])
    origem = _aba(modelo, "Gerências") if modelo else None
    linhas = [r for r in origem.iter_rows(min_row=2, max_col=4, values_only=True) if r[0]] if origem else \
        [(g, None, None, "Sim") for g in GERENCIAS_VALIDAS]
    for r in linhas:
        wg.append(list(r))
    wc = wb.create_sheet("Configurações")
    _cabecalho(wc, ["Chave", "Valor", "Descrição"], [26, 60, 80])
    antigos = {}
    origem = _aba(modelo, "Configurações") if modelo else None
    if origem:
        antigos = {str(r[0]).strip(): r[1] for r in origem.iter_rows(min_row=2, max_col=2, values_only=True) if r[0]}
    for chave, _, desc in CONFIG_PADRAO:
        wc.append([chave, antigos.get(chave), desc])
        wc.cell(wc.max_row, 2).alignment = Alignment(wrap_text=True, vertical="top")
    for nome in ("Log", "Mudanças do calendário"):
        _cabecalho(wb.create_sheet(nome), CABECALHOS[nome], [18, 28, 22, 13, 90, 12, 40, 30][:len(CABECALHOS[nome])])
    wb.save(arq)


def _arquivo(cfg):
    arq = caminho(cfg, cfg["controle"]["arquivo"])
    if not arq.exists():
        _criar(cfg, arq)
    return arq


# ---------------------------------------------------------------- leitura
def ler(cfg):
    """Devolve {'arquivo', 'atividades', 'gerencias': {nome: {...}}, 'config': {...}}."""
    arq = _arquivo(cfg)
    wb = load_workbook(arq, data_only=True)
    ctrl = {"arquivo": arq, "atividades": [], "gerencias": {}, "config": {}}
    ws = _aba(wb, "Atividades")
    if ws is None:
        raise ErroDados("O CONTROLE.xlsx não tem a aba 'Atividades'.")
    campos = [c[0] for c in COLUNAS]
    for i, row in enumerate(ws.iter_rows(min_row=2, max_col=len(campos), values_only=True), start=2):
        a = dict(zip(campos, row))
        a["data"] = parse_data(a["data"])
        if a["data"] is None or not a["mun"]:
            continue
        for k in SIM_NAO:
            a[k] = norm(a[k]) == "sim"
        for k in ("mun", "modalidade", "ger", "proc", "obs", "status", "link", "situacao", "origem", "assinatura", "chave"):
            a[k] = "" if a[k] is None else str(a[k]).strip()
        a["situacao"] = a["situacao"] or "Ativo"
        a["enviado_em"], a["atualizado"] = parse_carimbo(a["enviado_em"]), parse_carimbo(a["atualizado"])
        ctrl["atividades"].append(a)
    ws = _aba(wb, "Gerências")
    if ws is not None:
        for nome, gerente, email, ativa in ws.iter_rows(min_row=2, max_col=4, values_only=True):
            if nome:
                ctrl["gerencias"][str(nome).strip()] = {"gerente": gerente or "", "email": str(email or "").strip(),
                                                        "ativa": norm(ativa or "Sim") != "nao"}
    ws = _aba(wb, "Configurações")
    if ws is not None:
        for k, v in ws.iter_rows(min_row=2, max_col=2, values_only=True):
            if k and v not in (None, ""):
                ctrl["config"][str(k).strip()] = v
    wb.close()
    return ctrl


# ---------------------------------------------------------------- gravação
def salvar(cfg, atividades, eventos=(), logs=()):
    """Regrava a aba Atividades e acrescenta linhas em 'Mudanças do calendário' e 'Log'."""
    arq = _arquivo(cfg)
    bak = pasta_saida(cfg, "backups")
    shutil.copy2(arq, bak / f"{arq.stem}_{datetime.now():%Y%m%d_%H%M%S}{arq.suffix}")
    for antigo in sorted(bak.glob(f"{arq.stem}_*"))[:-20]:
        antigo.unlink(missing_ok=True)

    wb = load_workbook(arq)
    ws = _aba(wb, "Atividades")
    ws.delete_rows(2, ws.max_row)
    ordenar(atividades)
    for a in atividades:
        linha = [("SIM" if a[k] else "") if k in SIM_NAO else a[k] for k, _, _ in COLUNAS]
        ws.append(linha)
        r = ws.max_row
        for j, (k, _, _) in enumerate(COLUNAS, start=1):
            ws.cell(r, j).font = Font(name="Arial", size=10)
        ws.cell(r, 1).number_format = "DD/MM/YYYY"
        ws.cell(r, 5).number_format = "@"
        ws.cell(r, 10).number_format = "DD/MM/YYYY HH:MM"
        ws.cell(r, 17).number_format = "DD/MM/YYYY HH:MM"
        if a["revisar"]:
            for j in range(1, len(COLUNAS) + 1):
                ws.cell(r, j).fill = PatternFill("solid", fgColor="FFF2CC")
    ws.auto_filter.ref = f"A1:{get_column_letter(len(COLUNAS))}{max(ws.max_row, 2)}"

    agora = datetime.now().replace(microsecond=0)
    wm = _aba(wb, "Mudanças do calendário")
    for evento, mun, data, detalhe in eventos:
        wm.append([agora, evento, mun, data, detalhe])
        wm.cell(wm.max_row, 1).number_format = "DD/MM/YYYY HH:MM"
        wm.cell(wm.max_row, 4).number_format = "DD/MM/YYYY"
    wl = _aba(wb, "Log")
    for lg in logs:
        wl.append([lg.get("execucao") or agora, lg.get("municipio"), lg.get("datas"), lg.get("n"),
                   lg.get("destinatarios"), lg.get("status"), lg.get("link"), lg.get("erro")])
        wl.cell(wl.max_row, 1).number_format = "DD/MM/YYYY HH:MM"
    try:
        wb.save(arq)
    except PermissionError:
        raise ErroDados("Não consegui gravar o CONTROLE.xlsx. Feche-o se estiver aberto no Excel e tente de novo.")
