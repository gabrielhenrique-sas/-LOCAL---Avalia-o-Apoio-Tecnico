"""Cálculo das contagens, gráficos e geração do PDF do relatório."""
import textwrap
from collections import Counter
from datetime import datetime
from pathlib import Path
from xml.sax.saxutils import escape

import matplotlib

matplotlib.use("Agg")
import matplotlib.pyplot as plt  # noqa: E402
from matplotlib import font_manager  # noqa: E402
from reportlab.lib import colors  # noqa: E402
from reportlab.lib.enums import TA_CENTER, TA_RIGHT  # noqa: E402
from reportlab.lib.pagesizes import A4  # noqa: E402
from reportlab.lib.styles import ParagraphStyle  # noqa: E402
from reportlab.lib.units import cm  # noqa: E402
from reportlab.platypus import (BaseDocTemplate, Frame, HRFlowable, Image, KeepTogether, PageBreak,  # noqa: E402
                                PageTemplate, Paragraph, Spacer, Table, TableStyle)

from .util import fmt_data, norm, texto_datas  # noqa: E402

TINTA = "#1F2937"
AZUL = "#1F3864"
COR_BARRA = "#2F5FA7"
CINZA = "#6B7280"

_disp = {f.name for f in font_manager.fontManager.ttflist}
for _f in ("Arial", "Liberation Sans", "DejaVu Sans"):
    if _f in _disp:
        plt.rcParams["font.family"] = _f
        break

ESCALA_SIM = ["Sim, totalmente", "Parcialmente", "Não atendeu"]
ORDENS = {
    "esclareceu": ESCALA_SIM,
    "formato": ESCALA_SIM,
    "tempo": ["Curto demais", "Adequado", "Longo demais"],
    "explicar": ESCALA_SIM,
    "geral": ["Excelente", "Bom", "Regular", "Precisa melhorar"],
}

# (título da seção, enunciado, campo da resposta fechada, campo do comentário)
CATEGORIAS = [
    ("Esclarecimentos", "O Apoio Técnico esclareceu o que o município precisava?", "esclareceu", "comentario_esclareceu"),
    ("Formato", "O formato do Apoio Técnico atendeu as expectativas do município?", "formato", "comentario_formato"),
    ("Tempo de duração", "Como você avalia o tempo de duração do Apoio Técnico?", "tempo", None),
    ("Forma de explicar", "A forma de explicar (exemplos, linguagem usada e material usado) auxiliou na compreensão?",
     "explicar", "comentario_explicar"),
]
GERAL = ("Avaliação geral", "No geral, como você avalia o Apoio Técnico recebido?", "geral", "comentario_geral")


# ------------------------------------------------------------------ cálculo
def contar(respostas, campo):
    ordem = ORDENS[campo]
    por_norm = {norm(o): o for o in ordem}
    c = Counter()
    for r in respostas:
        v = r.get(campo, "")
        if not v:
            continue
        c[por_norm.get(norm(v), "Outras respostas")] += 1
    saida = [(o, c.get(o, 0)) for o in ordem]
    if c.get("Outras respostas"):
        saida.append(("Outras respostas", c["Outras respostas"]))
    return saida


def comentarios(respostas, campo_comentario, campo_fechado):
    itens = []
    for r in respostas:
        t = (r.get(campo_comentario) or "").strip()
        if t:
            itens.append((r.get(campo_fechado) or "-", t))
    return itens


def calcular(respostas):
    setores = Counter((r.get("setor") or "(não informado)") for r in respostas)
    setores = sorted(setores.items(), key=lambda kv: (-kv[1], kv[0]))
    blocos = []
    for titulo, enunciado, campo, campo_com in CATEGORIAS + [GERAL]:
        blocos.append({
            "titulo": titulo, "enunciado": enunciado, "campo": campo,
            "contagem": contar(respostas, campo),
            "comentarios": comentarios(respostas, campo_com, campo) if campo_com else [],
        })
    return {"n": len(respostas), "setores": setores, "blocos": blocos}


# ------------------------------------------------------------------ gráficos
def _rotulo(v, total):
    return f"{v} ({100 * v / total:.0f}%)" if total else "0"


def grafico(contagem, arquivo, horizontal=False):
    """Barras de uma só cor, com valor e percentual em cada barra. Devolve (largura, altura) em polegadas."""
    rot = [textwrap.fill(r, 46 if horizontal else 16) for r, _ in contagem]
    vals = [n for _, n in contagem]
    total = sum(vals)
    n = len(vals)
    largura = 15 / 2.54
    altura = max(1.2, 0.34 * n + 0.5) if horizontal else 5.6 / 2.54
    fig, ax = plt.subplots(figsize=(largura, altura))
    maximo = max(vals) if vals and max(vals) > 0 else 1
    pos = list(range(n))
    if horizontal:
        ax.barh(pos, vals, height=0.55, color=COR_BARRA)
        ax.set_yticks(pos)
        ax.set_yticklabels(rot, fontsize=7.5, color=TINTA)
        ax.invert_yaxis()
        ax.set_xlim(0, maximo * 1.35)
        ax.set_xticks([])
        for i, v in enumerate(vals):
            ax.text(v + maximo * 0.02, i, _rotulo(v, total), va="center", fontsize=7.5, color=TINTA)
        ocultar = ("top", "right", "bottom")
    else:
        ax.bar(pos, vals, width=0.5, color=COR_BARRA)
        ax.set_xticks(pos)
        ax.set_xticklabels(rot, fontsize=8, color=TINTA)
        ax.set_ylim(0, maximo * 1.3)
        ax.set_yticks([])
        for i, v in enumerate(vals):
            ax.text(i, v + maximo * 0.03, _rotulo(v, total), ha="center", va="bottom", fontsize=8, color=TINTA)
        ocultar = ("top", "right", "left")
    for s in ocultar:
        ax.spines[s].set_visible(False)
    for s in ax.spines.values():
        s.set_color("#9CA3AF")
    ax.tick_params(length=0)
    fig.tight_layout()
    fig.savefig(arquivo, dpi=200, transparent=True)
    plt.close(fig)
    return largura, altura


# ------------------------------------------------------------------ PDF
def _estilos():
    base = dict(fontName="Helvetica", textColor=colors.HexColor(TINTA))
    return {
        "titulo": ParagraphStyle("titulo", fontName="Helvetica-Bold", fontSize=16, leading=20, alignment=TA_CENTER,
                                 textColor=colors.black),
        "mun": ParagraphStyle("mun", fontName="Helvetica-Bold", fontSize=14, leading=18, textColor=colors.black),
        "info": ParagraphStyle("info", fontSize=10, leading=14, **base),
        "h1": ParagraphStyle("h1", fontName="Helvetica-Bold", fontSize=13, leading=17, textColor=colors.HexColor(AZUL),
                             spaceBefore=6, spaceAfter=2, keepWithNext=1),
        "h2": ParagraphStyle("h2", fontName="Helvetica-Bold", fontSize=11.5, leading=15,
                             textColor=colors.HexColor(AZUL), spaceBefore=10, spaceAfter=3, keepWithNext=1),
        "h3": ParagraphStyle("h3", fontName="Helvetica-Bold", fontSize=10, leading=13, spaceBefore=8, spaceAfter=2,
                             textColor=colors.HexColor(TINTA), keepWithNext=1),
        "perg": ParagraphStyle("perg", fontSize=9, leading=12, textColor=colors.HexColor(CINZA), fontName="Helvetica"),
        "td": ParagraphStyle("td", fontSize=8.5, leading=11, **base),
        "th": ParagraphStyle("th", fontName="Helvetica-Bold", fontSize=8.5, leading=11, textColor=colors.white),
        "nota": ParagraphStyle("nota", fontSize=9, leading=12, textColor=colors.HexColor(CINZA), fontName="Helvetica"),
    }


def _par(txt, estilo):
    return Paragraph(escape(str(txt)).replace("\n", "<br/>"), estilo)


def _tabela_comentarios(itens, st):
    dados = [[Paragraph("Resposta", st["th"]), Paragraph("Comentário", st["th"])]]
    for resp, texto in itens:
        dados.append([_par(resp, st["td"]), _par(texto, st["td"])])
    t = Table(dados, colWidths=[3.4 * cm, 11.6 * cm], repeatRows=1)
    t.setStyle(TableStyle([
        ("BACKGROUND", (0, 0), (-1, 0), colors.HexColor(AZUL)),
        ("VALIGN", (0, 0), (-1, -1), "TOP"),
        ("GRID", (0, 0), (-1, -1), 0.4, colors.HexColor("#BFBFBF")),
        ("ROWBACKGROUNDS", (0, 1), (-1, -1), [colors.white, colors.HexColor("#F3F4F6")]),
        ("LEFTPADDING", (0, 0), (-1, -1), 5), ("RIGHTPADDING", (0, 0), (-1, -1), 5),
        ("TOPPADDING", (0, 0), (-1, -1), 3), ("BOTTOMPADDING", (0, 0), (-1, -1), 4),
    ]))
    return t


def _imagem(contagem, pasta, nome, horizontal=False):
    arq = Path(pasta) / f"{nome}.png"
    larg, alt = grafico(contagem, str(arq), horizontal)
    return Image(str(arq), width=15 * cm, height=15 * cm * alt / larg)


def gerar_pdf(destino, municipio, secoes, gerencias, pasta_tmp, papel_timbrado):
    """secoes: lista de {'data': date, 'respostas': [...]}. Devolve o caminho do PDF."""
    st = _estilos()
    destino = Path(destino)
    destino.parent.mkdir(parents=True, exist_ok=True)
    pasta_tmp = Path(pasta_tmp)
    pasta_tmp.mkdir(parents=True, exist_ok=True)
    gerado_em = datetime.now().strftime("%d/%m/%Y %H:%M")

    def fundo(canvas, doc):
        canvas.saveState()
        if papel_timbrado and Path(papel_timbrado).exists():
            canvas.drawImage(str(papel_timbrado), 0, 0, width=A4[0], height=A4[1])
        canvas.setFont("Helvetica", 7.5)
        canvas.setFillColor(colors.HexColor(CINZA))
        canvas.drawRightString(A4[0] - 3 * cm, 4.25 * cm,
                               f"Gerado em {gerado_em}  |  Avaliações anônimas  |  Página {doc.page}")
        canvas.restoreState()

    doc = BaseDocTemplate(str(destino), pagesize=A4, leftMargin=3 * cm, rightMargin=3 * cm,
                          topMargin=3 * cm, bottomMargin=4.8 * cm,
                          title="Relatório de Avaliações do Apoio Técnico", author="")
    quadro = Frame(3 * cm, 4.8 * cm, A4[0] - 6 * cm, A4[1] - 7.8 * cm, id="corpo",
                   leftPadding=0, rightPadding=0, topPadding=0, bottomPadding=0)
    doc.addPageTemplates([PageTemplate(id="p", frames=[quadro], onPage=fundo)])

    datas = [s["data"] for s in secoes]
    modalidades = []
    for s in secoes:
        for m in s.get("modalidades", []):
            if m not in modalidades:
                modalidades.append(m)
    historia = [
        Paragraph("RELATÓRIO DE AVALIAÇÕES DO APOIO TÉCNICO", st["titulo"]),
        Spacer(1, 14),
        Paragraph(f"Município: {escape(municipio)}", st["mun"]),
        Spacer(1, 6),
        Paragraph(f"<b>Data do Apoio Técnico:</b> {texto_datas(datas)}", st["info"]),
        Paragraph(f"<b>Modalidade:</b> {escape(' e '.join(modalidades)) if modalidades else '-'}", st["info"]),
        Paragraph(f"<b>Gerências Participantes:</b> {escape(', '.join(gerencias)) if gerencias else '-'}", st["info"]),
        Spacer(1, 6),
        HRFlowable(width="100%", thickness=0.6, color=colors.HexColor("#9CA3AF")),
    ]

    for i, s in enumerate(secoes):
        if i > 0:
            historia.append(PageBreak())
        r = s["respostas"]
        titulo_sec = f"Apoio Técnico de {fmt_data(s['data'])}"
        if len(secoes) > 1 and s.get("modalidades"):          # com várias datas, cada seção diz a sua modalidade
            titulo_sec += f" - {' e '.join(s['modalidades'])}"
        if not r:
            historia += [Paragraph(titulo_sec, st["h1"]),
                         Paragraph("Não houve avaliações recebidas para esta data.", st["nota"])]
            continue
        calc = calcular(r)
        n = calc["n"]
        historia.append(Paragraph(f"{titulo_sec} ({n} {'resposta' if n == 1 else 'respostas'})", st["h1"]))
        tag = f"s{i}"

        historia.append(KeepTogether([
            Paragraph("Distribuição por setor", st["h2"]),
            _imagem(calc["setores"], pasta_tmp, f"{tag}_setor", horizontal=True),
        ]))
        historia.append(Paragraph("Avaliação por categoria", st["h2"]))
        for b in calc["blocos"][:-1]:
            historia.append(KeepTogether([
                Paragraph(b["titulo"], st["h3"]),
                Paragraph(escape(b["enunciado"]), st["perg"]),
                _imagem(b["contagem"], pasta_tmp, f"{tag}_{b['campo']}"),
            ]))
            if b["comentarios"]:
                historia.append(Paragraph(f"Comentários ({len(b['comentarios'])})", st["h3"]))
                historia.append(_tabela_comentarios(b["comentarios"], st))
        g = calc["blocos"][-1]
        historia.append(KeepTogether([
            Paragraph("Avaliação geral", st["h2"]),
            Paragraph(escape(g["enunciado"]), st["perg"]),
            _imagem(g["contagem"], pasta_tmp, f"{tag}_geral"),
        ]))
        if g["comentarios"]:
            historia.append(Paragraph(f"Comentários, elogios e sugestões ({len(g['comentarios'])})", st["h3"]))
            historia.append(_tabela_comentarios(g["comentarios"], st))

    doc.build(historia)
    return str(destino)
