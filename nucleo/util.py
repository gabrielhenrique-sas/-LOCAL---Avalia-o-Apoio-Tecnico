"""Funções pequenas usadas em vários módulos: texto, datas e semana."""
import re
import unicodedata
from datetime import date, datetime, timedelta

DIAS = ["Segunda", "Terça", "Quarta", "Quinta", "Sexta", "Sábado", "Domingo"]


def norm(s):
    """minúsculas, sem acentos, espaços normalizados."""
    s = unicodedata.normalize("NFD", "" if s is None else str(s))
    s = "".join(ch for ch in s if unicodedata.category(ch) != "Mn")
    return re.sub(r"\s+", " ", s.lower()).strip()


def slug(s):
    s = unicodedata.normalize("NFKD", str(s))
    s = "".join(c for c in s if not unicodedata.combining(c))
    return re.sub(r"[^A-Za-z0-9]+", "_", s).strip("_")


# ---------------------------------------------------------------- datas
_FORMATOS_DATA = ("%d/%m/%Y", "%Y-%m-%d", "%d/%m/%y", "%d-%m-%Y", "%d.%m.%Y")
_FORMATOS_CARIMBO = ("%d/%m/%Y %H:%M:%S", "%d/%m/%Y %H:%M", "%Y-%m-%d %H:%M:%S",
                     "%Y-%m-%dT%H:%M:%S", "%Y-%m-%d %H:%M")


def parse_data(v):
    if v is None or v == "":
        return None
    if isinstance(v, datetime):
        return v.date()
    if isinstance(v, date):
        return v
    if isinstance(v, (int, float)) and v > 30000:
        return (datetime(1899, 12, 30) + timedelta(days=float(v))).date()
    s = re.split(r"[ T]", str(v).strip())[0]
    for f in _FORMATOS_DATA:
        try:
            return datetime.strptime(s, f).date()
        except ValueError:
            pass
    return None


def parse_carimbo(v):
    if v is None or v == "":
        return None
    if isinstance(v, datetime):
        return v
    if isinstance(v, date):
        return datetime.combine(v, datetime.min.time())
    s = str(v).strip()
    for f in _FORMATOS_CARIMBO:
        try:
            return datetime.strptime(s, f)
        except ValueError:
            pass
    d = parse_data(s)
    return datetime.combine(d, datetime.min.time()) if d else None


def fmt_data(d):
    return d.strftime("%d/%m/%Y")


def fmt_hora(dh):
    return dh.strftime("%d/%m/%Y %H:%M") if dh else ""


def texto_datas(datas):
    t = [fmt_data(d) for d in sorted(datas)]
    if len(t) <= 1:
        return "".join(t)
    return ", ".join(t[:-1]) + " e " + t[-1]


def segunda_de(d):
    """Segunda-feira da semana que contém d."""
    return d - timedelta(days=d.weekday())
