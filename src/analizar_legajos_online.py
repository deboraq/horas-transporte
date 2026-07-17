"""Analiza archivos Legajos Online para inferir reglas."""
import re
import sys
from pathlib import Path

import pandas as pd

sys.path.insert(0, str(Path(__file__).resolve().parent))
from calculo_horas import _parse_time, calcular_dia, _parse_date


def to_min(v):
    if v is None or (isinstance(v, float) and pd.isna(v)):
        return None
    if hasattr(v, "hour"):
        return v.hour * 60 + v.minute
    m = re.search(r"(\d{1,2}):(\d{2})", str(v))
    if m:
        return int(m.group(1)) * 60 + int(m.group(2))
    return None


def dur(ini, fin):
    if ini is None or fin is None:
        return None
    return fin - ini if fin >= ini else 24 * 60 - ini + fin


def load(path):
    df = pd.read_excel(path, header=None, engine="calamine")
    rows = []
    for i in range(2, len(df) - 1):
        r = df.iloc[i]
        rows.append({
            "fecha": r[1],
            "cit": to_min(r[4]),
            "ent": to_min(r[5]),
            "sal": to_min(r[6]),
            "total_lo": to_min(r[7]),
            "norm": to_min(r[8]),
            "p50": to_min(r[9]),
            "p100": to_min(r[10]),
            "noct": to_min(r[14]),
            "dia": str(r[2]).strip().lower(),
            "turno": str(r[17]) if pd.notna(r[17]) else "",
        })
    return rows


def tmin(m):
    if m is None:
        return "  - "
    return f"{m:4d}"


files = [
    ("PERSONA_A", Path(r"c:\Users\Usr\Downloads\Detalle de Novedades  Legajos Online.xlsx")),
    ("PERSONA_B", Path(r"c:\Users\Usr\Downloads\Detalle de Novedades  Legajos Online (1).xlsx")),
]

for label, path in files:
    print("=" * 90)
    print(label, path.name)
    for x in load(path):
        cit, ent, sal = x["cit"], x["ent"], x["sal"]
        tc, te = dur(cit, sal), dur(ent, sal)
        fecha = _parse_date(x["fecha"])
        if fecha and ent is not None and sal is not None:
            eh = ent // 60
            em = ent % 60
            sh = sal // 60
            sm = sal % 60
            ours = calcular_dia(
                fecha,
                _parse_time(f"{eh:02d}:{em:02d}"),
                _parse_time(f"{sh:02d}:{sm:02d}"),
                citacion=_parse_time(f"{cit//60:02d}:{cit%60:02d}") if cit else None,
            )
            diff100 = (ours.horas_100_min or 0) - (x["p100"] or 0)
            diff50 = (ours.horas_50_min or 0) - (x["p50"] or 0)
            difft = (ours.total_min or 0) - (x["total_lo"] or 0)
        else:
            difft = diff50 = diff100 = 0

        print(
            f"{x['fecha']} {x['dia'][:9]:9} "
            f"LO tot={tmin(x['total_lo'])} cit={tmin(tc)} ent={tmin(te)} "
            f"n={tmin(x['norm'])} 50={tmin(x['p50'])} 100={tmin(x['p100'])} noct={tmin(x['noct'])} | "
            f"NUESTRO dT={difft:+4d} d50={diff50:+4d} d100={diff100:+4d} {x['turno'][:12]}"
        )
