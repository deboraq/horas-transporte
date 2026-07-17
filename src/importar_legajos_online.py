"""
Importa exportaciones de Legajos Online (Detalle de Novedades).
"""
from __future__ import annotations

import re
from dataclasses import dataclass
from datetime import date
from pathlib import Path
from typing import Optional

import pandas as pd

from calculo_horas import _parse_date, _parse_time


@dataclass
class FilaLegajosOnline:
    id_registro: str
    fecha: date
    dia: str
    feriado: bool
    citacion: Optional[object]
    entrada: Optional[object]
    salida: Optional[object]
    total_min: int
    horas_min: int
    horas_50_min: int
    horas_100_min: int
    descanso_min: Optional[int]
    descanso_no_tomado_min: Optional[int]
    cenas: int
    horas_nocturnas_min: int
    turno: str
    turno_horario: str
    nombre_informe: str = ""


def _celda_minutos(val) -> int:
    if val is None or (isinstance(val, float) and pd.isna(val)):
        return 0
    if hasattr(val, "hour"):
        return val.hour * 60 + val.minute
    m = re.search(r"(\d{1,2}):(\d{2})", str(val))
    if m:
        return int(m.group(1)) * 60 + int(m.group(2))
    return 0


def _celda_hora(val):
    if val is None or (isinstance(val, float) and pd.isna(val)):
        return None
    if hasattr(val, "hour"):
        return val
    m = re.search(r"(\d{1,2}):(\d{2})(?::(\d{2}))?", str(val))
    if m:
        from datetime import time
        s = int(m.group(3) or 0)
        return time(int(m.group(1)), int(m.group(2)), s)
    return _parse_time(val)


def importar_legajos_online(path: Path) -> tuple[str, list[FilaLegajosOnline]]:
    df = pd.read_excel(path, header=None, engine="calamine")
    nombre = ""
    if len(df) > 1:
        ultima = df.iloc[-1, 0]
        if isinstance(ultima, str) and "informe mensual" in ultima.lower():
            nombre = ultima.replace("Informe mensual", "").strip()

    filas: list[FilaLegajosOnline] = []
    for i in range(2, len(df) - 1):
        r = df.iloc[i]
        if pd.isna(r[1]):
            continue
        fecha = _parse_date(r[1])
        if not fecha:
            continue
        filas.append(
            FilaLegajosOnline(
                id_registro=str(r[0] or ""),
                fecha=fecha,
                dia=str(r[2] or "").strip(),
                feriado=str(r[3] or "").strip().lower() in ("si", "sí", "s"),
                citacion=_celda_hora(r[4]),
                entrada=_celda_hora(r[5]),
                salida=_celda_hora(r[6]),
                total_min=_celda_minutos(r[7]),
                horas_min=_celda_minutos(r[8]),
                horas_50_min=_celda_minutos(r[9]),
                horas_100_min=_celda_minutos(r[10]),
                descanso_min=_celda_minutos(r[11]) or None,
                descanso_no_tomado_min=_celda_minutos(r[12]) or None,
                cenas=int(r[13]) if pd.notna(r[13]) and str(r[13]).isdigit() else 0,
                horas_nocturnas_min=_celda_minutos(r[14]),
                turno=str(r[17] or "").strip(),
                turno_horario=str(r[18] or "").strip(),
            )
        )
    return nombre, filas


def legajo_desde_nombre(nombre: str, personal: dict[int, dict]) -> Optional[int]:
    n = nombre.upper().strip()
    for leg, info in personal.items():
        if info["nombre"].upper() in n or n in info["nombre"].upper():
            return leg
    # Coincidencia solo por apellido
    apellido = n.split()[0] if n else ""
    for leg, info in personal.items():
        if apellido and apellido in info["nombre"].upper():
            return leg
    return None
