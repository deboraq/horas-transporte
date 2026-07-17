"""
Motor de cálculo — reglas Camioneros / Bacar / Legajos Online.

Notas del delegado:
- Lun-vie: 8 hs; extras al 50%; nocturnas 21:00-06:00 al 100%
- Si pisa hora nocturna → TODAS las extras al 100%
- Sábado: 4 hs; extra antes 13:00 → 50%; después → 100%
- Domingo: todo al 100%; mínimo 6 hs si trabajó
- Cena: 1 si trabaja después de las 21:00
- Descanso: mínimo 12 hs entre jornadas; lo no cumplido → columna DESCANSO 100%
- Total Hs: desde citación hasta salida (si hay citación)
"""
from __future__ import annotations

from dataclasses import dataclass
from datetime import date, datetime, time, timedelta
from typing import Optional


DIAS_ES = ["lunes", "martes", "miercoles", "jueves", "viernes", "sabado", "domingo"]
JORNADA_LABORAL = 8 * 60
JORNADA_SABADO = 4 * 60
DOMINGO_MINIMO = 6 * 60
DESCANSO_MINIMO = 12 * 60
NOCTURNA_INICIO = time(21, 0)
NOCTURNA_FIN = time(6, 0)
SABADO_EXTRA_100_DESDE = time(13, 0)
CENA_DESDE = time(21, 0)


@dataclass
class ResultadoDia:
    fecha: date
    dia_semana: str
    es_feriado: bool
    franco: bool
    citacion: Optional[time]
    entrada: Optional[time]
    salida: Optional[time]
    turno_transporte: bool
    total_min: int
    horas_min: int
    horas_50_min: int
    horas_100_min: int
    horas_nocturnas_min: int
    descanso_min: Optional[int]
    descanso_100_min: int
    descanso_ok: Optional[bool]
    cenas: int
    dias: int
    observacion: str = ""

    @property
    def total_hs(self) -> str:
        return _fmt_min(self.total_min)

    @property
    def horas(self) -> str:
        return _fmt_min(self.horas_min)

    @property
    def horas_50(self) -> str:
        return _fmt_min(self.horas_50_min)

    @property
    def horas_100(self) -> str:
        return _fmt_min(self.horas_100_min)

    @property
    def horas_nocturnas(self) -> str:
        return _fmt_min(self.horas_nocturnas_min)

    @property
    def descanso(self) -> str:
        if self.descanso_min is None:
            return ""
        return _fmt_min(self.descanso_min)

    @property
    def descanso_100(self) -> str:
        return _fmt_min(self.descanso_100_min)


def _fmt_min(minutos: int) -> str:
    if minutos <= 0:
        return "00:00"
    h, m = divmod(minutos, 60)
    return f"{h:02d}:{m:02d}"


def _parse_time(val) -> Optional[time]:
    if val is None or val == "":
        return None
    if isinstance(val, time):
        return val
    if isinstance(val, datetime):
        return val.time()
    if isinstance(val, timedelta):
        total = int(val.total_seconds())
        h, rem = divmod(total, 3600)
        m, _ = divmod(rem, 60)
        return time(h % 24, m, 0)
    s = str(val).strip().lower()
    if s in ("", "fco", "franco", "libre", "-"):
        return None
    for fmt in ("%H:%M:%S", "%H:%M"):
        try:
            return datetime.strptime(s, fmt).time()
        except ValueError:
            pass
    return None


def _parse_date(val) -> Optional[date]:
    if val is None or val == "":
        return None
    if isinstance(val, datetime):
        return val.date()
    if isinstance(val, date):
        return val
    s = str(val).strip()
    for fmt in ("%d/%m/%Y", "%Y-%m-%d", "%d-%m-%Y"):
        try:
            return datetime.strptime(s, fmt).date()
        except ValueError:
            pass
    return None


def _es_franco(val) -> bool:
    if val is None:
        return False
    s = str(val).strip().lower()
    return s in ("si", "sí", "s", "1", "true", "fco", "franco", "x")


def _es_feriado(val) -> bool:
    if val is None:
        return False
    s = str(val).strip().lower()
    return s in ("si", "sí", "s", "1", "true", "f")


def _es_transporte(val) -> bool:
    if val is None:
        return False
    return "transporte" in str(val).strip().lower()


def _intervalos_trabajo(
    fecha: date,
    entrada: time,
    salida: time,
    citacion: Optional[time] = None,
) -> tuple[list[tuple[datetime, datetime]], list[tuple[datetime, datetime]]]:
    inicio_comp = citacion if citacion else entrada
    comp_ini = datetime.combine(fecha, inicio_comp)
    comp_fin = datetime.combine(fecha, salida)
    if comp_fin <= comp_ini:
        comp_fin += timedelta(days=1)

    fich_ini = datetime.combine(fecha, entrada)
    fich_fin = datetime.combine(fecha, salida)
    if fich_fin <= fich_ini:
        fich_fin += timedelta(days=1)

    return [(comp_ini, comp_fin)], [(fich_ini, fich_fin)]


def _duracion(intervalos: list[tuple[datetime, datetime]]) -> int:
    return sum(int((f - i).total_seconds() // 60) for i, f in intervalos)


def _minutos_en_franja(inicio: datetime, fin: datetime, desde: time, hasta: time) -> int:
    total = 0
    dia = inicio.date()
    while dia <= fin.date():
        a = datetime.combine(dia, desde)
        b = datetime.combine(dia, hasta)
        if desde > hasta:
            b += timedelta(days=1)
        ini = max(inicio, a)
        end = min(fin, b)
        if end > ini:
            total += int((end - ini).total_seconds() // 60)
        dia += timedelta(days=1)
    return total


def _minutos_nocturnos(intervalos: list[tuple[datetime, datetime]]) -> int:
    total = 0
    for inicio, fin in intervalos:
        total += _minutos_en_franja(inicio, fin, NOCTURNA_INICIO, NOCTURNA_FIN)
    return total


def _minutos_sabado_antes_13(intervalos: list[tuple[datetime, datetime]]) -> int:
    total = 0
    for inicio, fin in intervalos:
        d = inicio.date()
        while d <= fin.date():
            if d.weekday() == 5:
                total += _minutos_en_franja(inicio, fin, time(6, 0), SABADO_EXTRA_100_DESDE)
            d += timedelta(days=1)
    return total


def _minutos_sabado_despues_13(intervalos: list[tuple[datetime, datetime]]) -> int:
    total = 0
    for inicio, fin in intervalos:
        d = inicio.date()
        while d <= fin.date():
            if d.weekday() == 5:
                total += _minutos_en_franja(inicio, fin, SABADO_EXTRA_100_DESDE, time(23, 59))
                total += _minutos_en_franja(inicio, fin, time(0, 0), NOCTURNA_FIN)
            d += timedelta(days=1)
    return total


def _corresponde_cena(intervalos: list[tuple[datetime, datetime]]) -> bool:
    for inicio, fin in intervalos:
        if _minutos_en_franja(inicio, fin, CENA_DESDE, time(23, 59)) > 0:
            return True
        if _minutos_en_franja(inicio, fin, time(0, 0), NOCTURNA_FIN) > 0:
            return True
    return False


def _pago_domingo(total: int, nocturnas: int) -> tuple[int, int, int, int]:
    """Domingo: todo al 100%; mínimo 6 hs si trabajó."""
    if total <= 0:
        return 0, 0, 0, 0
    p100 = max(total, DOMINGO_MINIMO)
    return total, 0, 0, p100, nocturnas


def _calcular_pagos(
    intervalos: list[tuple[datetime, datetime]],
    fecha: date,
    es_feriado: bool,
    turno_transporte: bool,
) -> tuple[int, int, int, int, int]:
    total = _duracion(intervalos)
    if total == 0:
        return 0, 0, 0, 0, 0

    nocturnas = _minutos_nocturnos(intervalos)
    dow = fecha.weekday()

    if es_feriado:
        return total, 0, 0, total, nocturnas

    if dow == 6:
        return _pago_domingo(total, nocturnas)

    jornada = JORNADA_SABADO if dow == 5 else JORNADA_LABORAL
    norm = min(total, jornada)
    extra = max(0, total - jornada)

    # Sábado: jornada 4 hs; antes 13:00 → 50%; después y nocturnas → 100%
    if dow == 5:
        despues_13 = _minutos_sabado_despues_13(intervalos)
        en_extra = max(0, total - jornada)
        if nocturnas > 0:
            p50 = 0
            p100 = max(extra, nocturnas, despues_13)
        else:
            antes_13_extra = max(0, en_extra - despues_13)
            p50 = antes_13_extra
            p100 = despues_13
        return total, norm, p50, p100, nocturnas

    # Transporte: todo el extra al 100%
    if turno_transporte:
        return total, norm, 0, extra, nocturnas

    # Cronograma lun-vie
    if nocturnas > 0:
        # Si pisa hora nocturna → todas las extras al 100%
        p50 = 0
        p100 = max(extra, nocturnas)
    else:
        p50 = extra
        p100 = 0

    return total, norm, p50, p100, nocturnas


def _descanso_100(descanso_min: Optional[int]) -> int:
    if descanso_min is None:
        return 0
    if descanso_min >= DESCANSO_MINIMO:
        return 0
    return DESCANSO_MINIMO - descanso_min


def calcular_dia(
    fecha: date,
    entrada: Optional[time],
    salida: Optional[time],
    *,
    citacion: Optional[time] = None,
    franco: bool = False,
    es_feriado: bool = False,
    turno_transporte: bool = False,
    salida_anterior: Optional[datetime] = None,
    entrada_actual_dt: Optional[datetime] = None,
    observacion: str = "",
) -> ResultadoDia:
    dia_nom = DIAS_ES[fecha.weekday()]
    descanso_min = None
    descanso_ok = None
    if salida_anterior and entrada_actual_dt:
        descanso_min = int((entrada_actual_dt - salida_anterior).total_seconds() // 60)
        descanso_ok = descanso_min >= DESCANSO_MINIMO

    vacio = ResultadoDia(
        fecha=fecha,
        dia_semana=dia_nom,
        es_feriado=es_feriado,
        franco=franco,
        citacion=citacion,
        entrada=entrada,
        salida=salida,
        turno_transporte=turno_transporte,
        total_min=0,
        horas_min=0,
        horas_50_min=0,
        horas_100_min=0,
        horas_nocturnas_min=0,
        descanso_min=descanso_min,
        descanso_100_min=_descanso_100(descanso_min),
        descanso_ok=descanso_ok,
        cenas=0,
        dias=0,
        observacion=observacion,
    )

    if franco or entrada is None or salida is None:
        return vacio

    intervalos_comp, intervalos_fich = _intervalos_trabajo(fecha, entrada, salida, citacion)
    entrada_dt = datetime.combine(fecha, entrada)
    if salida_anterior and entrada_dt:
        descanso_min = int((entrada_dt - salida_anterior).total_seconds() // 60)
        descanso_ok = descanso_min >= DESCANSO_MINIMO

    total, norm, p50, p100, noct = _calcular_pagos(
        intervalos_comp, fecha, es_feriado, turno_transporte
    )
    cenas = 1 if _corresponde_cena(intervalos_fich) else 0

    return ResultadoDia(
        fecha=fecha,
        dia_semana=dia_nom,
        es_feriado=es_feriado,
        franco=False,
        citacion=citacion,
        entrada=entrada,
        salida=salida,
        turno_transporte=turno_transporte,
        total_min=total,
        horas_min=norm,
        horas_50_min=p50,
        horas_100_min=p100,
        horas_nocturnas_min=noct,
        descanso_min=descanso_min,
        descanso_100_min=_descanso_100(descanso_min),
        descanso_ok=descanso_ok,
        cenas=cenas,
        dias=1,
        observacion=observacion,
    )


def rango_periodo(anio: int, mes_fin: int) -> list[date]:
    if mes_fin == 1:
        inicio = date(anio - 1, 12, 26)
    else:
        inicio = date(anio, mes_fin - 1, 26)
    fin = date(anio, mes_fin, 25)
    dias = []
    d = inicio
    while d <= fin:
        dias.append(d)
        d += timedelta(days=1)
    return dias


def sumar_minutos(resultados: list[ResultadoDia]) -> dict[str, int]:
    return {
        "total": sum(r.total_min for r in resultados),
        "horas": sum(r.horas_min for r in resultados),
        "horas_50": sum(r.horas_50_min for r in resultados),
        "horas_100": sum(r.horas_100_min for r in resultados),
        "nocturnas": sum(r.horas_nocturnas_min for r in resultados),
        "descanso_100": sum(r.descanso_100_min for r in resultados),
        "cenas": sum(r.cenas for r in resultados),
        "dias": sum(r.dias for r in resultados),
    }
