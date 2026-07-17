/**
 * Control de Horas Transporte — Google Apps Script
 * Pegar en: Extensiones > Apps Script (después de subir Control_Horas.xlsx a Google Sheets)
 */
const JORNADA = 8 * 60;
const JORNADA_SAB = 4 * 60;
const DOMINGO_MIN = 6 * 60;
const DESCANSO_MIN = 12 * 60;

function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('Horas')
    .addItem('Calcular período', 'calcularPeriodo')
  .addToUi();
}

function calcularPeriodo() {
  const ss = SpreadsheetApp.getActive();
  const carga = ss.getSheetByName('Carga');
  const resultado = ss.getSheetByName('Resultado');
  const personal = leerPersonal(ss);

  const datos = carga.getDataRange().getValues();
  const filas = [];
  for (let i = 1; i < datos.length; i++) {
    const r = datos[i];
    if (!r[0] || !r[1]) continue;
    const fecha = parseFecha(r[0]);
    if (!fecha) continue;
    filas.push({
      fecha,
      legajo: Number(r[1]),
      nombre: r[2] || personal[Number(r[1])] || '',
      entrada: parseHora(r[3]),
      salida: parseHora(r[4]),
      citacion: parseHora(r[5]),
      franco: esSi(r[6]),
      feriado: esSi(r[7]),
      turno: String(r[8] || 'cronograma').toLowerCase(),
    });
  }

  filas.sort((a, b) => a.fecha - b.fecha || a.legajo - b.legajo);

  let salidaAnterior = null;
  const out = [['Fecha', 'Legajo', 'Nombre', 'Total', 'Normales', '50%', '100%', 'Nocturnas', 'Descanso', 'Descanso 100%', 'Cenas']];

  filas.forEach(f => {
    const res = calcularDia(f, salidaAnterior);
    if (f.entrada && f.salida && !f.franco) {
      salidaAnterior = finTurno(f.fecha, f.entrada, f.salida);
    }
    out.push([
      formatFecha(f.fecha),
      f.legajo,
      f.nombre,
      fmt(res.total),
      fmt(res.norm),
      fmt(res.p50),
      fmt(res.p100),
      fmt(res.noct),
      res.descanso !== null ? fmt(res.descanso) : '',
      fmt(res.desc100),
      res.cenas,
    ]);
  });

  resultado.clear();
  resultado.getRange(1, 1, out.length, out[0].length).setValues(out);
  resultado.getRange(1, 1, 1, out[0].length)
    .setBackground('#375623').setFontColor('#ffffff').setFontWeight('bold');
  SpreadsheetApp.getUi().alert('Listo: ' + (out.length - 1) + ' filas calculadas en hoja Resultado.');
}

function leerPersonal(ss) {
  const sh = ss.getSheetByName('Personal');
  const map = {};
  if (!sh) return map;
  const v = sh.getDataRange().getValues();
  for (let i = 1; i < v.length; i++) {
    if (v[i][0]) map[Number(v[i][0])] = String(v[i][1]);
  }
  return map;
}

function calcularDia(f, salidaAnterior) {
  const vacio = { total: 0, norm: 0, p50: 0, p100: 0, noct: 0, descanso: null, desc100: 0, cenas: 0 };
  if (f.franco || !f.entrada || !f.salida) return vacio;

  const entradaDt = combinar(f.fecha, f.entrada);
  let descanso = null;
  if (salidaAnterior) {
    descanso = Math.floor((entradaDt - salidaAnterior) / 60000);
  }
  const desc100 = descanso !== null && descanso < DESCANSO_MIN ? DESCANSO_MIN - descanso : 0;

  const inicioComp = combinar(f.fecha, f.citacion || f.entrada);
  const fin = combinar(f.fecha, f.salida);
  if (fin <= inicioComp) fin.setDate(fin.getDate() + 1);

  const total = Math.floor((fin - inicioComp) / 60000);
  const noct = minutosNocturnos(inicioComp, fin);
  const dow = f.fecha.getDay(); // 0=dom
  const transporte = f.turno.indexOf('transporte') >= 0;

  if (f.feriado || dow === 0) {
    const p100 = total > 0 ? Math.max(total, DOMINGO_MIN) : 0;
    return { total, norm: 0, p50: 0, p100, noct, descanso, desc100, cenas: cena(f.entrada, f.salida, f.fecha) };
  }

  const jornada = dow === 6 ? JORNADA_SAB : JORNADA;
  const norm = Math.min(total, jornada);
  const extra = Math.max(0, total - jornada);

  let p50 = 0, p100 = 0;

  if (dow === 6) {
    const desp13 = minSabDespues13(inicioComp, fin);
    if (noct > 0) { p50 = 0; p100 = Math.max(extra, noct, desp13); }
    else { p50 = Math.max(0, extra - desp13); p100 = desp13; }
  } else if (transporte) {
    p100 = extra;
  } else if (noct > 0) {
    p50 = 0;
    p100 = Math.max(extra, noct);
  } else {
    p50 = extra;
  }

  return { total, norm, p50, p100, noct, descanso, desc100, cenas: cena(f.entrada, f.salida, f.fecha) };
}

function minutosNocturnos(ini, fin) {
  let n = 0, t = new Date(ini);
  while (t < fin) {
    const h = t.getHours() + t.getMinutes() / 60;
    if (h >= 21 || h < 6) n++;
    t.setMinutes(t.getMinutes() + 1);
  }
  return n;
}

function minSabDespues13(ini, fin) {
  let n = 0, t = new Date(ini);
  while (t < fin) {
    if (t.getDay() === 6 && t.getHours() >= 13) n++;
    t.setMinutes(t.getMinutes() + 1);
  }
  return n;
}

function cena(entrada, salida, fecha) {
  const ini = combinar(fecha, entrada);
  const fin = combinar(fecha, salida);
  if (fin <= ini) fin.setDate(fin.getDate() + 1);
  let t = new Date(ini);
  while (t < fin) {
    if (t.getHours() >= 21 || t.getHours() < 6) return 1;
    t.setMinutes(t.getMinutes() + 1);
  }
  return 0;
}

function finTurno(fecha, entrada, salida) {
  const fin = combinar(fecha, salida);
  const ini = combinar(fecha, entrada);
  if (fin <= ini) fin.setDate(fin.getDate() + 1);
  return fin;
}

function combinar(fecha, hora) {
  const d = new Date(fecha);
  if (hora instanceof Date) {
    d.setHours(hora.getHours(), hora.getMinutes(), 0, 0);
  } else if (typeof hora === 'string') {
    const p = hora.split(':');
    d.setHours(+p[0], +p[1], 0, 0);
  }
  return d;
}

function parseFecha(v) {
  if (v instanceof Date) return v;
  const s = String(v).trim();
  const m = s.match(/(\d{1,2})[\/\-](\d{1,2})[\/\-](\d{4})/);
  if (m) return new Date(+m[3], +m[2] - 1, +m[1]);
  return null;
}

function parseHora(v) {
  if (!v && v !== 0) return null;
  if (v instanceof Date) return v;
  const s = String(v).trim().toLowerCase();
  if (s === 'si' || s === 'fco' || s === '') return null;
  const m = s.match(/(\d{1,2}):(\d{2})/);
  if (m) return m[1].padStart(2, '0') + ':' + m[2];
  return null;
}

function esSi(v) {
  return ['si', 'sí', 's', '1', 'true', 'fco', 'franco'].indexOf(String(v).trim().toLowerCase()) >= 0;
}

function fmt(min) {
  if (!min || min <= 0) return '00:00';
  const h = Math.floor(min / 60);
  const m = min % 60;
  return (h < 10 ? '0' : '') + h + ':' + (m < 10 ? '0' : '') + m;
}

function formatFecha(d) {
  const dd = ('0' + d.getDate()).slice(-2);
  const mm = ('0' + (d.getMonth() + 1)).slice(-2);
  return dd + '/' + mm + '/' + d.getFullYear();
}
