/**
 * Horas Transporte — Google Drive (carpeta compartida)
 *
 * Estructura en Mi unidad:
 *   Horas Transporte/
 *     datos.json          ← archivo principal (siempre actualizado)
 *     Respaldos/
 *       respaldo_2026-07-13_14-30.json
 *       respaldo_2026-07-13_18-45.json
 *       ...
 *
 * Compartir la carpeta + la URL de la app → otra persona usa los mismos datos.
 */

var CARPETA_NOMBRE = 'Horas Transporte';
var SUB_RESPALDOS = 'Respaldos';
var ARCHIVO_DATOS = 'datos.json';
var ARCHIVO_VIEJO = 'horas_transporte_nube.json';
var MAX_RESPALDOS = 30;

function doGet(e) {
  if (e && e.parameter && e.parameter.action === 'load') {
    return ContentService.createTextOutput(getDatosNube())
      .setMimeType(ContentService.MimeType.JSON);
  }
  return HtmlService.createHtmlOutputFromFile('ControlHoras')
    .setTitle('Horas Transporte')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
    .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

function doPost(e) {
  try {
    var body = e.postData && e.postData.contents ? e.postData.contents : '{}';
    saveDatosNube(body);
    return ContentService.createTextOutput(JSON.stringify({ ok: true }))
      .setMimeType(ContentService.MimeType.JSON);
  } catch (err) {
    return ContentService.createTextOutput(JSON.stringify({ ok: false, error: String(err) }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

function cargarNube() {
  return getDatosNube();
}

function guardarNube(jsonString) {
  saveDatosNube(jsonString);
  return getInfoCarpeta();
}

/** Ejecutar desde el editor: seleccionar probarDrive → Ejecutar */
function probarDrive() {
  var info = getInfoCarpeta();
  saveDatosNube(JSON.stringify({
    datos: {},
    personal: [],
    exportado: new Date().toISOString(),
    prueba: true
  }));
  Logger.log('OK carpeta: ' + info.url);
  return 'Carpeta creada: ' + info.url;
}

/** Info para mostrar en la planilla (nombre y link de la carpeta) */
function getInfoCarpeta() {
  var carpeta = getCarpetaPrincipal();
  var respaldos = getSubcarpetaRespaldos(carpeta);
  return {
    carpeta: carpeta.getName(),
    url: carpeta.getUrl(),
    archivo: ARCHIVO_DATOS,
    respaldos: respaldos.getName()
  };
}

function getCarpetaPrincipal() {
  var props = PropertiesService.getScriptProperties();
  var id = props.getProperty('CARPETA_HORAS_ID');
  if (id) {
    try {
      return DriveApp.getFolderById(id);
    } catch (e) {
      props.deleteProperty('CARPETA_HORAS_ID');
    }
  }
  var carpetas = DriveApp.getFoldersByName(CARPETA_NOMBRE);
  if (carpetas.hasNext()) {
    var c = carpetas.next();
    props.setProperty('CARPETA_HORAS_ID', c.getId());
    return c;
  }
  var nueva = DriveApp.createFolder(CARPETA_NOMBRE);
  props.setProperty('CARPETA_HORAS_ID', nueva.getId());
  return nueva;
}

function getSubcarpetaRespaldos(carpeta) {
  var subs = carpeta.getFoldersByName(SUB_RESPALDOS);
  if (subs.hasNext()) return subs.next();
  return carpeta.createFolder(SUB_RESPALDOS);
}

function buscarArchivoEnCarpeta(carpeta, nombre) {
  var files = carpeta.getFilesByName(nombre);
  return files.hasNext() ? files.next() : null;
}

function migrarDesdeRaizSiHaceFalta(carpeta) {
  var destino = buscarArchivoEnCarpeta(carpeta, ARCHIVO_DATOS);
  if (destino) return;

  var enCarpetaViejo = buscarArchivoEnCarpeta(carpeta, ARCHIVO_VIEJO);
  if (enCarpetaViejo) {
    enCarpetaViejo.setName(ARCHIVO_DATOS);
    return;
  }

  var viejos = DriveApp.getFilesByName(ARCHIVO_VIEJO);
  if (viejos.hasNext()) {
    var f = viejos.next();
    f.setName(ARCHIVO_DATOS);
    f.moveTo(carpeta);
  }
}

function getDatosNube() {
  var carpeta = getCarpetaPrincipal();
  migrarDesdeRaizSiHaceFalta(carpeta);
  var archivo = buscarArchivoEnCarpeta(carpeta, ARCHIVO_DATOS);
  if (!archivo) {
    return JSON.stringify({ datos: {}, personal: [], exportado: null });
  }
  return archivo.getBlob().getDataAsString() || JSON.stringify({ datos: {}, personal: [] });
}

function saveDatosNube(jsonString) {
  var carpeta = getCarpetaPrincipal();
  migrarDesdeRaizSiHaceFalta(carpeta);
  var archivo = buscarArchivoEnCarpeta(carpeta, ARCHIVO_DATOS);
  if (archivo) {
    archivo.setContent(jsonString);
  } else {
    carpeta.createFile(ARCHIVO_DATOS, jsonString, MimeType.PLAIN_TEXT);
  }
  guardarRespaldo(carpeta, jsonString);
}

function guardarRespaldo(carpeta, jsonString) {
  var sub = getSubcarpetaRespaldos(carpeta);
  var ahora = Utilities.formatDate(new Date(), Session.getScriptTimeZone() || 'America/Argentina/Buenos_Aires', 'yyyy-MM-dd_HH-mm');
  sub.createFile('respaldo_' + ahora + '.json', jsonString, MimeType.PLAIN_TEXT);
  limpiarRespaldosViejos(sub);
}

function limpiarRespaldosViejos(carpetaRespaldos) {
  var files = carpetaRespaldos.getFiles();
  var lista = [];
  while (files.hasNext()) {
    var f = files.next();
    if (String(f.getName()).indexOf('respaldo_') === 0) lista.push(f);
  }
  lista.sort(function (a, b) {
    return b.getLastUpdated().getTime() - a.getLastUpdated().getTime();
  });
  for (var i = MAX_RESPALDOS; i < lista.length; i++) {
    lista[i].setTrashed(true);
  }
}
