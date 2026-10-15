/**
 * Contador de eventos - guardado en Google Sheets
 * Pegar en la hoja: Extensiones > Apps Script
 * Después: Implementar > Nueva implementación > Aplicación web
 *   Ejecutar como: Yo
 *   Quién tiene acceso: Cualquier persona
 * La hoja puede estar vacía: las pestañas "Historial" y "Totales" se crean solas.
 *
 * Leer (doGet) es público: cualquiera con el enlace de la página ve el progreso.
 * Escribir (doPost) exige la CLAVE: solo quien la conoce puede registrar, editar o borrar.
 */

// CAMBIA ESTA CLAVE por una frase tuya. No la pongas en la página: se escribe en el navegador de quien registra.
var CLAVE = 'cambia-esta-clave';
var HOJA = 'Historial';
var VERSION = 4; // abre la URL /exec en el navegador: si no dice "version":4, falta publicar una versión nueva
var COLS = 6;
var ENCABEZADOS = [['Fecha', 'Hora', 'Evento', 'Cantidad', 'ID (no borrar)', 'Marca de tiempo (no borrar)']];
var FORMATOS = [['dd/mm/yyyy', 'hh:mm:ss', 'General', '0', '@', '0']];

function salida_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

function hoja_() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sh = ss.getSheetByName(HOJA);
  if (!sh) {
    sh = ss.insertSheet(HOJA);
    sh.setFrozenRows(1);
    sh.setColumnWidth(3, 240);
  }
  if (sh.getRange(1, COLS).getValue() === '') {   // hoja nueva o de una versión anterior con 5 columnas
    sh.getRange(1, 1, 1, COLS).setValues(ENCABEZADOS).setFontWeight('bold');
  }
  if (!ss.getSheetByName('Totales')) {
    var t = ss.insertSheet('Totales');
    t.getRange('A1:B1').setValues([['Evento', 'Total']]).setFontWeight('bold');
    t.getRange('A2').setFormula('=IFERROR(UNIQUE(FILTER(Historial!C2:C,Historial!C2:C<>"")),"")');
    t.getRange('B2').setFormula('=ARRAYFORMULA(IF(A2:A="","",SUMIF(Historial!C:C,A2:A,Historial!D:D)))');
    t.setColumnWidth(1, 240);
  }
  return sh;
}

// La página envía fecha y hora ya calculadas con la hora del dispositivo, así no depende de la zona horaria de la hoja.
function fila_(r) {
  var fecha = Number(r.fecha), hora = Number(r.hora);
  if (!(fecha > 0) || !(hora >= 0)) {   // envío sin fecha: se calcula con la hora de Colombia (UTC-5)
    var local = Number(r.ts) - 5 * 3600000;
    fecha = Math.floor(local / 86400000) + 25569;
    hora = (local % 86400000) / 86400000;
  }
  var nombre = String(r.nombre).replace(/^([=+\-@])/, "'$1"); // evita que un nombre se lea como fórmula
  return [fecha, hora, nombre, Number(r.n), String(r.id), Number(r.ts)];
}

// La página envía cambios sueltos: { op: 'put', id, nombre, n, ts, fecha, hora } o { op: 'del', id }.
// 'put' agrega la fila o la actualiza si ese ID ya existe; 'del' borra la fila de ese ID.
// Con ops vacío sirve para comprobar la clave.
function doPost(e) {
  var lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    var d = JSON.parse(e.postData.contents);
    if (d.token !== CLAVE) return salida_({ ok: false, error: 'Clave incorrecta' });
    if (d.action !== 'ops' || !Array.isArray(d.ops)) return salida_({ ok: false, error: 'Acción no válida' });

    var sh = hoja_();
    var ultima = sh.getLastRow();
    var filaDe = {};
    if (ultima > 1) {
      sh.getRange(2, 5, ultima - 1, 1).getValues().forEach(function (r, i) {
        if (r[0] !== '') filaDe[String(r[0])] = i + 2;
      });
    }

    d.ops.forEach(function (op) {
      if (!op || !op.id) return;
      var id = String(op.id), fila = filaDe[id];
      if (op.op === 'put') {
        if (!fila) { fila = sh.getLastRow() + 1; filaDe[id] = fila; }
        sh.getRange(fila, 1, 1, COLS).setValues([fila_(op)]).setNumberFormats(FORMATOS);
      } else if (op.op === 'del' && fila) {
        sh.deleteRow(fila);
        delete filaDe[id];
        for (var k in filaDe) if (filaDe[k] > fila) filaDe[k]--;
      }
    });
    return salida_({ ok: true, count: d.ops.length, version: VERSION });
  } catch (err) {
    return salida_({ ok: false, error: String(err) });
  } finally {
    lock.releaseLock();
  }
}

// Público: devuelve todos los registros para que la página muestre el progreso.
function doGet() {
  try {
    var sh = hoja_();
    var ultima = sh.getLastRow();
    var rows = [];
    if (ultima > 1) {
      sh.getRange(2, 1, ultima - 1, COLS).getValues().forEach(function (r) {
        var id = String(r[4]), n = Number(r[3]), ts = Number(r[5]);
        if (!ts && r[0] instanceof Date) ts = r[0].getTime();   // fila de una versión anterior sin marca de tiempo
        if (id && r[2] !== '' && n > 0 && ts > 0) rows.push({ id: id, nombre: String(r[2]).replace(/^'/, ''), n: n, ts: ts });
      });
    }
    return salida_({ ok: true, version: VERSION, rows: rows });
  } catch (err) {
    return salida_({ ok: false, error: String(err) });
  }
}
