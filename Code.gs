/**
 * Contador de eventos - guardado en Google Sheets
 * Pegar en la hoja: Extensiones > Apps Script
 * Después: Implementar > Nueva implementación > Aplicación web
 *   Ejecutar como: Yo
 *   Quién tiene acceso: Cualquier persona
 * La hoja puede estar vacía: las pestañas "Historial" y "Totales" se crean solas.
 */

// CAMBIA ESTA CLAVE por una frase tuya. La misma va en la página.
var CLAVE = 'cambia-esta-clave';
var HOJA = 'Historial';
var VERSION = 3; // abre la URL /exec en el navegador: si no dice "version": 3, falta publicar una versión nueva
var FORMATOS = [['dd/mm/yyyy', 'hh:mm:ss', 'General', '0', '@']];

function salida_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

function hoja_() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sh = ss.getSheetByName(HOJA);
  if (!sh) {
    sh = ss.insertSheet(HOJA);
    sh.getRange(1, 1, 1, 5)
      .setValues([['Fecha', 'Hora', 'Evento', 'Cantidad', 'ID (no borrar)']])
      .setFontWeight('bold');
    sh.setFrozenRows(1);
    sh.setColumnWidth(3, 240);
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
  return [fecha, hora, nombre, Number(r.n), String(r.id)];
}

// La página envía cambios sueltos: { op: 'put', id, nombre, n, ts } o { op: 'del', id }.
// 'put' agrega la fila o la actualiza si ese ID ya existe; 'del' borra la fila de ese ID.
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
        sh.getRange(fila, 1, 1, 5).setValues([fila_(op)]).setNumberFormats(FORMATOS);
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

// Solo para comprobar en el navegador que la aplicación web está activa.
function doGet() {
  return salida_({ ok: true, version: VERSION, mensaje: 'Contador de eventos: listo para recibir datos.' });
}
