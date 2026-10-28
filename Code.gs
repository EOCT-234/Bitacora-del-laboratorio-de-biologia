/**
 * Receptor de la Bitácora del Laboratorio de Biología (UNTRM).
 
 */
const SHEET_ID = '10IJwN5wLY_DscbgDeU_8pv-Jbqpw4u3YARSXPl0nun8';
const HOJA = 'Registros';
const COLS = ['ID','Fecha','Hora de ingreso','Responsable','Personal presente','Equipos','Otros equipos',
  'Materiales','Otros materiales','Áreas','Actividades','Foto ingreso','Hora de salida','Verificación de cierre',
  'Foto salida','Observaciones','Registrado'];

function doGet() { return salida({ ok: true, mensaje: 'Bitácora activa' }); }

function doPost(e) {
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(25000);
    const d = JSON.parse(e.postData.contents);
    const sh = hoja();
    const ultima = sh.getLastRow();
    const ids = ultima > 1 ? sh.getRange(2, 1, ultima - 1, 1).getValues().flat().map(String) : [];
    const i = ids.indexOf(String(d.id));

    if (d.accion === 'ingreso') {
      if (i >= 0) return salida({ ok: true, repetido: true }); // reintento: no duplicar
      const f = ultima + 1;
      const r = sh.getRange(f, 1, 1, COLS.length);
      r.setNumberFormat('@'); // evita que Sheets cambie fechas y horas
      r.setValues([[d.id, d.fecha, d.hora_ingreso, d.responsable, d.personal, d.equipos, d.otros_equipos,
        d.materiales, d.otros_materiales, d.areas, d.actividades, guardarFoto(d.foto, d.id + '_ingreso'),
        '', '', '', '', Utilities.formatDate(new Date(), 'America/Lima', 'yyyy-MM-dd HH:mm:ss')]]);
    } else if (d.accion === 'salida') {
      let f = i + 2;
      if (i < 0) { // ingreso no encontrado: se crea la fila para no perder la salida
        f = ultima + 1;
        sh.getRange(f, 1, 1, COLS.length).setNumberFormat('@');
        sh.getRange(f, 1, 1, 4).setValues([[d.id, '', '', '(ingreso no encontrado)']]);
      }
      const r = sh.getRange(f, 13, 1, 4);
      r.setNumberFormat('@');
      r.setValues([[d.hora_salida, d.cierre, guardarFoto(d.foto, d.id + '_salida'), d.observaciones]]);
    } else throw new Error('Acción no válida');
    return salida({ ok: true });
  } catch (err) {
    return salida({ ok: false, error: String(err && err.message || err) });
  } finally { try { lock.releaseLock(); } catch (x) {} }
}

function hoja() {
  const ss = SpreadsheetApp.openById(SHEET_ID);
  let sh = ss.getSheetByName(HOJA);
  if (!sh) sh = ss.insertSheet(HOJA);
  if (sh.getLastRow() === 0) {
    sh.appendRow(COLS);
    sh.getRange(1, 1, 1, COLS.length).setFontWeight('bold').setBackground('#0b2a5b').setFontColor('#ffffff');
    sh.setFrozenRows(1);
  }
  return sh;
}

function guardarFoto(b64, nombre) {
  if (!b64) return '';
  try { // si la foto falla, el registro igual se guarda
    const it = DriveApp.getFoldersByName('Bitacora_Fotos');
    const carpeta = it.hasNext() ? it.next() : DriveApp.createFolder('Bitacora_Fotos');
    const blob = Utilities.newBlob(Utilities.base64Decode(b64.split(',')[1]), 'image/jpeg', nombre + '.jpg');
    return carpeta.createFile(blob).getUrl();
  } catch (err) { return '(foto no guardada)'; }
}

function salida(o) {
  return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON);
}
