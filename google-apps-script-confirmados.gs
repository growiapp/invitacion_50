const SHEET_NAME = 'Hoja 1'; // pestaña de mensajes
const RSVP_SHEET = 'RSVPs';  // pestaña de confirmaciones

function doGet(e) {
  const action = e && e.parameter && e.parameter.action;

  if (action === 'getMensajes') {
    return getMensajes();
  }

  if (action === 'getRSVP') {
    return getRSVP();
  }

  return jsonOutput({ status: 'ok' });
}

function doPost(e) {
  const data = JSON.parse(e.postData.contents || '{}');

  if (data.action === 'addMensaje') {
    addMensaje(data.nombre, data.mensaje);
  }

  if (data.action === 'addRSVP') {
    addRSVP(data.nombre, data.cantidad, data.nota);
  }

  return jsonOutput({ status: 'ok' });
}

function getMensajes() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_NAME);
  const rows = sheet.getDataRange().getValues();
  const result = [];

  for (let i = rows.length - 1; i >= 1; i--) {
    if (rows[i][1] && rows[i][2]) {
      result.push({
        nombre: String(rows[i][1]).trim(),
        mensaje: String(rows[i][2]).trim()
      });
    }
  }

  return jsonOutput(result);
}

function addMensaje(nombre, mensaje) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_NAME);
  sheet.appendRow([new Date(), nombre || '', mensaje || '']);
}

function getRSVP() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(RSVP_SHEET);
  const rows = sheet.getDataRange().getValues();
  const result = [];

  for (let i = rows.length - 1; i >= 1; i--) {
    const nombre = rows[i][1];
    const cantidad = rows[i][2];
    const nota = rows[i][3];

    if (nombre) {
      result.push({
        nombre: String(nombre).trim(),
        cantidad: Math.max(1, parseInt(cantidad, 10) || 1),
        nota: nota ? String(nota).trim() : ''
      });
    }
  }

  return jsonOutput(result);
}

function addRSVP(nombre, cantidad, nota) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(RSVP_SHEET);
  sheet.appendRow([
    new Date(),
    nombre || '',
    Math.max(1, parseInt(cantidad, 10) || 1),
    nota || ''
  ]);
}

function jsonOutput(data) {
  return ContentService
    .createTextOutput(JSON.stringify(data))
    .setMimeType(ContentService.MimeType.JSON);
}
