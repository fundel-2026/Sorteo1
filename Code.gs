// ══════════════════════════════════════════════════════════
//  SORTEO APP — Google Apps Script Backend
//  Instrucciones:
//  1. Ve a sheets.new → crea una hoja nueva
//  2. Menú: Extensiones → Apps Script
//  3. Borra el código que hay y pega TODO este archivo
//  4. Clic en "Implementar" → "Nueva implementación"
//     · Tipo: Aplicación web
//     · Ejecutar como: Yo
//     · Quién tiene acceso: Cualquier persona
//  5. Copia la URL que aparece y pégala en index.html y admin.html
//     donde dice: const ENDPOINT = '...';
// ══════════════════════════════════════════════════════════

const HOJA = 'Participantes';

function iniciarHoja_() {
  const ss    = SpreadsheetApp.getActiveSpreadsheet();
  let sheet   = ss.getSheetByName(HOJA);
  if (!sheet) {
    sheet = ss.insertSheet(HOJA);
    sheet.appendRow(['Fecha', 'Nombre', 'Email', 'Teléfono', 'Ticket', 'ID']);
    sheet.getRange(1, 1, 1, 6).setFontWeight('bold').setBackground('#7C3AED').setFontColor('#ffffff');
    sheet.setColumnWidth(1, 180);
    sheet.setColumnWidth(2, 200);
    sheet.setColumnWidth(3, 220);
  }
  return sheet;
}

/* ── GET: devuelve todos los participantes ── */
function doGet(e) {
  try {
    const sheet   = iniciarHoja_();
    const lastRow = sheet.getLastRow();
    let participants = [];

    if (lastRow > 1) {
      const data = sheet.getRange(2, 1, lastRow - 1, 6).getValues();
      participants = data
        .filter(r => r[2])
        .map(r => ({
          fecha:    r[0] instanceof Date ? r[0].toISOString() : String(r[0]),
          nombre:   String(r[1]),
          email:    String(r[2]),
          telefono: String(r[3]),
          ticket:   String(r[4]),
          id:       String(r[5])
        }));
    }

    return ContentService
      .createTextOutput(JSON.stringify({ ok: true, participants }))
      .setMimeType(ContentService.MimeType.JSON);

  } catch (err) {
    return ContentService
      .createTextOutput(JSON.stringify({ ok: false, error: err.message }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

/* ── POST: agrega un participante ── */
function doPost(e) {
  try {
    const sheet = iniciarHoja_();
    const data  = JSON.parse(e.postData.contents);

    // Verificar duplicado por email
    const lastRow = sheet.getLastRow();
    if (lastRow > 1) {
      const emails = sheet.getRange(2, 3, lastRow - 1, 1).getValues().flat()
                         .map(x => String(x).toLowerCase().trim());
      if (emails.includes(data.email.toLowerCase().trim())) {
        const idx     = emails.indexOf(data.email.toLowerCase().trim());
        const tickets = sheet.getRange(2, 5, lastRow - 1, 1).getValues().flat();
        return ContentService
          .createTextOutput(JSON.stringify({ ok: false, duplicate: true, ticket: tickets[idx] }))
          .setMimeType(ContentService.MimeType.JSON);
      }
    }

    sheet.appendRow([
      new Date().toISOString(),
      data.nombre,
      data.email,
      data.telefono || '',
      data.ticket,
      data.id
    ]);

    return ContentService
      .createTextOutput(JSON.stringify({ ok: true, ticket: data.ticket }))
      .setMimeType(ContentService.MimeType.JSON);

  } catch (err) {
    return ContentService
      .createTextOutput(JSON.stringify({ ok: false, error: err.message }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}
