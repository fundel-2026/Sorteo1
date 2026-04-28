// ══════════════════════════════════════════════════════════
//  SORTEO APP — Google Apps Script Backend
//  INSTRUCCIONES:
//  1. En tu Google Sheet ve a Extensiones → Apps Script
//  2. Borra el código anterior y pega este completo
//  3. Clic en "Implementar" → "Administrar implementaciones"
//     → Edita la implementación existente → Nueva versión → Implementar
//  (Si es la primera vez: Implementar → Nueva implementación →
//   Tipo: App web · Ejecutar como: Yo · Acceso: Cualquier persona)
// ══════════════════════════════════════════════════════════

const HOJA = 'Participantes';

function iniciarHoja_() {
  const ss  = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(HOJA);
  if (!sheet) {
    sheet = ss.insertSheet(HOJA);
    sheet.appendRow(['Fecha', 'Nombre', 'Email', 'Teléfono', 'Curso', 'Ticket', 'ID']);
    sheet.getRange(1, 1, 1, 7).setFontWeight('bold')
         .setBackground('#7C3AED').setFontColor('#ffffff');
  }
  return sheet;
}

/* ── Todo entra por GET para evitar problemas de CORS ── */
function doGet(e) {
  const action = (e.parameter.action || 'list');

  if (action === 'add') {
    return agregarParticipante_(e.parameter);
  }
  return listarParticipantes_();
}

/* ── Listar participantes ── */
function listarParticipantes_() {
  try {
    const sheet   = iniciarHoja_();
    const lastRow = sheet.getLastRow();
    let participants = [];

    if (lastRow > 1) {
      const data = sheet.getRange(2, 1, lastRow - 1, 7).getValues();
      participants = data
        .filter(r => r[2])
        .map(r => ({
          fecha:    r[0] instanceof Date ? r[0].toISOString() : String(r[0]),
          nombre:   String(r[1]),
          email:    String(r[2]),
          telefono: String(r[3]),
          curso:    String(r[4]),
          ticket:   String(r[5]),
          id:       String(r[6])
        }));
    }

    return out_({ ok: true, participants });
  } catch(err) {
    return out_({ ok: false, error: err.message });
  }
}

/* ── Agregar participante ── */
function agregarParticipante_(p) {
  try {
    const sheet   = iniciarHoja_();
    const lastRow = sheet.getLastRow();

    // Verificar email duplicado
    if (lastRow > 1) {
      const emails  = sheet.getRange(2, 3, lastRow - 1, 1).getValues().flat()
                           .map(x => String(x).toLowerCase().trim());
      const tickets = sheet.getRange(2, 5, lastRow - 1, 1).getValues().flat();
      const idx     = emails.indexOf(p.email.toLowerCase().trim());
      if (idx !== -1) {
        return out_({ ok: false, duplicate: true, ticket: tickets[idx] });
      }
    }

    sheet.appendRow([
      new Date().toISOString(),
      p.nombre    || '',
      p.email     || '',
      p.telefono  || '',
      p.curso     || '',
      p.ticket    || '',
      p.id        || ''
    ]);

    return out_({ ok: true, ticket: p.ticket });
  } catch(err) {
    return out_({ ok: false, error: err.message });
  }
}

function out_(obj) {
  return ContentService
    .createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}
