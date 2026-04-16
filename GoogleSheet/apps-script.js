// ══════════════════════════════════════════════════════
// VINGATEFUEL — Google Apps Script Backend
// ══════════════════════════════════════════════════════
//
// OPSÆTNING:
// 1. Opret et nyt Google Sheet
// 2. Gå til Udvidelser → Apps Script
// 3. Indsæt HELE denne fil
// 4. Skift SECRET_TOKEN til dit eget hemmelige ord
// 5. Klik Deploy → Ny implementering → Webapp
// 6. Kør som: Dig selv
//    Adgang: Alle (du er den eneste der kender URL'en)
// 7. Kopiér URL og indsæt i appen under Indstillinger
//
// Kolonner:
//   A  Dato              (1)
//   B  Tidspunkt         (2)
//   C  Odometer (miles)  (3)
//   D  Liter             (4)
//   E  Pris pr. liter    (5)
//   F  Total pris        (6)
//   G  Note              (7)
//   H  ISO Timestamp     (8)
//   I  Entry ID          (9)
//   J  Udeladt           (10)  ← ny
//   K  Slettet           (11)  ← ny
//   L  Sidst opdateret   (12)  ← ny
// ══════════════════════════════════════════════════════

const SECRET_TOKEN = ''; // ← SKIFT DETTE til dit eget hemmelige token!
const SHEET_NAME   = 'Tankninger';

// Ensure the sheet has all 12 header columns (upgrades older sheets)
function ensureHeaders(sheet) {
  const lastCol = sheet.getLastColumn();
  if (lastCol < 10) {
    sheet.getRange(1, 10).setValue('Udeladt');
    sheet.getRange(1, 10).setFontWeight('bold').setBackground('#1a2a38').setFontColor('#f0a040');
    sheet.setColumnWidth(10, 80);
  }
  if (lastCol < 11) {
    sheet.getRange(1, 11).setValue('Slettet');
    sheet.getRange(1, 11).setFontWeight('bold').setBackground('#1a2a38').setFontColor('#f0a040');
    sheet.setColumnWidth(11, 80);
  }
  if (lastCol < 12) {
    sheet.getRange(1, 12).setValue('Sidst opdateret');
    sheet.getRange(1, 12).setFontWeight('bold').setBackground('#1a2a38').setFontColor('#f0a040');
    sheet.setColumnWidth(12, 180);
  }
}

function doPost(e) {
  try {
    const data = JSON.parse(e.postData.contents);

    if (SECRET_TOKEN !== '' && data.secret !== SECRET_TOKEN) {
      return ContentService
        .createTextOutput(JSON.stringify({ status: 'error', message: 'Ugyldig nøgle' }))
        .setMimeType(ContentService.MimeType.JSON);
    }

    const ss        = SpreadsheetApp.getActiveSpreadsheet();
    const timestamp = new Date(data.timestamp);
    const formattedDate = Utilities.formatDate(timestamp, Session.getScriptTimeZone(), 'dd-MM-yyyy');
    const formattedTime = Utilities.formatDate(timestamp, Session.getScriptTimeZone(), 'HH:mm');

    let sheet = ss.getSheetByName(SHEET_NAME);
    if (!sheet) {
      sheet = ss.insertSheet(SHEET_NAME);
      sheet.appendRow([
        'Dato', 'Tidspunkt', 'Odometer (miles)',
        'Liter', 'Pris pr. liter (kr.)', 'Total pris (kr.)',
        'Note', 'ISO Timestamp', 'Entry ID',
        'Udeladt', 'Slettet', 'Sidst opdateret'
      ]);
      const header = sheet.getRange(1, 1, 1, 12);
      header.setFontWeight('bold').setBackground('#1a2a38').setFontColor('#f0a040');
      sheet.setFrozenRows(1);
      sheet.setColumnWidth(1, 100);
      sheet.setColumnWidth(2, 80);
      sheet.setColumnWidth(3, 140);
      sheet.setColumnWidth(4, 80);
      sheet.setColumnWidth(5, 140);
      sheet.setColumnWidth(6, 130);
      sheet.setColumnWidth(7, 250);
      sheet.setColumnWidth(8, 180);
      sheet.setColumnWidth(9, 220);
      sheet.setColumnWidth(10, 80);
      sheet.setColumnWidth(11, 80);
      sheet.setColumnWidth(12, 180);
    } else {
      ensureHeaders(sheet);
    }

    const incomingUpdatedAt = data.updatedAt || data.timestamp || '';

    // Check if entryId already exists → update metadata if incoming is newer
    if (data.entryId && sheet.getLastRow() > 1) {
      const lastRow      = sheet.getLastRow();
      const idValues     = sheet.getRange(2, 9, lastRow - 1, 1).getValues().flat();
      const existingIdx  = idValues.findIndex(id => id === data.entryId);

      if (existingIdx !== -1) {
        const sheetRow        = existingIdx + 2; // 1-indexed + header
        const existingUpdated = sheet.getRange(sheetRow, 12).getValue();
        const shouldUpdate    = incomingUpdatedAt && (
          !existingUpdated ||
          (!isNaN(new Date(incomingUpdatedAt)) && new Date(incomingUpdatedAt) > new Date(existingUpdated))
        );

        if (shouldUpdate) {
          sheet.getRange(sheetRow, 10).setValue(data.excluded  ? 'Ja' : 'Nej');
          sheet.getRange(sheetRow, 11).setValue(data.isDeleted ? 'Ja' : 'Nej');
          sheet.getRange(sheetRow, 12).setValue(incomingUpdatedAt);
        }

        return ContentService
          .createTextOutput(JSON.stringify({ status: 'ok', message: 'updated' }))
          .setMimeType(ContentService.MimeType.JSON);
      }
    }

    // New entry — append row
    sheet.appendRow([
      formattedDate,
      formattedTime,
      data.odometer      || '',
      data.liters        || '',
      data.pricePerLiter || '',
      data.totalPrice    || '',
      data.note          || '',
      data.timestamp,
      data.entryId       || '',
      data.excluded  ? 'Ja' : 'Nej',
      data.isDeleted ? 'Ja' : 'Nej',
      incomingUpdatedAt
    ]);

    return ContentService
      .createTextOutput(JSON.stringify({ status: 'ok' }))
      .setMimeType(ContentService.MimeType.JSON);

  } catch (err) {
    return ContentService
      .createTextOutput(JSON.stringify({ status: 'error', message: err.toString() }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

function doGet(e) {
  try {
    const action = (e && e.parameter && e.parameter.action) ? e.parameter.action : '';
    const secret = (e && e.parameter && e.parameter.secret) ? e.parameter.secret : '';

    if (SECRET_TOKEN !== '' && secret !== SECRET_TOKEN) {
      return ContentService
        .createTextOutput(JSON.stringify({ status: 'error', message: 'Ugyldig nøgle' }))
        .setMimeType(ContentService.MimeType.JSON);
    }

    if (action === 'list') {
      const ss    = SpreadsheetApp.getActiveSpreadsheet();
      const sheet = ss.getSheetByName(SHEET_NAME);
      const entries = [];

      if (sheet && sheet.getLastRow() > 1) {
        const dataLastCol = sheet.getLastColumn();
        const rows = sheet.getRange(2, 1, sheet.getLastRow() - 1, dataLastCol).getValues();
        rows.forEach(row => {
          if (!row[7]) return; // skip if no timestamp
          entries.push({
            id:            row[8] || ('fuel:' + row[7]),
            timestamp:     row[7],
            odometer:      Number(row[2]) || 0,
            liters:        Number(row[3]) || 0,
            pricePerLiter: row[4] !== '' ? Number(row[4]) : null,
            totalPrice:    row[5] !== '' ? Number(row[5]) : null,
            note:          row[6] || '',
            excluded:      (row.length > 9  ? row[9]  : '') === 'Ja',
            isDeleted:     (row.length > 10 ? row[10] : '') === 'Ja',
            updatedAt:     row.length > 11  ? (row[11] || '') : '',
            synced:        true
          });
        });
      }

      entries.sort((a, b) => b.odometer - a.odometer);

      return ContentService
        .createTextOutput(JSON.stringify({ status: 'ok', entries: entries }))
        .setMimeType(ContentService.MimeType.JSON);
    }

    // ping / default
    return ContentService
      .createTextOutput(JSON.stringify({ status: 'ok', message: 'VingateFuel API kører ✓' }))
      .setMimeType(ContentService.MimeType.JSON);

  } catch (err) {
    return ContentService
      .createTextOutput(JSON.stringify({ status: 'error', message: err.toString() }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}
