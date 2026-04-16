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
// ══════════════════════════════════════════════════════

const SECRET_TOKEN = ''; // ← SKIFT DETTE til dit eget hemmelige token!
const SHEET_NAME   = 'Tankninger';

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
        'Note', 'ISO Timestamp', 'Entry ID'
      ]);
      const header = sheet.getRange(1, 1, 1, 9);
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
    }

    // Prevent duplicate entries: skip if entryId already exists in the sheet
    if (data.entryId && sheet.getLastRow() > 1) {
      const existingIds = sheet.getRange(2, 9, sheet.getLastRow() - 1, 1).getValues().flat();
      if (existingIds.includes(data.entryId)) {
        return ContentService
          .createTextOutput(JSON.stringify({ status: 'ok', message: 'duplicate' }))
          .setMimeType(ContentService.MimeType.JSON);
      }
    }

    sheet.appendRow([
      formattedDate,
      formattedTime,
      data.odometer      || '',
      data.liters        || '',
      data.pricePerLiter || '',
      data.totalPrice    || '',
      data.note          || '',
      data.timestamp,
      data.entryId       || ''
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
        const rows = sheet.getRange(2, 1, sheet.getLastRow() - 1, 9).getValues();
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
