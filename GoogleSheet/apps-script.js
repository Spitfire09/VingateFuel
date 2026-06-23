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

const SECRET_TOKEN     = ''; // ← SKIFT DETTE til dit eget hemmelige token!
const SHEET_NAME       = 'Tankninger';
const CHECK_SHEET_NAME = 'ServiceTjek';

function ensureFuelHeaders(sheet) {
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

function ensureCheckHeaders(sheet) {
  const lastCol = sheet.getLastColumn();
  if (lastCol < 8) {
    sheet.getRange(1, 8).setValue('Note');
    sheet.getRange(1, 8).setFontWeight('bold').setBackground('#1a2a38').setFontColor('#f0a040');
    sheet.setColumnWidth(8, 250);
  }
  if (lastCol < 9) {
    sheet.getRange(1, 9).setValue('Slettet');
    sheet.getRange(1, 9).setFontWeight('bold').setBackground('#1a2a38').setFontColor('#f0a040');
    sheet.setColumnWidth(9, 80);
  }
  if (lastCol < 10) {
    sheet.getRange(1, 10).setValue('Slettet tidspunkt');
    sheet.getRange(1, 10).setFontWeight('bold').setBackground('#1a2a38').setFontColor('#f0a040');
    sheet.setColumnWidth(10, 180);
  }
  if (lastCol < 11) {
    sheet.getRange(1, 11).setValue('Sidst opdateret');
    sheet.getRange(1, 11).setFontWeight('bold').setBackground('#1a2a38').setFontColor('#f0a040');
    sheet.setColumnWidth(11, 180);
  }
}

function getOrCreateFuelSheet(ss) {
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
    ensureFuelHeaders(sheet);
  }
  return sheet;
}

function getOrCreateCheckSheet(ss) {
  let sheet = ss.getSheetByName(CHECK_SHEET_NAME);
  if (!sheet) {
    sheet = ss.insertSheet(CHECK_SHEET_NAME);
    sheet.appendRow([
      'Dato', 'Tidspunkt', 'Odometer (miles)',
      'Type', 'Handling', 'ISO Timestamp',
      'Check ID', 'Note', 'Slettet',
      'Slettet tidspunkt', 'Sidst opdateret'
    ]);
    const header = sheet.getRange(1, 1, 1, 11);
    header.setFontWeight('bold').setBackground('#1a2a38').setFontColor('#f0a040');
    sheet.setFrozenRows(1);
    sheet.setColumnWidth(1, 100);
    sheet.setColumnWidth(2, 80);
    sheet.setColumnWidth(3, 140);
    sheet.setColumnWidth(4, 100);
    sheet.setColumnWidth(5, 110);
    sheet.setColumnWidth(6, 180);
    sheet.setColumnWidth(7, 220);
    sheet.setColumnWidth(8, 250);
    sheet.setColumnWidth(9, 80);
    sheet.setColumnWidth(10, 180);
    sheet.setColumnWidth(11, 180);
  } else {
    ensureCheckHeaders(sheet);
  }
  return sheet;
}

function doPost(e) {
  try {
    const data = JSON.parse(e.postData.contents);

    if (SECRET_TOKEN !== '' && data.secret !== SECRET_TOKEN) {
      return ContentService
        .createTextOutput(JSON.stringify({ status: 'error', message: 'Ugyldig nøgle' }))
        .setMimeType(ContentService.MimeType.JSON);
    }

    const ss          = SpreadsheetApp.getActiveSpreadsheet();
    const timestamp   = new Date(data.timestamp);
    const timezone    = Session.getScriptTimeZone();
    const dateValue   = isNaN(timestamp) ? new Date() : timestamp;
    const formattedDate = Utilities.formatDate(dateValue, timezone, 'dd-MM-yyyy');
    const formattedTime = Utilities.formatDate(dateValue, timezone, 'HH:mm');
    const recordType    = data.recordType === 'check' ? 'check' : 'fuel';

    if (recordType === 'check') {
      const sheet = getOrCreateCheckSheet(ss);
      const incomingUpdatedAt = data.updatedAt || data.timestamp || '';

      if (data.checkId && sheet.getLastRow() > 1) {
        const lastRow      = sheet.getLastRow();
        const idValues     = sheet.getRange(2, 7, lastRow - 1, 1).getValues().flat();
        const existingIdx  = idValues.findIndex(id => id === data.checkId);

        if (existingIdx !== -1) {
          const sheetRow        = existingIdx + 2;
          const existingUpdated = sheet.getRange(sheetRow, 11).getValue();
          const shouldUpdate    = incomingUpdatedAt && (
            !existingUpdated ||
            (!isNaN(new Date(incomingUpdatedAt)) && new Date(incomingUpdatedAt) > new Date(existingUpdated))
          );

          if (shouldUpdate) {
            sheet.getRange(sheetRow, 1).setValue(formattedDate);
            sheet.getRange(sheetRow, 2).setValue(formattedTime);
            sheet.getRange(sheetRow, 3).setValue(data.odometer || '');
            sheet.getRange(sheetRow, 4).setValue(data.checkType || 'oil');
            sheet.getRange(sheetRow, 5).setValue(data.action || 'checked');
            sheet.getRange(sheetRow, 6).setValue(data.timestamp || '');
            sheet.getRange(sheetRow, 8).setValue(data.note || '');
            sheet.getRange(sheetRow, 9).setValue(data.isDeleted ? 'Ja' : 'Nej');
            sheet.getRange(sheetRow, 10).setValue(data.deletedAt || '');
            sheet.getRange(sheetRow, 11).setValue(incomingUpdatedAt);
          }

          return ContentService
            .createTextOutput(JSON.stringify({ status: 'ok', message: 'updated' }))
            .setMimeType(ContentService.MimeType.JSON);
        }
      }

      sheet.appendRow([
        formattedDate,
        formattedTime,
        data.odometer || '',
        data.checkType || 'oil',
        data.action || 'checked',
        data.timestamp || '',
        data.checkId || '',
        data.note || '',
        data.isDeleted ? 'Ja' : 'Nej',
        data.deletedAt || '',
        incomingUpdatedAt
      ]);

      return ContentService
        .createTextOutput(JSON.stringify({ status: 'ok' }))
        .setMimeType(ContentService.MimeType.JSON);
    }

    const sheet = getOrCreateFuelSheet(ss);
    const incomingUpdatedAt = data.updatedAt || data.timestamp || '';

    if (data.entryId && sheet.getLastRow() > 1) {
      const lastRow      = sheet.getLastRow();
      const idValues     = sheet.getRange(2, 9, lastRow - 1, 1).getValues().flat();
      const existingIdx  = idValues.findIndex(id => id === data.entryId);

      if (existingIdx !== -1) {
        const sheetRow        = existingIdx + 2;
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
      const ss         = SpreadsheetApp.getActiveSpreadsheet();
      const fuelSheet  = ss.getSheetByName(SHEET_NAME);
      const checkSheet = ss.getSheetByName(CHECK_SHEET_NAME);
      const entries = [];
      const checks  = [];

      if (fuelSheet && fuelSheet.getLastRow() > 1) {
        const dataLastCol = fuelSheet.getLastColumn();
        const rows = fuelSheet.getRange(2, 1, fuelSheet.getLastRow() - 1, dataLastCol).getValues();
        rows.forEach(row => {
          if (!row[7]) return;
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

      if (checkSheet && checkSheet.getLastRow() > 1) {
        const checkLastCol = checkSheet.getLastColumn();
        const rows = checkSheet.getRange(2, 1, checkSheet.getLastRow() - 1, checkLastCol).getValues();
        rows.forEach(row => {
          if (!row[5]) return;
          checks.push({
            id:        row[6] || ('check:' + row[5]),
            timestamp: row[5],
            odometer:  Number(row[2]) || 0,
            type:      row[3] === 'coolant' ? 'coolant' : 'oil',
            action:    row[4] === 'topped_up' ? 'topped_up' : 'checked',
            note:      row.length > 7 ? (row[7] || '') : '',
            isDeleted: (row.length > 8 ? row[8] : '') === 'Ja',
            deletedAt: row.length > 9 ? (row[9] || '') : '',
            updatedAt: row.length > 10 ? (row[10] || '') : '',
            synced:    true
          });
        });
      }

      entries.sort((a, b) => b.odometer - a.odometer);
      checks.sort((a, b) => b.odometer - a.odometer);

      return ContentService
        .createTextOutput(JSON.stringify({ status: 'ok', entries: entries, checks: checks }))
        .setMimeType(ContentService.MimeType.JSON);
    }

    return ContentService
      .createTextOutput(JSON.stringify({ status: 'ok', message: 'VingateFuel API kører ✓' }))
      .setMimeType(ContentService.MimeType.JSON);

  } catch (err) {
    return ContentService
      .createTextOutput(JSON.stringify({ status: 'error', message: err.toString() }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}
