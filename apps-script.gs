// Google Apps Script — in dein Google Sheet einfügen
// Extensions → Apps Script → diesen Code einfügen → Speichern → Als Web-App deployen

function doPost(e) {
  try {
    var data   = JSON.parse(e.postData.contents);
    var sheet  = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();

    // Kopfzeile beim ersten Eintrag automatisch anlegen
    if (sheet.getLastRow() === 0) {
      var headers = ['Zeitstempel', 'Name', 'E-Mail', 'Telefon', 'Berufsfeld', 'Beitrag'];
      sheet.appendRow(headers);
      sheet.getRange(1, 1, 1, headers.length)
           .setFontWeight('bold')
           .setBackground('#0A0A0A')
           .setFontColor('#FFFFFF');
      sheet.setFrozenRows(1);
    }

    sheet.appendRow([
      data.timestamp,
      data.name,
      data.email,
      data.phone,
      data.field,
      data.contribution
    ]);

    return ContentService
      .createTextOutput(JSON.stringify({ status: 'success' }))
      .setMimeType(ContentService.MimeType.JSON);

  } catch (err) {
    return ContentService
      .createTextOutput(JSON.stringify({ status: 'error', message: err.toString() }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

// Test-Funktion — direkt in Apps Script ausführen um die Verbindung zu prüfen
function test() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  Logger.log('Sheet: ' + sheet.getName() + ' | Zeilen: ' + sheet.getLastRow());
}
