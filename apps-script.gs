// Google Apps Script — in dein Google Sheet einfügen
// Extensions → Apps Script → diesen Code einfügen → Speichern → Als Web-App deployen

function doPost(e) {
  try {
    var data   = JSON.parse(e.postData.contents);
    var sheet  = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();

    // Eintrag bearbeiten
    if (data.action === 'update') {
      var lastRow = sheet.getLastRow();
      if (lastRow > 1) {
        var emails = sheet.getRange(2, 3, lastRow - 1, 1).getValues();
        for (var i = 0; i < emails.length; i++) {
          if (emails[i][0] === data.originalEmail) {
            sheet.getRange(i + 2, 2, 1, 5).setValues([[
              data.name, data.email, data.phone, data.field, data.contribution
            ]]);
            break;
          }
        }
      }
      return ContentService
        .createTextOutput(JSON.stringify({ status: 'success' }))
        .setMimeType(ContentService.MimeType.JSON);
    }

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

function doGet(e) {
  var sheet   = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  var lastRow = sheet.getLastRow();
  var rows    = [];

  if (lastRow > 1) {
    var data = sheet.getRange(2, 1, lastRow - 1, 6).getValues();
    for (var i = 0; i < data.length; i++) {
      rows.push({
        timestamp:    data[i][0],
        name:         data[i][1],
        email:        data[i][2],
        phone:        data[i][3],
        field:        data[i][4],
        contribution: data[i][5]
      });
    }
  }

  var json     = JSON.stringify(rows);
  var callback = e.parameter.callback;

  if (callback) {
    return ContentService
      .createTextOutput(callback + '(' + json + ')')
      .setMimeType(ContentService.MimeType.JAVASCRIPT);
  }

  return ContentService
    .createTextOutput(json)
    .setMimeType(ContentService.MimeType.JSON);
}

// Test-Funktion — direkt in Apps Script ausführen um die Verbindung zu prüfen
function test() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  Logger.log('Sheet: ' + sheet.getName() + ' | Zeilen: ' + sheet.getLastRow());
}
