function createAttestation() {

  var allCoachesFolder = DriveApp.getFolderById('1gESuH-QxnMffGDtduvMA-iBjmZlatdX5');
  var sheet = SpreadsheetApp.getActiveSheet();
  var lastRow = sheet.getLastRow();
  var lastColumn = sheet.getLastColumn();

  var coaches = sheet.getRange(2, 1, lastRow - 1, lastColumn).getDisplayValues();
  //  alternative: hardcoding id of spreadsheet
  //  var coaches = Sheets.Spreadsheets.Values.get('1cX8RJc_tcuhLUYz6oZeLCUy-id-RLsQcW-neN-KFaQE' , 'A2:E4');

  var templateId = '1EcxZcinp7PV4q1UpatL8T7CSTmHqWMjAly5cwAR35hE';

  //    date of report
  var frenchMonths = { "janvier": "01", "février": "02", "mars": "03", "avril": "04", "mai": "05", "juin": "06", "juillet": "07", "août": "08", "septembre": "09", "octobre": "10", "novembre": "11", "décembre": "12"}
  var date = coaches[0][3];
  var year = date.split(' ')[1];
  var monthDigit = frenchMonths[date.split(' ')[0]];

  for (var i= 0; i < coaches.length; i++) {
    if (coaches[i][4] == 'bénévole' ) {

      //    name of coach
      var coachName = coaches[i][0];

      //    hours
      var hours = coaches[i][1];

      //    monthly total in euros
      var total = coaches[i][2];

      //    creating copy of template and getting doc id
      var documentId = DriveApp.getFileById(templateId).makeCopy().getId();

      //    changing name of doc
      var file = DriveApp.getFileById(documentId);
      file.setName(`${year}-${monthDigit}_${coachName}_attestation`);

      //    changing template
      var body = DocumentApp.openById(documentId).getBody();

      body.replaceText('{{name}}', coachName);
      body.replaceText('{{hours}}', hours);
      body.replaceText('{{date}}', date);
      body.replaceText('{{year}}', year);
      body.replaceText('{{total}}', total);

      //    Find coach folder
      var folders = allCoachesFolder.getFolders();
      var doesntExists = true;
      var coachFolder = '';

      while (folders.hasNext()) {
        var folder = folders.next();

        if (folder.getName() === coachName) {
          doesntExists = false;
          coachFolder = folder;
          break;
        }
      }

      // if coach doesn't have their folder on the drive, creates one
      if (doesntExists == true) {
        Logger.log("created a new folder for " + coachName)
        coachFolder =  allCoachesFolder.createFolder(coachName);
      }

      // move file to coach's folder
      file.moveTo(coachFolder);
      Logger.log("created attestation for " + coachName);
    }
  }
  Logger.log("done");
}
