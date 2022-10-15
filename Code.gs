function onEdit(e) {
    // Prevent errors if no object is passed.
    if (!e) return;
    // Get the active sheet.
    e.source.getActiveSheet()
        // Set the cell you want to update with the date.
        .getRange('O1')
        // Update the date.
        .setValue(new Date());
    // Get the active sheet.
    // e.source.getActiveSheet()
    //     // Set the cell you want to update with the user.
    //     .getRange('O2')
    //     // Update the user (only email is available, and only if security settings allow).
    //     .setValue(e.user.getEmail() );
}

//for expenditure query import
function importQueryExp() {
    const filename = "CSVFile_Exp.csv"; // Please set the filename of CSV file on your Google Drive.
    const file = DriveApp.getFilesByName(filename);
    if (!file.hasNext()) {
      throw new Error(`"${filename}" was not found.`);
    }
    const csv = file.next().getBlob().getDataAsString();
    const values = Utilities.parseCsv(csv, ",");
    const sheet = SpreadsheetApp.getActiveSheet();
    sheet.getRange(3, 1, sheet.getLastRow(), sheet.getLastColumn()).clearContent();
    sheet.getRange(2, 1, values.length, values[0].length).setValues(values);
  }

//for labor query import
  function importQueryLabor() {
    const filename = "CSVFile_Labor.csv"; // Set the filename of CSV file on your Google Drive.
    const file = DriveApp.getFilesByName(filename);
    if (!file.hasNext()) {
      throw new Error(`"${filename}" was not found.`);
    }
    const csv = file.next().getBlob().getDataAsString();
    const values = Utilities.parseCsv(csv, ",");
    const sheet = SpreadsheetApp.getActiveSheet();
    sheet.getRange(3, 1, sheet.getLastRow(), 7).clearContent();
    sheet.getRange(2, 1, values.length, values[0].length).setValues(values);

    const filenameP = "CSVFile_Payno.csv"; // to update the last payroll posted in Banner
    const fileP = DriveApp.getFilesByName(filenameP);
    if (!fileP.hasNext()) {
      throw new Error(`"${filenameP}" was not found.`);
    }
    const csvP= fileP.next().getBlob().getDataAsString();
    //console.log("csvP:", csvP);
    const valuesP = Utilities.parseCsv(csvP, ",");
    sheet.getRange(2, 10).setValue(valuesP[1]);
  }

//for payrate query import
  function importQueryPay() {
    const filename = "CSVFile_Payrate.csv"; // Set the filename of CSV file on your Google Drive.
    const file = DriveApp.getFilesByName(filename);
    if (!file.hasNext()) {
      throw new Error(`"${filename}" was not found.`);
    }
    const csv = file.next().getBlob().getDataAsString();
    const values = Utilities.parseCsv(csv, ",");
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName('TOAD Query Results');
    sheet.getRange(2, 1, sheet.getLastRow(), sheet.getLastColumn()).clearContent();
    sheet.getRange(1, 1, values.length, values[0].length).setValues(values);
  }