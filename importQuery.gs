function importQueryExp() {
    const filename = "CSVFile_ExpendITDwBudget.csv"; // Please set the filename of CSV file on your Google Drive.
    const file = DriveApp.getFilesByName(filename);
    if (!file.hasNext()) {
      throw new Error(`"${filename}" was not found.`);
    }
    const csv = file.next().getBlob().getDataAsString();
    const values = Utilities.parseCsv(csv, ",");
    const ssE = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ssE.getSheetByName('Import');
    sheet.getRange(3, 1, sheet.getLastRow(), sheet.getLastColumn()).clearContent();
    sheet.getRange(2, 1, values.length, values[0].length).setValues(values);
    sheet.getRange('P1').setValue(new Date());
  }

function importQueryExpF1() {
    const filename = "CSVFile_Exp_F1.csv"; // Please set the filename of CSV file on your Google Drive.
    const file = DriveApp.getFilesByName(filename);
    if (!file.hasNext()) {
      throw new Error(`"${filename}" was not found.`);
    }
    const fileNext = file.next();
    const csv = fileNext.getBlob().getDataAsString();
    const values = Utilities.parseCsv(csv, ",");
    const ssE = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ssE.getSheetByName('Import');
    sheet.getRange(3, 1, sheet.getLastRow(), sheet.getLastColumn()).clearContent();
    sheet.getRange(2, 1, values.length, values[0].length).setValues(values);

    //var dateC = fileNext.getDateCreated();
    var dateC = fileNext.getLastUpdated();
    // Logger.log(dateC);
    // var name = fileNext.getName()
    // Logger.log(name);

    sheet.getRange('J1').setValue(dateC); //set date from created csv file info
  }

  function importQueryLabor() {
    const filename = "CSVFile_Labor.csv"; // Set the filename of CSV file on your Google Drive.
    const file = DriveApp.getFilesByName(filename);
    if (!file.hasNext()) {
      throw new Error(`"${filename}" was not found.`);
    }
    const csv = file.next().getBlob().getDataAsString();
    const values = Utilities.parseCsv(csv, ",");
    const ssL = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ssL.getSheetByName('Import');
    sheet.getRange(3, 1, sheet.getLastRow(), 7).clearContent();
    sheet.getRange(2, 1, values.length, values[0].length).setValues(values);
    sheet.getRange('J1').setValue(new Date()); 

    const filenameP = "CSVFile_PayNo.csv"; // to update the last payroll posted in Banner
    const fileP = DriveApp.getFilesByName(filenameP);
    if (!fileP.hasNext()) {
      throw new Error(`"${filenameP}" was not found.`);
    }
    const csvP= fileP.next().getBlob().getDataAsString();
    //console.log("csvP:", csvP);
    const valuesP = Utilities.parseCsv(csvP, ",");
    ssL.getRange('J2').setValue(valuesP[1]); //set R value
  }

function importQueryLaborF1() {
    const filename = "CSVFile_Labor_F1.csv"; // Set the filename of CSV file on your Google Drive.
    const file = DriveApp.getFilesByName(filename);
    if (!file.hasNext()) {
      throw new Error(`"${filename}" was not found.`);
    }
    const fileNext = file.next();
    const csv = fileNext.getBlob().getDataAsString();
    const values = Utilities.parseCsv(csv, ",");
    const ssL = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ssL.getSheetByName('Import');
    sheet.getRange(3, 1, sheet.getLastRow(), 7).clearContent();
    sheet.getRange(2, 1, values.length, values[0].length).setValues(values);

    //var dateC = fileNext.getDateCreated();
    var dateC = fileNext.getLastUpdated();
    // Logger.log(dateC);
    // var name = fileNext.getName()
    // Logger.log(name);

    sheet.getRange('J1').setValue(dateC); //set date from created csv file info
    
    const filenameP = "CSVFile_PayNo.csv"; // to update the last payroll posted in Banner
    const fileP = DriveApp.getFilesByName(filenameP);
    if (!fileP.hasNext()) {
      throw new Error(`"${filenameP}" was not found.`);
    }
    
    const csvP= fileP.next().getBlob().getDataAsString();
    //console.log("csvP:", csvP);
    const valuesP = Utilities.parseCsv(csvP, ",");
    sheet.getRange('J2').setValue(valuesP[1]); //set R value
  }

  function importQueryPay() {
    const filename = "CSVFile_Payrate.csv"; // Set the filename of CSV file on your Google Drive.
    const file = DriveApp.getFilesByName(filename);
    if (!file.hasNext()) {
      throw new Error(`"${filename}" was not found.`);
    }
    const csv = file.next().getBlob().getDataAsString();
    const values = Utilities.parseCsv(csv, ",");
    const ssP = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ssP.getSheetByName('TOAD Query Results');
    sheet.getRange(2, 1, sheet.getLastRow(), sheet.getLastColumn()).clearContent();
    sheet.getRange(1, 1, values.length, values[0].length).setValues(values);
    sheet.getRange('I1').setValue(new Date());
  }

  // function onEdit(e) {
//     // Prevent errors if no object is passed.
//     if (!e) return;
//     // Get the active ss.
//     e.source.getActiveSheet()
//         // Set the cell you want to update with the date.
//         .getRange('O1')
//         // Update the date.
//         .setValue(new Date());
//     // Get the active ss.
//     // e.source.getActiveSheet()
//     //     // Set the cell you want to update with the user.
//     //     .getRange('O2')
//     //     // Update the user (only email is available, and only if security settings allow).
//     //     .setValue(e.user.getEmail() );
// }
