// TODO: REPLACE ALL PARAMETERS WITH VALUES SPECIFIC TO EXPERIMENT
/* ============= START OF PARAMETERS ============= */

// url of the spreadsheet to write the results
var RESULTS_URL = "results_url"; // e.g. "https://docs.google.com/spreadsheets/d/ABCXYZ/edit" 
// mapping from spreadsheet row counts to url of spreadsheet
var urls = {
  size1: "url1", // e.g. 10000: "https://docs.google.com/spreadsheets/d/ABCXYZ/edit"
  size2: "url2",
  // ...
};
// name of experiment to be written to results sheet
var EXPER_NAME = "pivot table tests";
// sheet name of results spreadsheet to be written to
var SHEET_NAME = "method 1";

// TODO: Change values in createPivotTable function

/* ============= END OF PARAMETERS ============= */

/*  Runs the experiment on spreadsheet of size `size` as specified by the mapping in `urls`.
    This is the main function to be called for running a trial of the experiment. */
function experiment(size) {
  var result = createPivotTable(urls[size]);
  var results_sheet = SpreadsheetApp.openByUrl(RESULTS_URL).getSheetByName(SHEET_NAME);
  writeToSheet(results_sheet, size, result);
}

/*  Writes the date, experiment name, size, and result(trial time) to a spreadsheet, 
    and highlights the background of the result. */
function writeToSheet(sheet, size, result) {
  var time = new Date();
  var lastRow = sheet.getLastRow() + 1;
  sheet.getRange(lastRow, 1).setValue(Utilities.formatDate(time, 'America/Chicago', 'MMMM dd, yyyy HH:mm:ss Z'));
  lastRow++;
  sheet.getRange(lastRow, 1).setValue(EXPER_NAME);
  lastRow++;
  sheet.getRange(lastRow, 1).setValue(size);
  lastRow++;
  sheet.getRange(lastRow, 1).setValue(result).setBackground("orange");
}

/*  Measures time to create pivot table on the spreadsheet specified by `url`.
    Experiment is performed on copy of the spreadsheet. */
function createPivotTable(url) {
  var ss = SpreadsheetApp.openByUrl(url);
  // perform experiment on copy of spreadsheet
  // copy is added to "Recent", not to location of original spreadsheet
  var ss = ss.copy(ss.getName() + "_" + Date.now());
  var activeid = ss.getActiveSheet().getSheetId();

  var pivotTableParams = {};

  // The source indicates the range of data you want to put in the table.
  // optional arguments: startRowIndex, startColumnIndex, endRowIndex, endColumnIndex
  pivotTableParams.source = { // replace
    sheetId: activeid
  };

  // Group rows, the 'sourceColumnOffset' corresponds to the column number in the source range
  // eg: 0 to group by the first column
  pivotTableParams.rows = [{ // replace
    sourceColumnOffset: 1,
    sortOrder: "ASCENDING"
  }];

  // Defines how a value in a pivot table should be calculated.
  pivotTableParams.values = [{ // replace
    summarizeFunction: "SUM",
    sourceColumnOffset: 8
  }];

  // Create a new sheet which will contain our Pivot Table
  var pivotTableSheet = ss.insertSheet("pivot");
  var pivotTableSheetId = pivotTableSheet.getSheetId();

  // Add Pivot Table to new sheet
  // Meaning we send an 'updateCells' request to the Sheets API
  // Specifying via 'start' the sheet where we want to place our Pivot Table
  // And in 'rows' the parameters of our Pivot Table
  var request = { // replace
    "updateCells": {
      "rows": {
        "values": [{
          "pivotTable": pivotTableParams
        }]
      },
      "start": {
        "sheetId": pivotTableSheetId
      },
      "fields": "pivotTable"
    }
  };

  var startDate = new Date();
  Sheets.Spreadsheets.batchUpdate({ 'requests': [request] }, ss.getId());
  // get value to ensure update is complete
  var val = pivotTableSheet.getRange(4, 2).getValue(); // replace
  var endDate = new Date();
  console.log("val is " + val);

  // clean up (necessary if not working on copy)
  ss.deleteSheet(pivotTableSheet);

  ret = endDate.getTime() - startDate.getTime();
  return ret;
}