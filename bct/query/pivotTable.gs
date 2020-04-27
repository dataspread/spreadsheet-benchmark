/* ============= START OF PARAMETERS ============= */

// url to the spreadsheet to write the results
var RESULTS_URL = "results_url"; // e.g. "https://docs.google.com/spreadsheets/d/ABCXYZ/edit" 
// spreadsheet sizes to run experiment on
// script may time out if sizes is too large, so sizes should be subset of urls
var sizes = [size1, size2, ...]; // e.g. [10000, 20000]
// mapping from spreadsheet row counts to url of spreadsheet
var urls = {
  size1: "url1", // e.g. 10000: "https://docs.google.com/spreadsheets/d/ABCXYZ/edit"
  size2: "url2",
  // ...
};
// name of experiment to be written to results sheet
var EXPER_NAME = "pivot table tests"
// sheet name of results spreadsheet to be written to
var SHEET_NAME = "method 1"

// TODO: Change values in createPivotTable function

/* ============= END OF PARAMETERS ============= */

/*  Runs experiments on all spreadsheets specified by `sizes` array.
    This is the main function to be called for running the experiment. */
function loop() {
  var results = [];
  var tenruns = []; // to average across 10 trials
  for (j = 0; j < 10; j++) {
    for (i = 0; i < sizes.length; i++) {
      var ret = createPivotTable(urls[sizes[i]]);
      results.push(ret);
    }
    tenruns.push(results);
    results = [];
  }
  averageStats(tenruns);
}

/*  Takes in an array of trial times for all spreadsheet sizes and writes
    the trial times and average time to the results sheet. 
    The average excludes the max and min trial times for that spreadsheet size. */
function averageStats(times) {
  var results_sheet = SpreadsheetApp.openByUrl(RESULTS_URL).getSheetByName(SHEET_NAME);
  var perSize = []; // get all the times for one size sheet
  for (i = 0; i < times[0].length; i++) {
    perSize.push([]);
  }
  for (i = 0; i < times.length; i++) {
    for (j = 0; j < times[i].length; j++) {
      perSize[j].push(times[i][j]);
    }
  }

  var results = [];
  for (z = 0; z < perSize.length; z++) {
    cur = perSize[z];
    // write ALL trial times to results sheet (including max and min)
    writeInter(results_sheet, cur);
    // remove min and max trial times
    cur.splice(cur.indexOf(Math.min.apply(null, cur)), 1);
    cur.splice(cur.indexOf(Math.max.apply(null, cur)), 1);
    var sum = 0;
    for (j = 0; j < cur.length; j++) {
      sum += cur[j];
    }
    results.push(sum / cur.length);
  }
  // write average times and metadata to results sheet
  writeToSheet(results_sheet, results);
}

/*  Writes the date, experiment name, trial times, sizes, and averaged results to a spreadsheet, 
    and highlights the background of the result. 
    This function is called by `averageStats`. */
function writeToSheet(sheet, results) {
  var time = new Date();
  var lastRow = sheet.getLastRow() + 1;
  sheet.getRange(lastRow, 1).setValue(Utilities.formatDate(time, 'America/Chicago', 'MMMM dd, yyyy HH:mm:ss Z'));
  lastRow++;
  sheet.getRange(lastRow, 1).setValue(EXPER_NAME);
  lastRow++;
  // write all sizes to sheet
  for (i = 0; i < results.length; i++) {
    sheet.getRange(lastRow, i + 1).setValue(sizes[i]);
  }
  lastRow++;
  // write all average times to sheet
  for (i = 0; i < results.length; i++) {
    sheet.getRange(lastRow, i + 1).setValue(results[i]).setBackground("orange");
  }
}

/*  Writes the intermediate trial times for one sized sheet.
    This is a helper function called by `writeToSheet`. */
function writeInter(sheet, results) {
  var lastRow = sheet.getLastRow() + 1;
  for (i = 0; i < results.length; i++) {
    sheet.getRange(lastRow, i + 1).setValue(results[i]);
  }
}

/*  Measures time to create pivot table on the spreadsheet specified by `url`. */
function createPivotTable(url) {
  var ss = SpreadsheetApp.openByUrl(url);
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