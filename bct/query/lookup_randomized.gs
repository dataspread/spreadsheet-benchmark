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

// value to be passed into the `is_sorted` argument of vlookup formula
var IS_SORTED = "TRUE";
// name of experiment to be written to results sheet
var EXPER_NAME = "lookup tests";
// sheet name of results spreadsheet to be written to
var SHEET_NAME = "method 1";

// TODO: Change values in vlookup function

/* ============= END OF PARAMETERS ============= */

/*  Runs the experiment on spreadsheet of size `size` as specified by the mapping in `urls`.
    This is the main function to be called for running a trial of the experiment. */
function experiment(size) {
  var result = vlookup(size, urls[size]);
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

/*  Measures time to vlookup `size` rows on the spreadsheet specified by `url`.
    Experiment is performed on copy of the spreadsheet. */
function vlookup(size, url) {
  var ss = SpreadsheetApp.openByUrl(url);
  // perform experiment on copy of spreadsheet
  // copy is added to "Recent", not to location of original spreadsheet
  var ss = ss.copy(ss.getName() + "_" + Date.now());
  var sheet = ss.getActiveSheet();
  sheet.insertColumnBefore(1);
  var rws = sheet.getLastRow();

  // add data to look for, numbers 1-size in ascending order
  var data = sheet.getRange(2, 1, rws, 1).getValues();
  for (z = 0; z < rws; z++) {
    data[z][0] = z;
  }
  sheet.getRange(2, 1, rws, 1).setValues(data);

  var startDate = new Date();
  var oldval = sheet.getRange(4, 18).getValue(); // replace
  sheet.getRange(4, 18).setFormula("=VLOOKUP(9123, A1:J" + size + ", 3, " + IS_SORTED + ")"); // replace
  // get value to ensure update is complete
  var count = sheet.getRange(4, 18).getValue(); // replace
  var endDate = new Date();

  // clean up (necessary if not working on copy)
  sheet.getRange(4, 18).setValue(oldval); // replace
  sheet.deleteColumn(1);

  ret = endDate.getTime() - startDate.getTime();
  return ret;
}