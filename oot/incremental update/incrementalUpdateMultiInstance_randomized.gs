// TODO: REPLACE ALL PARAMETERS WITH VALUES SPECIFIC TO EXPERIMENT
/* ============= START OF PARAMETERS ============= */

// url of the spreadsheet to write the results
var RESULTS_URL = "results_url"; // e.g. "https://docs.google.com/spreadsheets/d/ABCXYZ/edit" 
// spreadsheet url of data. since number of instances will be varied, spreadsheet size stays constant
var url = "url" // e.g. "https://docs.google.com/spreadsheets/d/ABCXYZ/edit";
// number of rows in the spreadsheet specified by url
var SIZE = 0; // e.g. 90000

// name of experiment to be written to results sheet
var EXPER_NAME = "incremental update multi instance tests";
// sheet name of results spreadsheet to be written to
var SHEET_NAME = "method 1";

// TODO: Change values in `sum` function

/* ============= END OF PARAMETERS ============= */

/*  Runs the experiment on spreadsheet of size `size` with `num_instances` of formulas.
    This is the main function to be called for running a trial of the experiment. */
function experiment(num_instances) {
  var result = sum(SIZE, url, num_instances);
  var results_sheet = SpreadsheetApp.openByUrl(RESULTS_URL).getSheetByName(SHEET_NAME);
  writeToSheet(results_sheet, num_instances, result);
}

/*  Writes the date, experiment name, size, and result (trial time) to a spreadsheet, 
    and highlights the background of the result. */
function writeToSheet(sheet, num_instances, result) {
  var time = new Date();
  var lastRow = sheet.getLastRow() + 1;
  sheet.getRange(lastRow, 1).setValue(Utilities.formatDate(time, 'America/Chicago', 'MMMM dd, yyyy HH:mm:ss Z'));
  lastRow++;
  sheet.getRange(lastRow, 1).setValue(EXPER_NAME);
  lastRow++;
  sheet.getRange(lastRow, 1).setValue(num_instances);
  lastRow++;
  sheet.getRange(lastRow, 1).setValue(result).setBackground("orange");
}

/*  Measures time to compute `inst` formulas on spreadsheet of `size` rows with url `url`.
    Experiment is performed on copy of the spreadsheet. */
function sum(size, url, inst) {
  ret = [];
  var ss = SpreadsheetApp.openByUrl(url);
  // perform experiment on copy of spreadsheet
  // copy is added to "Recent", not to location of original spreadsheet
  var ss = ss.copy(ss.getName() + "_" + Date.now());
  var sheet = ss.getActiveSheet();
  var oldval = sheet.getRange(2, 10).getValue();
  var form_string = "=COUNTIF(J2:J" + size + ", 1)";
  var value_to_set = oldval == 1 ? 0 : 1;
  var data = sheet.getRange(1, 18, inst, 18).getValues();
  for (z = 0; z < inst; z++) {    // insert form_string inst times
    data[z][1] = form_string;
  }
  sheet.getRange(1, 18, inst, 18).setValues(data);   // set them all at once
  var res = sheet.getRange(1, 18, inst, 18).getValues();   // get them all again to force computation

  var date = new Date();
  sheet.getRange(2, 10).setValue(value_to_set); // change 1 value
  var res = sheet.getRange(1, 18, inst, 18).getValues();   // get them all again to force recomputation
  var endDate = new Date();

  sheet.getRange(2, 10).setValue(oldval);
  // clean up, 0 out the formulas
  for (z = 0; z < inst; z++) {
    data[z][1] = "0";
  }
  sheet.getRange(1, 18, inst, 18).setValues(data);

  ret = endDate.getTime() - date.getTime();
  return ret;
}
