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
var EXPER_NAME = "redundant tests";
// sheet name of results spreadsheet to be written to
var SHEET_NAME = "method 1";

// TODO: Change values in `insertFormula` function

/* ============= END OF PARAMETERS ============= */

/*  Runs the experiment on spreadsheet of size `size` as specified by the mapping in `urls`.
    This is the main function to be called for running a trial of the experiment. */
function experiment(size) {
  var results_sheet = SpreadsheetApp.openByUrl(RESULTS_URL).getSheetByName(SHEET_NAME);
  var result = insertFormula(size, urls[size]);
  writeToSheet(results_sheet, size, result);
}

/*  Writes the date, experiment name, size, and result(trial time) to a spreadsheet, 
    and highlights the background of the result. */
function writeToSheet(sheet, size, result) {
  var time = new Date();
  var lastRow = sheet.getLastRow() + 1;
  sheet.getRange(lastRow, 1).setValue(Utilities.formatDate(time, 'America/Chicago', 'MMMM dd, yyyy HH:mm:ss Z'));
  lastRow++;
  sheet.getRange(lastRow, 1).setValue(EXPER_NAME + "(single formula");
  sheet.getRange(lastRow, 2).setValue(EXPER_NAME + "(multi formula");
  lastRow++;
  sheet.getRange(lastRow, 1).setValue(size);
  lastRow++;
  sheet.getRange(lastRow, 1).setValue(result[0]).setBackground("orange");
  sheet.getRange(lastRow, 2).setValue(result[1]).setBackground("orange");
}

/*  Measures time to compute a formula one vs multiple times on spreadsheet of 
    size `size` specified by `url`.
    Experiment is performed on copy of the spreadsheet. */
function insertFormula(size, url) {
  var ss = SpreadsheetApp.openByUrl(url);
  // perform experiment on copy of spreadsheet
  // copy is added to "Recent", not to location of original spreadsheet
  var ss = ss.copy(ss.getName() + "_" + Date.now());
  var sheet = ss.getActiveSheet();
  var prev_data = sheet.getRange(1, 18, 6, 1).getValues();
  var warm_start = "=COUNTIF(A1:Q" + size + ", 2017)";
  var form_string = "=COUNTIF(A1:Q" + size + ", 2016)";
  // set an independent formula to reduce inital overhead of using a new spreedsheet
  var data = sheet.getRange(1, 18).getValues();
  data[0][0] = warm_start
  sheet.getRange(1, 18).setValues(data);
  var data = sheet.getRange(1, 18).getValues();
  // multi instances
  var start_date = new Date();
  // set formula to subsequent rows and calculate time
  for (i = 0; i < 5; i++) {
    var row = i + 2; // i+2 because after warm start and sheets are 1-indexed
    var data = sheet.getRange(row, 18).getValues();
    data[0][0] = form_string
    sheet.getRange(row, 18).setValues(data);
    var data = sheet.getRange(row, 18).getValues();
    // calculate time for first  of 5 formula
    if (i == 0) {
      first_date = new Date();
      first_duration = first_date.getTime() - start_date.getTime();
    }
  }
  // calculate time for all 5 formulas
  end_date = new Date();
  end_duration = end_date.getTime() - start_date.getTime();

  // clean up, set back to original values
  sheet.getRange(1, 18, 6, 1).setValues(prev_data);

  return [first_duration, end_duration];
}