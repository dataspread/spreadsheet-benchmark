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
var EXPER_NAME = "incremental update tests";
// sheet name of results spreadsheet to be written to
var SHEET_NAME = "method 1";

// TODO: Change values in countif function

/* ============= END OF PARAMETERS ============= */

/*  Runs the experiment on spreadsheet of size `size` as specified by the mapping in `urls`.
    This is the main function to be called for running a trial of the experiment. */
function experiment(size) {
    var results_sheet = SpreadsheetApp.openByUrl(RESULTS_URL).getSheetByName(SHEET_NAME);
    // result is of form [initial time, recalc time]
    var result = countif(size, urls[size]);
    writeToSheet(results_sheet, size, result);
}

/*  Writes the date, experiment name, size, and result (trial time) to a spreadsheet, 
    and highlights the background of the result. */
function writeToSheet(sheet, size, result) {
    var time = new Date();
    var lastRow = sheet.getLastRow() + 1;
    sheet.getRange(lastRow, 1).setValue(Utilities.formatDate(time, 'America/Chicago', 'MMMM dd, yyyy HH:mm:ss Z'));
    lastRow++;
    sheet.getRange(lastRow, 1).setValue(EXPER_NAME + " (initial)");
    sheet.getRange(lastRow, 2).setValue(EXPER_NAME + " (recalc)");
    lastRow++;
    sheet.getRange(lastRow, 1).setValue(size);
    lastRow++;
    // initial
    sheet.getRange(lastRow, 1).setValue(result[0]).setBackground("orange");
    // recalc
    sheet.getRange(lastRow, 2).setValue(result[1]).setBackground("orange");
}

/*  Measures time to countif initially and update the countif on `size` rows on 
    the spreadsheet specified by `url`.
    Experiment is performed on copy of the spreadsheet. */
function countif(size, url) {
    ret = [];
    var ss = SpreadsheetApp.openByUrl(url);
    // perform experiment on copy of spreadsheet
    // copy is added to "Recent", not to location of original spreadsheet
    var ss = ss.copy(ss.getName() + "_" + Date.now());
    var sheet = ss.getActiveSheet();
    var oldval = sheet.getRange(1, 6).getValue(); // replace
    // do initial count
    var date = new Date();
    sheet.getRange(4, 18).setFormula("=COUNTIF(J2:J" + size + ", 1)"); // replace
    var firstCount = sheet.getRange(4, 18).getValue(); // replace
    var endDate = new Date();
    ret.push(endDate.getTime() - date.getTime());

    // now change one value to trigger recomputation
    var secondDate = new Date();
    sheet.getRange(1, 6).setValue(2016); // replace
    var secondCount = sheet.getRange(4, 18).getValue(); // replace
    var secondEndDate = new Date();
    ret.push(secondEndDate.getTime() - secondDate.getTime());

    // clean up
    sheet.getRange(1, 6).setValue(oldval); // replace
    console.log("first count is " + firstCount);
    console.log("second count is " + secondCount);
    return ret;
}