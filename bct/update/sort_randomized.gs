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
var EXPER_NAME = "sort tests"
// sheet name of results spreadsheet to be written to
var SHEET_NAME = "method 1"

// TODO: Change values in sort function

/* ============= END OF PARAMETERS ============= */

/*  Runs the experiment on spreadsheet of size `size` as specified by the mapping in `urls`.
    This is the main function to be called for running a trial of the experiment. */
function experiment(size) {
    var result = sort(size, urls[size]);
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

/*  Measures time to sort `size` rows of the spreadsheet specified by `url`.
    Experiment is performed on copy of the spreadsheet. */
function sort(size, url) {
    var ss = SpreadsheetApp.openByUrl(url);
    // perform experiment on copy of spreadsheet
    // copy is added to "Recent", not to location of original spreadsheet
    var ss = ss.copy(ss.getName() + "_" + Date.now());
    var sheet = ss.getActiveSheet();
    // TODO: Change range to be sorted and column to sort on
    var range = sheet.getRange("A1:O" + (size)); // replace

    var startDate = new Date();
    range.sort({ column: 1, ascending: false }); // replace
    // get value to ensure sort is complete
    var x = sheet.getRange(5, 2).getValue(); // replace
    var endDate = new Date();

    // clean up (necessary if not working on copy)
    range.sort({ column: 1, ascending: true }); // replace
    var ret = endDate.getTime() - startDate.getTime();
    return ret;
}