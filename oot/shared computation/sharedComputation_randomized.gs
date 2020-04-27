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
var EXPER_NAME = "shared computation tests";
// sheet name of results spreadsheet to be written to
var SHEET_NAME = "method 1";
// whether experiment repeats computation or reuse intermediate results
var IS_REPEATED = true;

// TODO: Change values in `repeated` and `reusable` functions

/* ============= END OF PARAMETERS ============= */

/*  Runs the experiment on spreadsheet of size `size` as specified by the mapping in `urls`.
    This is the main function to be called for running a trial of the experiment. */
function experiment(size) {
    if (IS_REPEATED) {
        var result = repeated(size, urls[size]);
    } else {
        var result = reusable(size, urls[size]);
    }
    var results_sheet = SpreadsheetApp.openByUrl(RESULTS_URL).getSheetByName(SHEET_NAME);
    writeToSheet(results_sheet, size, result);
}

/*  Writes the date, experiment name, size, and result(trial time) to a spreadsheet, 
    and highlights the background of the result. */
function writeToSheet(sheet, size, results) {
    var time = new Date();
    var lastRow = sheet.getLastRow() + 1;
    sheet.getRange(lastRow, 1).setValue(Utilities.formatDate(time, 'America/Chicago', 'MMMM dd, yyyy HH:mm:ss Z'));
    lastRow++;
    sheet.getRange(lastRow, 1).setValue(EXPER_NAME);
    lastRow++;
    sheet.getRange(lastRow, 1).setValue(size);
    lastRow++;
    sheet.getRange(lastRow, 1).setValue(results).setBackground("orange");
}

/*  Measures time to compute sum using intermediate results on spreadsheet of 
    size `size` specified by `url`.
    Experiment is performed on copy of the spreadsheet. */
function reusable(size, url) {
    var ss = SpreadsheetApp.openByUrl(url);
    // perform experiment on copy of spreadsheet
    // copy is added to "Recent", not to location of original spreadsheet
    var ss = ss.copy(ss.getName() + "_" + Date.now());
    var sheet = ss.getActiveSheet();
    // insert two columns at the beginning
    sheet.insertColumnBefore(1);
    sheet.insertColumnBefore(1);
    // insert formulas that use previous cell's result
    var data = sheet.getRange(1, 1, size, 2).getValues();
    data[0][0] = "1";
    data[0][1] = "=A1";
    for (z = 1; z < size; z++) {
        data[z][0] = z + 1;
        data[z][1] = "=B" + z + "+A" + (z + 1);
    }
    sheet.getRange(1, 1, size, 2).setValues(data);

    var oldVal = sheet.getRange(1, 1).getValue();
    console.log(oldVal);

    var date = new Date();
    // update inital value and get the final result
    sheet.getRange(1, 1).setValue(oldVal + 1);
    var count = sheet.getRange(size, 2).getValue();
    var endDate = new Date();

    console.log("count is " + count);
    // clean up and delete first two columns
    sheet.deleteColumn(1);
    sheet.deleteColumn(1);

    ret = endDate.getTime() - date.getTime();
    return ret;
}

/*  Measures time to compute sum without using intermediate results on spreadsheet of 
    size `size` specified by `url`.
    Experiment is performed on copy of the spreadsheet. */
function repeated(size, url) {
    var ss = SpreadsheetApp.openByUrl(url);
    // perform experiment on copy of spreadsheet
    // copy is added to "Recent", not to location of original spreadsheet
    var ss = ss.copy(ss.getName() + "_" + Date.now());
    var sheet = ss.getActiveSheet();
    // insert two columns at the beginning
    sheet.insertColumnBefore(1);
    sheet.insertColumnBefore(1);
    // insert formulas that calculate from scratch
    var data = sheet.getRange(1, 1, size, 2).getValues();
    data[0][0] = "1";
    data[0][1] = "=A1";
    for (z = 1; z < size; z++) {
        data[z][0] = z + 1;
        data[z][1] = "=SUM(A1:A" + (z + 1) + ")";
    }
    sheet.getRange(1, 1, size, 2).setValues(data);

    var oldVal = sheet.getRange(1, 1).getValue();
    console.log(oldVal);

    var date = new Date();
    // update inital value and get the final result
    sheet.getRange(1, 1).setValue(oldVal + 1);
    var count = sheet.getRange(size, 2).getValue();
    var endDate = new Date();

    console.log("count is " + count);
    // clean up and delete first two columns
    sheet.deleteColumn(1);
    sheet.deleteColumn(1);

    ret = endDate.getTime() - date.getTime();
    return ret;
}