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
var EXPER_NAME = "find and replace tests";
// sheet name of results spreadsheet to be written to
var SHEET_NAME = "method 1";

// TODO: Change values in `fandr` function and arguments to fandr

/* ============= END OF PARAMETERS ============= */

/*  Runs the experiment on spreadsheet of size `size` as specified by the mapping in `urls`.
    This is the main function to be called for running a trial of the experiment. */
function experiment(size) {
    var results_sheet = SpreadsheetApp.openByUrl(RESULTS_URL).getSheetByName(SHEET_NAME);
    if (IS_PRESENT) {
        var result = fandr(size, urls[size], "present", "replace"); // replace
    } else {
        // 
        var result = fandr(size, urls[size], "absent", "replace"); // replace
    }
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

/*  Measures time to do find and replace on the spreadsheet of size `size` specified by `url`.
    Experiment is performed on copy of the spreadsheet. */
function fandr(size, url, find, repl) {
    console.log("Starting " + size);
    var ss = SpreadsheetApp.openByUrl(url);
    // perform experiment on copy of spreadsheet
    // copy is added to "Recent", not to location of original spreadsheet
    var ss = ss.copy(ss.getName() + "_" + Date.now());
    var sheet = ss.getActiveSheet();
    var rws = sheet.getLastRow();
    var i, find, repl;

    var startDate = new Date();
    var data = sheet.getRange(1, 4, rws, 3).getValues(); // replace
    // check all rows in column for `find` and replace with `repl`
    for (i = 0; i < rws; i++) {
        try {
            data[i][1] = data[i][1].replace(find, repl);
        }
        catch (err) { continue; }
    }
    sheet.getRange(1, 4, rws, 3).setValues(data); // replace
    var endDate = new Date();

    console.log("Ending " + size);
    ret = endDate.getTime() - startDate.getTime();
    return ret;
}