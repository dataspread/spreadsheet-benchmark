// TODO: REPLACE ALL PARAMETERS WITH VALUES SPECIFIC TO EXPERIMENT
/* ============= START OF PARAMETERS ============= */

// url of the spreadsheet to write the results
var RESULTS_URL = "results_url"; // e.g. "https://docs.google.com/spreadsheets/d/ABCXYZ/edit" 
// mapping from spreadsheet row counts to id of folder containing spreadsheet
// folder must ONLY contain only the desired spreadsheet. 
// should be ok if there are copies of the same spreadsheet
var urls = {
    // If the folder url looks like `https://drive.google.com/drive/u/0/folders/ABCXYZ`
    // the folder id is "ABCXYZ"
    size1: "folder_id1", // e.g. 10000: "ABCXYZ"
    size2: "folder_id2",
    // ...
};
// name of experiment to be written to results sheet
var EXPER_NAME = "open tests"
// sheet name of results spreadsheet to be written to
var SHEET_NAME = "method 1"

/* ============= END OF PARAMETERS ============= */

/*  Runs the experiment on spreadsheet of size `size` as specified by the mapping in `urls`.
    This is the main function to be called for running a trial of the experiment. */
function experiment(size) {
    var result = open(urls[size]);
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

/*  Measures time to open spreadsheet in folder specified by `folder_id`.
    Experiment is performed on copy of the spreadsheet. */
function open(folder_id) {
    var folder = DriveApp.getFolderById(folder_id);
    var file = folder.getFiles().next();
    var file = file.makeCopy(); // comment to open original

    var startDate = new Date();
    var ss = SpreadsheetApp.open(file).getActiveSheet();
    var lastRow = ss.getLastRow() + 1;
    var lastCol = ss.getLastColumn();
    // get value to ensure whole sheet is loaded
    var val = ss.getRange(lastRow, lastCol).getValue();
    var endDate = new Date();

    var ret = endDate.getTime() - startDate.getTime();
    return ret;
}