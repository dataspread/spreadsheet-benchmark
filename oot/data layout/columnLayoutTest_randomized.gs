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

// whether to access sequentially or random
var SEQUENTIAL_ACCESS = true;
// name of experiment to be written to results sheet
var EXPER_NAME = "column layout tests"
// sheet name of results spreadsheet to be written to
var SHEET_NAME = "method 1"

// TODO: Change values in sequential and random functions

/* ============= END OF PARAMETERS ============= */


/*  Runs the experiment on spreadsheet of size `size` as specified by the mapping in `urls`.
    This is the main function to be called for running a trial of the experiment. */
function experiment(size) {
    if (SEQUENTIAL_ACCESS) {
        var result = sequential(size, urls[size]);
    } else {
        var result = random(size, urls[size]);
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

/*  Measures time to access `size` rows sequentially within the first column on the 
    spreadsheet specified by `url`.
    Experiment is performed on copy of the spreadsheet. */
function sequential(size, url) {
    var ss = SpreadsheetApp.openByUrl(url);
    // perform experiment on copy of spreadsheet
    // copy is added to "Recent", not to location of original spreadsheet
    var ss = ss.copy(ss.getName() + "_" + Date.now());
    var sheet = ss.getActiveSheet();
    var startDate = new Date();
    for (i = 1; i < size; i++) {
        var temp = sheet.getRange(i, 1);
    }
    var endDate = new Date();

    return (endDate.getTime() - startDate.getTime());
}

/*  Measures time to access `size` rows in random order on the spreadsheet specified by `url`.
    Experiment is performed on copy of the spreadsheet. */
function random(size, url) {
    var ss = SpreadsheetApp.openByUrl(url);
    // perform experiment on copy of spreadsheet
    // copy is added to "Recent", not to location of original spreadsheet
    var ss = ss.copy(ss.getName() + "_" + Date.now());
    var sheet = ss.getActiveSheet();
    var rows = [];
    for (i = 1; i < size; i++) {
        rows.push(i);
    }
    // shuffle order of rows to access
    var shuffled = shuffle(rows);
    var startDate = new Date();
    for (j = 0; j < shuffled.length; j++) {
        var temp = sheet.getRange(shuffled[j], 1);
    }
    var endDate = new Date();

    return (endDate.getTime() - startDate.getTime());
}

/* Shuffles `array1` by swapping each element randomly */
function shuffle(array1) {
    var ctr = array1.length, temp, index;

    while (ctr > 0) {
        index = Math.floor(Math.random() * ctr);
        ctr--;
        temp = array1[ctr];
        array1[ctr] = array1[index];
        array1[index] = temp;
    }
    return array1;
}