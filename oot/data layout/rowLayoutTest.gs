// TODO: REPLACE ALL PARAMETERS WITH VALUES SPECIFIC TO EXPERIMENT
/* ============= START OF PARAMETERS ============= */

// url of the spreadsheet to write the results
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

// whether to range access or random column access
var RANGE_ACCESS = true;
// name of experiment to be written to results sheet
var EXPER_NAME = "row layout tests"
// sheet name of results spreadsheet to be written to
var SHEET_NAME = "method 1";

/* ============= END OF PARAMETERS ============= */

/*  Runs experiments on all spreadsheets specified by `sizes` array.
    This is the main function to be called for running the experiment. */
function loop() {
  var results = [];
  var tenruns = []; // to average across 10 trials
  for (j = 0; j < 10; j++) {
    for (i = 0; i < sizes.length; i++) {
      var size = sizes[i];
      if (RANGE_ACCESS) {
        var ret = rangeAccess(size, urls[size]);
      } else {
        var ret = randomColumnAccess(size, urls[size]);
      }
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

/*  Measures time to range access `size` row spreadsheet specified by `url`.
    Experiment is performed on copy of the spreadsheet. */
function rangeAccess(size, url) {
  var ss = SpreadsheetApp.openByUrl(url);
  var sheet = ss.getActiveSheet();
  var startDate = new Date();
  var result = 0;
  var resultCell = sheet.getRange(4, 19); // single cell
  var oldval = resultCell.getValue();
  resultCell.setFormula("=COUNT(A2:R" + size + ")");
  result += resultCell.getValue();

  var endDate = new Date();

  // clean up
  resultCell.setValue(oldval);

  return (endDate.getTime() - startDate.getTime());
}


function columnToLetter(column) {
  var temp, letter = '';
  while (column > 0) {
    temp = (column - 1) % 26;
    letter = String.fromCharCode(temp + 65) + letter;
    column = (column - temp - 1) / 26;
  }
  return letter;
}

/*  Measures time to access columns in sheet in random order on `size` row spreadsheet specified by `url`.
    Experiment is performed on copy of the spreadsheet. */
function randomColumnAccess(size, url) {
  var ss = SpreadsheetApp.openByUrl(url);
  var sheet = ss.getActiveSheet();
  var columns = [];
  var lastColumn = sheet.getLastColumn();
  for (var i = 1; i <= lastColumn; i++) {
    columns.push(columnToLetter(i));
  }
  var shuffled = shuffle(columns);
  var result = 0;
  var startDate = new Date();
  var resultCell = sheet.getRange(4, 19);
  var oldval = resultCell.getValue();

  for (var j = 0; j < shuffled.length; j++) {
    resultCell.setFormula("=COUNT(" + shuffled[j] + "1:" + shuffled[j] + size + ")");
    result += resultCell.getValue();
  }
  var endDate = new Date();

  // clean up
  resultCell.setValue(oldval);

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