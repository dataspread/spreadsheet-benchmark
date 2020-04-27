/* ============= START OF PARAMETERS ============= */

// url to the spreadsheet to write the results
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
// name of experiment to be written to results sheet
var EXPER_NAME = "incremental update tests"
// sheet name of results spreadsheet to be written to
var SHEET_NAME = "method 1"

// TODO: Change values in countif function

/* ============= END OF PARAMETERS ============= */

/*  Runs experiments on all spreadsheets specified by `sizes` array.
    This is the main function to be called for running the experiment. */
function loop() {
  var results1 = [];
  var results2 = [];
  var tenruns1 = []; // to average across 10 trials
  var tenruns2 = [];
  for (j = 0; j < 10; j++) {
    for (i = 0; i < sizes.length; i++) {
      var results = countif(sizes[i], urls[sizes[i]]);
      results1.push(results[0]);
      results2.push(results[1]);
    }
    tenruns1.push(results1);
    tenruns2.push(results2);
    results1 = [];
    results2 = [];
  }
  averageStats(tenruns1, "initial");
  averageStats(tenruns2, "recalc");
}

/*  Takes in an array of trial times for all spreadsheet sizes and writes
    the trial times and average time to the results sheet. 
    The average excludes the max and min trial times for that spreadsheet size. */
function averageStats(times, description) {
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
  writeToSheet(results_sheet, results, description);
}

// this function writes all of the trial times for each size sheet to a datasheet and
// highlights the averaged times for each size sheet
// helper function called by averageStats
function writeToSheet(sheet, results, description) {
  var time = new Date();
  var lastRow = sheet.getLastRow() + 1;
  sheet.getRange(lastRow, 1).setValue(Utilities.formatDate(time, 'America/Chicago', 'MMMM dd, yyyy HH:mm:ss Z'));
  lastRow++;
  sheet.getRange(lastRow, 1).setValue(EXPER_NAME + "(" + description + ")");
  lastRow++;
  for (i = 0; i < results.length; i++) {
    sheet.getRange(lastRow, i + 1).setValue(sizes[i]);
  }
  lastRow++;
  for (i = 0; i < results.length; i++) {
    sheet.getRange(lastRow, i + 1).setValue(results[i]).setBackground("orange");
  }
}

// writes the intermediate trial times for one sized sheet
// helper function called by writeToSheet
function writeInter(sheet, results) {
  var lastRow = sheet.getLastRow() + 1;
  for (i = 0; i < results.length; i++) {
    sheet.getRange(lastRow, i + 1).setValue(results[i]);
  }
}

/*  Measures time to countif initially and update the countif on `size` rows on 
    the spreadsheet specified by `url`. */
function countif(size, url) {
  ret = [];
  var ss = SpreadsheetApp.openByUrl(url);
  var sheet = ss.getActiveSheet();
  var oldval = sheet.getRange(2, 10).getValue(); // replace
  var oldval2 = sheet.getRange(4, 18).getValue();
  // do initial count
  var date = new Date();
  sheet.getRange(4, 18).setFormula("=COUNTIF(J2:J" + size + ", 1)"); // replace
  var firstCount = sheet.getRange(4, 18).getValue(); // replace
  var endDate = new Date();
  ret.push(endDate.getTime() - date.getTime());

  // now change one value to trigger recomputation
  var secondDate = new Date();
  sheet.getRange(2, 10).setValue(2016); // replace
  var secondCount = sheet.getRange(4, 18).getValue(); // replace
  var secondEndDate = new Date();
  ret.push(secondEndDate.getTime() - secondDate.getTime());

  // clean up
  sheet.getRange(2, 10).setValue(oldval); // replace
  sheet.getRange(4, 18).setValue(oldval2);

  console.log("first count is " + firstCount);
  console.log("second count is " + secondCount);
  return ret;
}
