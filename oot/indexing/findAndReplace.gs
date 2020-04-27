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
// name of experiment to be written to results sheet
var EXPER_NAME = "find and replace tests";
// sheet name of results spreadsheet to be written to
var SHEET_NAME = "method 1";

// TODO: Change values in `fandr` function and arguments to fandr

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
      // `replace` string should not be in the search column
      var ret1 = fandr(sizes[i], urls[sizes[i]], "present", "replace"); // replace `present` with string in search column
      var ret2 = fandr(sizes[i], urls[sizes[i]], "absent", "replace"); // replace `absent` with string not in search column
      results1.push(ret1);
      results2.push(ret2);
    }
    tenruns1.push(results1);
    tenruns2.push(results2);
    results1 = [];
    results2 = [];
  }
  averageStats(tenruns1, EXPER_NAME + " (present)");
  averageStats(tenruns2, EXPER_NAME + " (absent)");
}

/*  Takes in an array of trial times for all spreadsheet sizes and writes
    the trial times and average time to the results sheet. 
    The average excludes the max and min trial times for that spreadsheet size. */
function averageStats(times, experiment) {
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
  writeToSheet(results_sheet, results, experiment);
}

/*  Writes the date, experiment name, trial times, sizes, and averaged results to a spreadsheet, 
    and highlights the background of the result. 
    This function is called by `averageStats`. */
function writeToSheet(sheet, results, experiment) {
  var time = new Date();
  var lastRow = sheet.getLastRow() + 1;
  sheet.getRange(lastRow, 1).setValue(Utilities.formatDate(time, 'America/Chicago', 'MMMM dd, yyyy HH:mm:ss Z'));
  lastRow++;
  sheet.getRange(lastRow, 1).setValue(experiment);
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

/*  Measures time to do find and replace on the spreadsheet of size `size` specified by `url`. */
function fandr(size, url, find, repl) {
  var ss = SpreadsheetApp.openByUrl(url);
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

  // clean up
  for (i = 0; i < rws; i++) {
    try {
      data[i][1] = data[i][1].replace(repl, find);
    }
    catch (err) { continue; }
  }
  sheet.getRange(1, 4, rws, 3).setValues(data); // replace

  ret = endDate.getTime() - startDate.getTime();
  return ret;
}