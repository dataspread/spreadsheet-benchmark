// TODO: REPLACE ALL PARAMETERS WITH VALUES SPECIFIC TO EXPERIMENT
/* ============= START OF PARAMETERS ============= */

// url of the spreadsheet to write the results
var RESULTS_URL = "results_url"; // e.g. "https://docs.google.com/spreadsheets/d/ABCXYZ/edit" 
// spreadsheet url of data. since number of instances will be varied, spreadsheet size stays constant
var url = "url" // e.g. "https://docs.google.com/spreadsheets/d/ABCXYZ/edit";
// number of rows in the spreadsheet specified by url
var SIZE = 0; // e.g. 90000
// each number represents how many instances of the formula to run with the experiment
// script may time out if array is too large
var num_instances = [inst1, inst2, ...]; // e.g. [5, 100]

// name of experiment to be written to results sheet
var EXPER_NAME = "incremental update multi instance tests";
// sheet name of results spreadsheet to be written to
var SHEET_NAME = "method 1";

// TODO: Change values in `sum` function

/* ============= END OF PARAMETERS ============= */

/*  Runs the experiment on spreadsheet of size `size` with `num_instances` of formulas.
    This is the main function to be called for running a trial of the experiment. */
function loop() {
  var results1 = []; // for inital fandr
  var tenruns1 = []; // to average across 10 trials
  for (j = 0; j < 10; j++) {
    for (i = 0; i < num_instances.length; i++) {
      var ret1 = sum(SIZE, url, num_instances[i]);
      results1.push(ret1);
    }
    console.log(results1);
    tenruns1.push(results1);
    results1 = [];
  }
  averageStats(tenruns1);
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
    sheet.getRange(lastRow, i + 1).setValue(num_instances[i]);
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

/*  Measures time to compute `inst` formulas on spreadsheet of `size` rows with url `url`. */
function sum(size, url, inst) {
  ret = [];
  var ss = SpreadsheetApp.openByUrl(url);
  var sheet = ss.getActiveSheet();
  var oldval = sheet.getRange(2, 10).getValue();
  var form_string = "=COUNTIF(J2:J" + size + ", 1)";
  var value_to_set = oldval == 1 ? 0 : 1;
  sheet.insertColumnBefore(18);
  var data = sheet.getRange(1, 18, inst, 18).getValues();
  for (z = 0; z < inst; z++) {    // insert form_string inst times
    data[z][1] = form_string;
  }
  sheet.getRange(1, 18, inst, 18).setValues(data);   // set them all at once
  var res = sheet.getRange(1, 18, inst, 18).getValues();   // get them all again to force computation

  var startDate = new Date();
  sheet.getRange(2, 10).setValue(value_to_set); // change 1 value
  var res = sheet.getRange(1, 18, inst, 18).getValues();   // get them all again to force recomputation
  var endDate = new Date();

  sheet.getRange(2, 10).setValue(oldval);
  // clean up
  sheet.deleteColumn(18);

  ret = endDate.getTime() - startDate.getTime();
  return ret;
}
