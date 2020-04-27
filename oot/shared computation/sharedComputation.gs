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

/*  Runs experiments on all spreadsheets specified by `sizes` array.
    This is the main function to be called for running the experiment. */
function loop() {
  var results = [];
  var tenruns = []; // to average across 10 trials
  for (j = 0; j < 10; j++) {
    for (i = 0; i < sizes.length; i++) {
      var size = sizes[i];
      if (IS_REPEATED) {
        var ret = repeated(size, urls[size]);
      } else {
        var ret = reusable(size, urls[size]);
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

/*  Measures time to compute sum using intermediate results on spreadsheet of 
    size `size` specified by `url`.
    Experiment is performed on copy of the spreadsheet. */
function reusable(size, url) {
  var ss = SpreadsheetApp.openByUrl(url);
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