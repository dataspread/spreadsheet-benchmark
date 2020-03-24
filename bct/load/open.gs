// all sheets need to be opened from a starter sheet, this is that sheet:
var start_url = "url"

// same old data sheet: 
var DATA_SHEET = "url"

var size = [150, 60000, 10000, 20000, 30000, 40000, 50000, 60000, 70000, 80000, 90000];

// these folders should each hold the one sheet you want to open
// ---------------------------------------- WEATHER NO FORMULA ----------------------------------------
var folder_ids = ["folder_1_id", "folder_2_id", "..."];

// ---------------------------------------- FORMULA ----------------------------------------
//var folder_ids = ["folder_1_id", "folder_2_id", "..."];

var EXPER_NAME = "open weather NO formula";

function loop() {
  var results = [];
  var tenruns = []; // to average across 10 trials
  for (j = 0; j < 10; j ++) {
    for (i = 0; i < sizes.length; i++) {
      var ret = open(i);
      results.push(ret);
    }
    tenruns.push(results);
    results = [];
  }
  averageStats(tenruns);
}

// times is arry of 10 arrays, each size # sizes tested
// this function removes the min and max outliers and average the remaining 8 times for each size spreadsheet
function averageStats(times) {
  var results_sheet = SpreadsheetApp.openByUrl(DATA_SHEET).getActiveSheet();
  var perSize = []; // get the all times for one size sheet
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
    writeInter(results_sheet, cur);
    cur.splice(cur.indexOf(Math.min.apply(null, cur)), 1);
    cur.splice(cur.indexOf(Math.max.apply(null, cur)), 1);
    var sum = 0;
    for (j = 0; j < cur.length; j++) {
      sum+=cur[j];
    }
    results.push(sum/cur.length);
  }

  writeToSheet(results_sheet, results);
}

// this function writes all of the trial times for each size sheet to a datasheet and
// highlights the averaged times for each size sheet
// helper function called by averageStats
function writeToSheet(sheet, results) {
  var time = new Date();
  var lastRow = sheet.getLastRow()+1;
  sheet.getRange(lastRow, 1).setValue(Utilities.formatDate(time, 'America/Chicago', 'MMMM dd, yyyy HH:mm:ss Z'));
  lastRow++;
  sheet.getRange(lastRow, 1).setValue(EXPER_NAME);
  lastRow++;
  for (i = 0; i < results.length; i++) {
      sheet.getRange(lastRow, i+1).setValue(sizes[i]);
  }
  lastRow++;
  for (i = 0; i < results.length; i++) {
    sheet.getRange(lastRow, i+1).setValue(results[i]).setBackground("orange");
  }
}

// writes the intermediate trial times for one sized sheet
// helper function called by writeToSheet
function writeInter(sheet, results) {
  var lastRow = sheet.getLastRow()+1;
  for (i = 0; i < results.length; i++) {
    sheet.getRange(lastRow, i+1).setValue(results[i]);
  }
}

function open(index) { 
  var date = new Date();
  var start_time = date.getTime();
  
  var folder = DriveApp.getFolderById(folder_ids[index]);
  var sheet = folder.getFiles().next();
  var ss = SpreadsheetApp.open(sheet).getActiveSheet();
  var lastRow = ss.getLastRow()+1;
  var lastCol = ss.getLastColumn();
  var val = ss.getRange(lastRow,lastCol).getValue();
  var endDate = new Date();
  
  var time = endDate.getTime() - start_time;
  return time;
}