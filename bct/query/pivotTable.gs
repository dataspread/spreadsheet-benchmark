var DATA_SHEET = "datasheet_url"
var sizes = [150, 6000, 10000, 20000, 30000, 40000, 50000, 60000, 70000, 80000, 90000];

// ---------------------------------------- FORMULA ----------------------------------------
var urls = ["url1", "url2", "..."];

// ---------------------------------------- NO FORMULA ----------------------------------------
//var urls = ["url1", "url2", "..."];

var EXPER_NAME = "sort tests formula"
function loop() {
  var results = [];
  var tenruns = []; // to average across 10 trials
  for (j = 0; j < 10; j ++) {
    for (i = 0; i < sizes.length; i++) {
      var ret = createPivotTable(urls[i]);
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


function createPivotTable(url) {
  var ss=SpreadsheetApp.openByUrl(url);
  var sheet=ss.getActiveSheet();
  var activeid = ss.getActiveSheet().getSheetId();
  var r = ss.getDataRange();  
  
  var pivotTableParams = {};
  
  // The source indicates the range of data you want to put in the table.
  // optional arguments: startRowIndex, startColumnIndex, endRowIndex, endColumnIndex
  pivotTableParams.source = {
    sheetId: activeid
  };
  
  // Group rows, the 'sourceColumnOffset' corresponds to the column number in the source range
  // eg: 0 to group by the first column
  pivotTableParams.rows = [{
    sourceColumnOffset: 0,
    sortOrder: "ASCENDING"
  }];
  
  // Defines how a value in a pivot table should be calculated.
  pivotTableParams.values = [{
    summarizeFunction: "SUM",
    sourceColumnOffset: 14
  }];
    
  // Create a new sheet which will contain our Pivot Table
  var pivotTableSheet = ss.insertSheet("pivot");
  var pivotTableSheetId = pivotTableSheet.getSheetId();
  
  // Add Pivot Table to new sheet
  // Meaning we send an 'updateCells' request to the Sheets API
  // Specifying via 'start' the sheet where we want to place our Pivot Table
  // And in 'rows' the parameters of our Pivot Table
  var request = {
    "updateCells": {
      "rows": {
        "values": [{
          "pivotTable": pivotTableParams
        }]
      },
      "start": {
        "sheetId": pivotTableSheetId
      },
      "fields": "pivotTable"
    }
  };
  var date = new Date();
  Sheets.Spreadsheets.batchUpdate({'requests': [request]}, ss.getId());
  var val = pivotTableSheet.getRange(4, 2).getValue();
  var endDate = new Date();
  console.log("val is " + val);
  ret = endDate.getTime() - date.getTime();

  //clean up
  ss.deleteSheet(pivotTableSheet);
  return ret;
}