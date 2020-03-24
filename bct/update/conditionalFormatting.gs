var DATA_SHEET = "datasheet_url"
var sizes = [150, 6000, 10000, 20000, 30000, 40000, 50000, 60000, 70000, 80000, 90000];

// ---------------------------------------- FORMULA ----------------------------------------
var urls = ["url1", "url2", "..."];

// ---------------------------------------- NO FORMULA ----------------------------------------
//var urls = ["url1", "url2", "..."];

var EXPER_NAME = "cf tests no formula; condition > 2"
function loop() {
  var results = [];
  var tenruns = []; // to average across 10 trials
  for (j = 0; j < 10; j ++) {
    for (i = 0; i < sizes.length; i++) {
      var ret = cf(urls[i], sizes[i]);
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

// performs the conditional formating experiment on a spreadsheet of size size
// with the url url
function cf(url, size) {
  var ss = SpreadsheetApp.openByUrl(url);
  var r = ss.getActiveSheet().getDataRange();
  var id = ss.getActiveSheet().getSheetId();
  var sheet = ss.getActiveSheet();
  
  var request = {
  "requests": [
    {
      "addConditionalFormatRule": {
        "rule": {
          "ranges": [
            {
              "sheetId": id,
              "startColumnIndex": 9,
              "endColumnIndex": 10,
            },
          ],
          "booleanRule": {
            "condition": {
              "type": "NUMBER_GREATER",
              "values": [
                {
                  "userEnteredValue": "2"
                }
              ]
            },
            "format": {
              "backgroundColor": {
                "blue": 0.5,
                "red": 0.5,
              }
            }
          }
        },
        "index": 0
      }
    }
  ]
  };
  
  var clearRequest = {
    "requests": [
    {
      "deleteConditionalFormatRule": {
        "index": 0,
        "sheetId": id
      } 
    }
  ]
  }
  
  var date = new Date();
  Sheets.Spreadsheets.batchUpdate(JSON.stringify(request), ss.getId());
  console.log(sheet.getRange(size, 9).getBackground());
  var endDate = new Date();
  ret = endDate.getTime() - date.getTime();
  console.log("time is " + ret);

  // clean up
  Sheets.Spreadsheets.batchUpdate(JSON.stringify(clearRequest), ss.getId());
  return ret;
}