var DATA_SHEET = "datasheet_url"
var sizes = [150, 6000, 10000, 20000, 30000, 40000, 50000, 60000, 70000, 80000, 90000];

// ---------------------------------------- FORMULA ----------------------------------------
var urls = ["url1", "url2", "..."];

// ---------------------------------------- NO FORMULA ----------------------------------------
//var urls = ["url1", "url2", "..."];

var EXPER_NAME = "fandr formula non-exist"
function loop() {
  console.log("new trial");
  var results1 = [];
  var results2 = [];
  var tenruns1 = []; // to average across 10 trials
  var tenruns2 = [];
  for (j = 0; j <= 10; j ++) {
    for (i = 0; i < sizes.length; i++) {
      var experiment_name = EXPER_NAME;
      var ret1 = fandr(sizes[i], urls[i], 'yaawwn', "woohoo");
      var ret2 = fandr(sizes[i], urls[i], "woohoo", 'yaawwn');
      results1.push(ret1);
      results2.push(ret2);
    }
    tenruns1.push(results1);
    tenruns2.push(results2);
    results1 = [];
    results2=[];
  }
  averageStats(tenruns1, EXPER_NAME);
  averageStats(tenruns2, EXPER_NAME + " back");
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

function fandr(size, url, find, repl) {
  console.log("Starting " + size);
  var ss=SpreadsheetApp.openByUrl(url);
  var sheet = ss.getActiveSheet();
  var rws=sheet.getLastRow();
  var i,j,a,find,repl;
  var date = new Date();
  sheet.getRange(1,18).setValue(date.getTime());
  
  var data = sheet.getRange(1,3,rws,3).getValues();
  for (i=0;i<rws;i++) {
       try {
         data[i][1]=data[i][1].replace(find,repl);
        }
       catch (err) {continue;}
  }
  
  sheet.getRange(1,3,rws,3).setValues(data);
  console.log("Ending " + size);
  var endDate = new Date();
  sheet.getRange(2,18).setValue(endDate.getTime());
  var time = sheet.getRange(2,18).getValue() - sheet.getRange(1,18).getValue()
  sheet.getRange(3,18).setValue(time);
  return time
}