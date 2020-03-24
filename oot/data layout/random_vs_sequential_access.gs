var DATA_SHEET = "url";

var size = 80000;
var url = "url";

var EXPER_NAME = "80k row random access"
function loop() {

  var results = [];
  var tenruns = []; // to average across 10 trials
  for (z = 0; z <= 10; z ++) {
    //var trial = [sequential(size, url)];
    var trial = [random(size, url)];
    tenruns.push(trial);
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

function sequential(size, url) {
  var sheet = SpreadsheetApp.openByUrl(url).getActiveSheet();
  var start = new Date();  
  for (i = 1; i < size; i++) {
    var temp = sheet.getRange(1,i);
  }
  var end = new Date();
  
  return (end.getTime() - start.getTime());
}

function random(size, url) {
  var sheet = SpreadsheetApp.openByUrl(url).getActiveSheet();
  var rows = [];
  for (i = 1; i < size; i++) {
    rows.push(i);
  }
  var shuffled = shuffle(rows);
  var start = new Date();  
  for (j = 0; j < shuffled.length; j++) {
    var temp = sheet.getRange(1, shuffled[j]);
  }
  var end = new Date();
  
  return (end.getTime() - start.getTime());
}

// the worlds dumbest way of getting random numbers between 1 and size without duplicates
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