var DATA_SHEET = "datasheet_url"
var sizes = [150, 6000, 10000, 20000, 30000, 40000, 50000, 60000, 70000, 80000, 90000];

// ---------------------------------------- FORMULA ----------------------------------------
var urls = ["url1", "url2", "..."];

// ---------------------------------------- NO FORMULA ----------------------------------------
//var urls = ["url1", "url2", "..."];


var EXPER_NAME = "airbnb countif 90k sheet multi instance"
function loop() {

  var results1 = []; // for inital fandr
  var tenruns1 = []; // to average across 10 trials
  for (j = 0; j <= 10; j ++) {
    for (i = 0; i < f_instances.length; i++) {
      var ret1 = sum(size, url, f_instances[i]);
      results1.push(ret1);
    }
    console.log(results1);
    tenruns1.push(results1);
    results1 = [];
  }
  averageStats(tenruns1);
}

function averageStats(t) {
  // t is arry of 10 arrays, each size # experiments
  // remove min and max
  var results_sheet = SpreadsheetApp.openByUrl(DATA_SHEET).getActiveSheet();
  var perNumber = [];
  for (i = 0; i < t[0].length; i++) {
    perNumber.push([]);
  }
  for (i = 0; i < t.length; i++) {
    for (j = 0; j < t[i].length; j++) {
      perNumber[j].push(t[i][j]);
    }
  }

  var results = [];
  for (z = 0; z < perNumber.length; z++) {
    cur = perNumber[z];
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

function writeToSheet(sheet, results) {
  var time = new Date();
  var lastRow = sheet.getLastRow()+1;
  sheet.getRange(lastRow, 1).setValue(Utilities.formatDate(time, 'America/Chicago', 'MMMM dd, yyyy HH:mm:ss Z'));
  lastRow++;
  sheet.getRange(lastRow, 1).setValue(EXPER_NAME);
  lastRow++;
  for (i = 0; i < results.length; i++) {
      sheet.getRange(lastRow, i+1).setValue(f_instances[i]);
  }
  lastRow++;
  for (i = 0; i < results.length; i++) {
    sheet.getRange(lastRow, i+1).setValue(results[i]).setBackground("orange");
  }
}

function writeInter(sheet, results) {
  var lastRow = sheet.getLastRow()+1;
  for (i = 0; i < results.length; i++) {
    sheet.getRange(lastRow, i+1).setValue(results[i]);
  }
}

function sum(size, url, inst) {
    ret = [];
    var sheet = SpreadsheetApp.openByUrl(url).getActiveSheet();
    var oldval = sheet.getRange(2,10).getValue();
    var form_string = "=COUNTIF(J2:J" + size + ", 1)";
    var value_to_set = oldval == 1? 0:1;
    var data = sheet.getRange(1,18,inst,18).getValues();
    for (z=0;z<inst;z++) {    // insert 5 of same formula
      data[z][1]=form_string;
    }
    sheet.getRange(1,18,inst,18).setValues(data);   // set them all at once
    var res = sheet.getRange(1,18,inst,18).getValues();   // get them all again to force computation
  
    var date = new Date();  
    sheet.getRange(2,10).setValue(value_to_set); // change 1 value
    var res = sheet.getRange(1,18,inst,18).getValues();   // get them all again to force recomputation
    var endDate = new Date();
  
    sheet.getRange(2,10).setValue(oldval);
    // clean up, 0 out the formulas
    for (z=0;z<inst;z++) {
      data[z][1]="0";
    }
    ret = endDate.getTime() - date.getTime();
    return ret;
}

