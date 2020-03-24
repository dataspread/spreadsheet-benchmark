// previous url: https://docs.google.com/spreadsheets/d/1EFwjPEvK74Py3jWNer9rlE5JsSmwMHVGlGPUNuIBEiQ/edit#gid=0
var DATA_SHEET = "https://docs.google.com/spreadsheets/d/18m0aPe84ZH4_MOOsgyNhUCQPkTFMbgyDCjTpDMKAcdc/edit#gid=0";


// previous urls
// NO FORMULA ---------------------
//var size = 150;
//var url = "https://docs.google.com/spreadsheets/d/1uScL1qDTNruOIkwBMLd4yyyy92MewGjfmXLKKt3ajkU/edit#gid=1208735331"
//var size = 10000;
//var url = "https://docs.google.com/spreadsheets/d/1tlg22qrIJ47GqOL3T6Unr0OjNrJhrU99MElV1ItMJ6g/edit#gid=107740236"
//var size = 20000
//var url = "https://docs.google.com/spreadsheets/d/1vIdwCaXV2ESwnXBkUF-GXlqgNCx-FsUu1H-I7ViVVCY/edit#gid=1901691147";
//var size = 50000
//var url = "https://docs.google.com/spreadsheets/d/1Rk3Id238caghcqA3PtAbLdRToCrdAjgt1PAI4_s_Zec/edit#gid=408570086";
//var size = 80000;
//var url = "https://docs.google.com/spreadsheets/d/17beCIdzMnv4sYjy8wWhe303VQnXUKAVtji6O3prhpn0/edit#gid=2057009412";

// 20k, 50k, 80k
// 20k: https://docs.google.com/spreadsheets/d/1OJAKik4Ju3yq-u3ArmLkHOlr7bcbH16soWANFQ_Oktw/edit
// 50k: https://docs.google.com/spreadsheets/d/1SpF4psgD1UC26FRJYRcnI87JUib6bHJQb1hkkue7kL0/edit
// 80k: https://docs.google.com/spreadsheets/d/1Ii4WaFwAlbCTKawfx1NRW7uk6aWBygwrGDWPIye5-yE/edit
// 10k: "https://docs.google.com/spreadsheets/d/1kgKdaOFx9gDUziGIsPKcs8hF875-KbfKwkQCARH73-s/edit#gid=1708501637
//30k: https://docs.google.com/spreadsheets/d/1IaRR_SQt_MBDBQJSrmueU5ROC_N3-UEwrjslXovRRb0/edit#gid=1724899661
// 20k: https://docs.google.com/spreadsheets/d/1XAAEgkasuNgDbOSpLHXCJnmEXUobmkhsqTyVd1XrZ1E/edit#gid=249543903
var sizes = [20000] //10000, 20000, 30000
var urls = ["https://docs.google.com/spreadsheets/d/1XAAEgkasuNgDbOSpLHXCJnmEXUobmkhsqTyVd1XrZ1E/edit#gid=249543903"];

var EXPER_NAME = "airbnb random row fetching tests"
function loop() {

  var results = []; // for inital fandr
  var tenruns = []; // to average across 10 trials
  
  for (z = 0; z < 10; z ++) {
    for (i = 0; i < sizes.length; i++) {
      //var ret = sequential(sizes[i], urls[i]);
      var ret = random(sizes[i], urls[i]);
      results.push(ret);
    }
    tenruns.push(results);
    results = [];
  }
  averageStats(tenruns);
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
      sheet.getRange(lastRow, i+1).setValue(sizes[i]);
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

function sequential(size, url) {
  var sheet = SpreadsheetApp.openByUrl(url).getActiveSheet();
  
  var start = new Date();
  var result = 0;
  var resultCell = sheet.getRange(4,19) // single cell
  resultCell.setFormula("=COUNT(A2:R" + size + ")");
  result += resultCell.getValue();
 
  var end = new Date();
  
  return (end.getTime() - start.getTime());
}


function columnToLetter(column)
{
  var temp, letter = '';
  while (column > 0)
  {
    temp = (column - 1) % 26;
    letter = String.fromCharCode(temp + 65) + letter;
    column = (column - temp - 1) / 26;
  }
  return letter;
}

function random(size, url) {
  var sheet = SpreadsheetApp.openByUrl(url).getActiveSheet();
  var rows = [];
  var lastColumn = sheet.getLastColumn();
  for (i = 1; i <= lastColumn; i++) {
    rows.push(columnToLetter(i));
  }
  var shuffled = shuffle(rows);
  var result = 0;
  var start = new Date();  
  var resultCell = sheet.getRange(4,19)
  for (j = 0; j < shuffled.length; j++) {
    resultCell.setFormula("=COUNT("+shuffled[j] + "1:" + shuffled[j] + size+")");
    result += resultCell.getValue();
  }
  var end = new Date();
  
  return (end.getTime() - start.getTime());
}

function shuffle(arra1) {
    var ctr = arra1.length, temp, index;

    while (ctr > 0) {
        index = Math.floor(Math.random() * ctr);
        ctr--;
        temp = arra1[ctr];
        arra1[ctr] = arra1[index];
        arra1[index] = temp;
    }
    return arra1;
}
