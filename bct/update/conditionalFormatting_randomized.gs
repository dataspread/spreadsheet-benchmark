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
var EXPER_NAME = "cf tests";
// sheet name of results spreadsheet to be written to
var SHEET_NAME = "method 1";

// TODO: Change values in cf function

/* ============= END OF PARAMETERS ============= */


/*  Runs the experiment on spreadsheet of size `size` as specified by the mapping in `urls`.
    This is the main function to be called for running a trial of the experiment. */
function experiment(size) {
  var result = cf(urls[size]);
  var results_sheet = SpreadsheetApp.openByUrl(RESULTS_URL).getSheetByName(SHEET_NAME);
  writeToSheet(results_sheet, size, result);
}

/*  Writes the date, experiment name, size, and result(trial time) to a spreadsheet, 
    and highlights the background of the result. */
function writeToSheet(sheet, size, result) {
  var time = new Date();
  var lastRow = sheet.getLastRow() + 1;
  sheet.getRange(lastRow, 1).setValue(Utilities.formatDate(time, 'America/Chicago', 'MMMM dd, yyyy HH:mm:ss Z'));
  lastRow++;
  sheet.getRange(lastRow, 1).setValue(EXPER_NAME);
  lastRow++;
  sheet.getRange(lastRow, 1).setValue(size);
  lastRow++;
  sheet.getRange(lastRow, 1).setValue(result).setBackground("orange");
}


/*  Performs the conditional formatting experiment on a copy of the spreadsheet 
    specified by `url` */
function cf(url) {
  var ss = SpreadsheetApp.openByUrl(url);
  // perform experiment on copy of spreadsheet
  // copy is added to "Recent", not to location of original spreadsheet
  var ss = ss.copy(ss.getName() + "_" + Date.now());
  var id = ss.getActiveSheet().getSheetId();

  // conditional formatting rule for spreadsheet
  var request = { // replace
    "requests": [
      {
        "addConditionalFormatRule": {
          "rule": {
            /* FILL IN
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
            */
          },
          "index": 0 // REPLACE
        }
      }
    ]
  };

  var startDate = new Date();
  // update spreadsheet with conditional formatting rule
  Sheets.Spreadsheets.batchUpdate(JSON.stringify(request), ss.getId());
  var endDate = new Date();

  // clearing conditional formatting rule to reset spreadsheet state
  // needed if experiment is repeated and copies are not used
  var clearRequest = { // replace
    "requests": [
      {
        /* FILL IN
        "deleteConditionalFormatRule": {
            "index": 0,
            "sheetId": id
        } 
        */
      }
    ]
  };
  Sheets.Spreadsheets.batchUpdate(JSON.stringify(clearRequest), ss.getId());

  ret = endDate.getTime() - startDate.getTime();
  console.log("time is " + ret);
  return ret;
}