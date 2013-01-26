
/**
* Creates a new sheet with prodtest summary by test failures across different instances 
* Prereq: To run this script, 
*         1. name each sheet with the instance name
*         2. Use the colum A for the test name
* look at https://docs.google.com/a/salesforce.com/spreadsheet/ccc?key=0AtKOKETWzOCOdDFQaXRxaWJKQ05ybWNrOTY3a0JUS1E#gid=0 for example
*/
function getTestFailuresByInstance() {
 
  var newSheet = SpreadsheetApp.getActiveSpreadsheet();
  //newSheet.getSheetByName("Summary");
  /*
  * If you get an error that says 'You already have a sheet with that name. Please try again', 
  * delete/rename the sheet named 'Summary' and run the script again
  */
  newSheet.insertSheet("Summary");
 
  // get test failures by instance and put it in a map
  var sheets = SpreadsheetApp.getActiveSpreadsheet().getSheets();
  Logger.log("No of sheets "+ sheets.length);
  var testMap = {};  // map-> {testname=insance1|instance2} Eg. {testCreatePage=NA1|NA2|NA3}
  for(var i = 0; i < sheets.length; i++) {
    var sheetName = sheets[i].getName();
    if(sheetName == "Summary")
      continue;
    Logger.log("---- " + sheetName + "----");
    var rowRange = sheets[i].getDataRange();
    var rangeValue = sheets[i].getDataRange().getValues();
    var numRows = rowRange.getNumRows();
    
    for(var j = 0; j < numRows; j++) {
      var testCase = rangeValue[j][0];
      if (testCase.indexOf("test") == 0) {
        if(testCase in testMap)
          testMap[testCase] = testMap[testCase] +"|"+sheetName;
        else
          testMap[testCase] = sheetName;
      }
    }
  }
  Logger.log(testMap);   // hit Ctrl+Enter to view the log
  
  // put the values back in the summary sheet
  newSheet.getRange("A1").setValue("TestCase");
  newSheet.getRange("A1").setFontWeight("bold")
  var count = 1;
  var colIndex = {}
  for(var i = 0; i < sheets.length; i++) {
    instName = sheets[i].getName();
    if (instName != "Summary" && instName != "Overview" ) {
      var rIndex = String.fromCharCode(65+count);
      colIndex[instName] = rIndex;
      newSheet.getRange(rIndex+"1").setValue(instName);
      newSheet.getRange(rIndex+"1").setFontWeight("bold")
      count ++;
    }
  }
  var row = 2;
  for(var test in testMap) {
    newSheet.getRange("A"+row).setValue(test);
    var instances = new String(testMap[test]).split("|");
    for (var i = 0; i < instances.length; i++) {
      newSheet.getRange(colIndex[instances[i]]+row).setValue("Y");
      newSheet.getRange(colIndex[instances[i]]+row).setBackgroundColor('#B03313');
    }
    row ++;    
  }
}
