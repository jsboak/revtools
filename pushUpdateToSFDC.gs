function creatonedittrigger(funcname) {
  if(ScriptApp.getProjectTriggers().filter(t => t.getHandlerFunction() == funcname).length == 0) {
    ScriptApp.newTrigger(funcname).forSpreadsheet(SpreadsheetApp.getActive()).onEdit().create();
  }
}

function logPush(e) {

  Logger.log("Event: " + JSON.stringify(e));
  if(isTokenValid()) {
    if (e.value || e.oldValue) { //If this was a single-cell change

    //{"columnEnd":2,"columnStart":2,"rowEnd":4,"rowStart":4}
      
      var accountRow = e.range.rowStart;
      var accountId = SpreadsheetApp.getActiveSheet().getRange(accountRow,26).getValue();
      var updateAccountURL = userProperties.getProperty(baseURLPropertyName) + "/services/data/v57.0/sobjects/Account/" + accountId;

      var columnNumber = e.range.columnStart;

      var sfdcFieldId = SpreadsheetApp.getActiveSheet().getRange(2,columnNumber).getValue();

      var patchPayload = {}

      patchPayload[sfdcFieldId] =  e.value ? e.value : "";

      Logger.log(patchPayload);

      Logger.log(SFDChttpRequest(updateAccountURL, "patch", JSON.stringify(patchPayload), false));

    }
  } else {
    return onHomepage(e); 
  }
  
}