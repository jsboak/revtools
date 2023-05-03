function creatonedittrigger(functionName) {

  try {
    SpreadsheetApp.getActiveSpreadsheet().getName();
  } catch(error) {
    SpreadsheetApp.getActiveSpreadsheet().renameActiveSheet("Untitled spreadsheet");
  }

  if(ScriptApp.getProjectTriggers().filter(t => t.getTriggerSourceId() == SpreadsheetApp.getActiveSpreadsheet().getId()).length == 0) {
    ScriptApp.newTrigger(functionName).forSpreadsheet(SpreadsheetApp.getActiveSpreadsheet()).onEdit().create();
  }
}

function logPush(e) {

  Logger.log("Event: " + JSON.stringify(e));
  if(isTokenValid()) {
    if (e.value || e.oldValue) { //If this was a single-cell change
      var columnNumber = e.range.columnStart;
      var sfdcFieldId = SpreadsheetApp.getActiveSheet().getRange(2,columnNumber).getValue();

      //{"columnEnd":2,"columnStart":2,"rowEnd":4,"rowStart":4}
      if(sfdcFieldId.toString().startsWith("opp-")) {

        sfdcFieldId = sfdcFieldId.toString().substring(4);

        Logger.log("Pushing update to opportunity");
        var accountRow = e.range.rowStart;
        var opportunityId = SpreadsheetApp.getActiveSheet().getRange(accountRow,25).getValue();
        var updateAccountURL = userProperties.getProperty(baseURLPropertyName) + "/services/data/v57.0/sobjects/Opportunity/" + opportunityId;
        var patchPayload = {}
        patchPayload[sfdcFieldId] =  e.value ? e.value : "";

        Logger.log(SFDChttpRequest(updateAccountURL, "patch", JSON.stringify(patchPayload), false));

      } else {
          var accountRow = e.range.rowStart;
          var accountId = SpreadsheetApp.getActiveSheet().getRange(accountRow,26).getValue();
          var updateAccountURL = userProperties.getProperty(baseURLPropertyName) + "/services/data/v57.0/sobjects/Account/" + accountId;

          var patchPayload = {}

          patchPayload[sfdcFieldId] =  e.value ? e.value : "";

          Logger.log("Pushing update to Account");

          Logger.log(SFDChttpRequest(updateAccountURL, "patch", JSON.stringify(patchPayload), false));
        }
      }
  } else {
    return onHomepage(e); 
  }
}





