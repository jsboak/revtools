function createUpdateSfdcPull() {
  if(ScriptApp.getProjectTriggers().filter(t => t.getHandlerFunction() == "updateSheetFromSfdcPull").length == 0) {
    ScriptApp.newTrigger("updateSheetFromSfdcPull").forSpreadsheet(SpreadsheetApp.getActiveSpreadsheet()).onOpen().create();
  }
}

function updateSheetFromSfdcPull() {

  var pulledSfdcData = JSON.parse(retrievePullDataFromSfdc());
  var widgetRowMap = mapAccountIdRows();
  var rowToUpdate;
  var accountId;
  var cellToUpdate;

  var territoryMap = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Territory Map");
  var fieldIds = territoryMap.getDataRange();

  Logger.log("Updating sheet with pulled-SFDC data");
  for (let i = 1; i < 26; i++) { //Iterate over columns

    var fieldId = fieldIds.getCell(2,i).getValue().toString();

    if(fieldId.startsWith("opp-") || fieldId != "") {

      for (let j=0; j < pulledSfdcData.totalSize; j++) {

        accountId = pulledSfdcData.records[j].Id;
        rowToUpdate = widgetRowMap[accountId];
        cellToUpdate = territoryMap.getRange(rowToUpdate,i);

        if(fieldId == "opp-Name" && pulledSfdcData.records[j].Opportunities != null) {
          
          updatedValue = pulledSfdcData.records[j].Opportunities.records[0].Name;

          var opportunityId = pulledSfdcData.records[j].Opportunities.records[0].Id;
          var opportunityHyperlink = SpreadsheetApp.newRichTextValue()
            .setText(updatedValue)
            .setLinkUrl(userProperties.getProperty(baseURLPropertyName) + "/lightning/r/Opportunity/" + opportunityId + "/view")
            .build();
          cellToUpdate.setRichTextValue(opportunityHyperlink);

        } else if (fieldId.startsWith("opp-") && pulledSfdcData.records[j].Opportunities) {

          cellToUpdate.setValue(pulledSfdcData.records[j].Opportunities.records[0][fieldId.substring(4)]);

        } else if (fieldId == "Name") {

          var accountHyperLink = SpreadsheetApp.newRichTextValue()
              .setText(pulledSfdcData.records[j].Name)
              .setLinkUrl(userProperties.getProperty(baseURLPropertyName) + "/lightning/r/Account/" + accountId + "/view")
              .build();
          cellToUpdate.setRichTextValue(accountHyperLink);

        } else {
          cellToUpdate.setValue(pulledSfdcData.records[j][fieldId]);
        }
      }
    }
  }
}

function retrievePullDataFromSfdc() {

  var territoryMap = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Territory Map");

  var fields = territoryMap.getRange("A2:2").getValues()[0];
  let uniqueFields = [...new Set(fields)];

  Logger.log(uniqueFields);

  if(currentSfdcUser == null) {
    getCurrentSfdcUser();
  }

  var oppQuery = "";
  var accountQuery = "SELECT+Id,+";
  
  //iterate over columns
  for (let i = 0; i < fields.length; i++) {

    
    fieldId = uniqueFields[i];

    if(fieldId != "" && fieldId != null) {

      Logger.log(uniqueFields[i]);

      if(fieldId.startsWith("opp-")) {
        fieldId = "Opportunity." + fieldId.substring(4);
        oppQuery = oppQuery + fieldId + ",+";
      } else {

        accountQuery = accountQuery + fieldId + ",+";

      }
    }    
  }

  if(oppQuery != "") {
    oppQuery = ",+(SELECT+Opportunity.Id,+" + oppQuery.substring(0,oppQuery.length-2) + "+FROM+Account.Opportunities+WHERE+IsClosed+=+FALSE+LIMIT+1)";

  }

  accountQuery = accountQuery.substring(0, accountQuery.length-2) + `${oppQuery}` + `+from+Account+WHERE+OwnerId='${currentSfdcUser}'`;

  var getDataURL = '/services/data/v57.0/query/?q='+accountQuery;

  if(isTokenValid) {
    var sfdcData = salesforceEntryPoint(userProperties.getProperty(baseURLPropertyName) + getDataURL,"get","",false);
  }

  Logger.log("Retrieved pull-data from SFDC");

  return sfdcData;

}



















