function createUpdateSfdcPull() {
  if(ScriptApp.getProjectTriggers().filter(t => t.getHandlerFunction() == "updateSheetFromSfdcPull").length == 0) {
    ScriptApp.newTrigger("updateSheetFromSfdcPull").forSpreadsheet(SpreadsheetApp.getActiveSpreadsheet()).onOpen().create();
  }
}

function updateSheetFromSfdcPull() {

  if (SpreadsheetApp.getActiveSpreadsheet().getActiveSheet().getName() == "Territory Map") {
    updateTerritoryMap();
  } else if (SpreadsheetApp.getActiveSpreadsheet().getActiveSheet().getName() == "Opportunities" ) {

    updateOpportunityReport()

  }
}

function updateOpportunityReport() {

  var pulledOppSfdcData = JSON.parse(retrievePullOppDataFromSfdc());
  var widgetRowMap = mapOppIdRows();
  var rowToUpdate;
  var OppId;
  var cellToUpdate;

  var opportunityReport = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Opportunities");
  var fieldIds = opportunityReport.getDataRange();

  Logger.log("Updating opportunities with pulled-SFDC data");

  for (let i = 1; i < 26; i++) { //Iterate over columns

    var fieldId = fieldIds.getCell(2,i).getValue().toString();

    if(fieldId != "") {

      for (let j=0; j < pulledOppSfdcData.totalSize; j++) {

        OppId = pulledOppSfdcData.records[j].Id;

        rowToUpdate = widgetRowMap[OppId];

        try {
          cellToUpdate = opportunityReport.getRange(rowToUpdate,i);
        } catch (error) {

          SpreadsheetApp.getActive().toast("Opportunity " + OppId + " not found or owned by current user.", "Error")
          return
        }

        if (fieldId == "Name") {

          var oppHyperLink = SpreadsheetApp.newRichTextValue()
              .setText(pulledOppSfdcData.records[j].Name)
              .setLinkUrl(userProperties.getProperty(baseURLPropertyName) + "/lightning/r/Opportunity/" + OppId + "/view")
              .build();
          cellToUpdate.setRichTextValue(oppHyperLink);

        } else {
          cellToUpdate.setValue(pulledOppSfdcData.records[j][fieldId]);
        }
      }
    }
  }
  SpreadsheetApp.getActive().toast("Opportunities updated from Salesforce data.", "Update", "3");
  
}

function updateTerritoryMap() {
  var pulledSfdcData = JSON.parse(retrievePullTerritoryDataFromSfdc());
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
        try {
          cellToUpdate = territoryMap.getRange(rowToUpdate,i);
        } catch (error) {
          SpreadsheetApp.getActive().toast("Account " + accountId + " not found or owned by current user.", "Error")
          return
        }

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
  SpreadsheetApp.getActive().toast("Territory Map updated from Salesforce data.", "Update", "3");
}

function retrievePullOppDataFromSfdc() {
  var opportunityReport = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Opportunities");

  var fields = opportunityReport.getRange("A2:2").getValues()[0];
  let uniqueFields = [...new Set(fields)];

  Logger.log(uniqueFields);

  if(currentSfdcUser == null) {
    getCurrentSfdcUser();
  }

  var oppQuery = "SELECT+Id,+";

  for (let i = 0; i < fields.length; i++) {
    
    fieldId = uniqueFields[i];

    if(fieldId != "" && fieldId != null) {

      Logger.log(uniqueFields[i]);

      oppQuery = oppQuery + fieldId + ",+";

    }
  }

  oppQuery = oppQuery.substring(0, oppQuery.length-2) + `+from+Opportunity+WHERE+OwnerId='${currentSfdcUser}'`;
  
  var getOppDataURL = '/services/data/v57.0/query/?q='+oppQuery;

  if(isTokenValid) {
    var oppSfdcData = salesforceEntryPoint(userProperties.getProperty(baseURLPropertyName) + getOppDataURL,"get","",false);
  }

  Logger.log("Retrieved Opportunity data from SFDC");

  return oppSfdcData;

}

function retrievePullTerritoryDataFromSfdc() {

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

  Logger.log("Retrieved Territory Map data from SFDC");

  return sfdcData;

}

function mapOppIdRows() {

  var opportunityReport = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Opportunities");
  var oppIds = opportunityReport.getDataRange();

  var oppIdRowMap = {}

  var numColumns = oppIds.getValues()[0].length;

  for (let i = 2; i < oppIds.getValues().length; i++) {

    oppId = oppIds.getCell(i+1,numColumns).getValue();

    oppIdRowMap[oppId] = i+1;

  }

  return oppIdRowMap;
}

















