function retrieveDataFromSfdc() {

  var territoryMap = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Territory Map");
  var fieldIds = territoryMap.getDataRange();

  if(currentSfdcUser == null) {
    getCurrentSfdcUser();
  }

  var oppQuery = "";
  var accountQuery = "SELECT+";
  
  //iterate over columns
  for (let i = 1; i < 26; i++) {

    fieldId = fieldIds.getCell(2,i).getValue().toString();

    if(fieldId != "") {

      if(fieldId.startsWith("opp-")) {
        fieldId = "Opportunity." + fieldId.substring(4);
        oppQuery = oppQuery + fieldId + ",+";
      } else {

        accountQuery = accountQuery + fieldId + ",+";

      }
    }    
  }

  if(oppQuery != "") {
    oppQuery = ",+(SELECT+" + oppQuery.substring(0,oppQuery.length-2) + "+FROM+Account.Opportunities+WHERE+IsClosed+=+FALSE+LIMIT+1)";

  }

  accountQuery = accountQuery.substring(0, accountQuery.length-2) + `${oppQuery}` + `+from+Account+WHERE+OwnerId='${currentSfdcUser}'`;

  Logger.log("accountQuery: " + accountQuery);

  var getDataURL = '/services/data/v57.0/query/?q='+accountQuery;
  var sfdcData = salesforceEntryPoint(userProperties.getProperty(baseURLPropertyName) + getDataURL,"get","",false);

  Logger.log(sfdcData);
  
}



















