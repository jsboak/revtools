function createNewTerritoryMap(e) {

  var territorySheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Territory Map");
  if(territorySheet != null) {
    Logger.log("Territory map already exists - adding to existing map");

    addDataToNewTerritoryMap(territorySheet, e); //Need to create new method for adding columns to existing sheet

  } else {

    Logger.log("Creating new territory map");

    var territoryMapSheet = createSheetForNewTerritory(e);

    addDataToNewTerritoryMap(territoryMapSheet, e);

    PropertiesService.getUserProperties().setProperty("territoryMapName", territoryMapSheet.getName());

  }

  return CardService.newActionResponseBuilder()
    .setNotification(CardService.newNotification()
        .setText("Retrieved your accounts from Salesforce"))
    .setNavigation(CardService.newNavigation().popCard())

    .build();

}

function addDataToNewTerritoryMap(territorySheet, e) {

  var territoryMatrix = [];
  var accounts = JSON.parse(retrieveAccountsOwnedByCurrentUser(e));
  var numColumns = e.formInputs.sfdc_territory_fields.length;
  var territoryMap = territorySheet.getRange(3,2,accounts.totalSize,numColumns);
  var accountNameMatrix = [];
  var accountIdMatrix = [];

  for(let j=0; j < accounts.totalSize; j++) {

    var accountRow = [];
    var accountId = accounts.records[j].Id;

    var accountNameField = [SpreadsheetApp.newRichTextValue().setText(accounts.records[j].Name).
      setLinkUrl(userProperties.getProperty(baseURLPropertyName) + "/lightning/r/Account/" + accountId + "/view").build()];
    var accountIdField = [accountId];

    for(i=0; i < numColumns; i++) {

      var accountField = e.formInputs.sfdc_territory_fields[i];
      if(accountField == "null") {
        accountField = "";
      }      
      accountRow.push(accounts.records[j][accountField]);
    }
    accountNameMatrix.push(accountNameField);
    accountIdMatrix.push(accountIdField);
    territoryMatrix.push(accountRow)
  }

  Logger.log(accountNameMatrix);
  var accountNameMap = territorySheet.getRange(3,1,accounts.totalSize,1);
  accountNameMap.setRichTextValues(accountNameMatrix);

  territorySheet.getRange(3,2,accounts.totalSize, numColumns+1).setHorizontalAlignment("center");
  territoryMap.setValues(territoryMatrix);

  var accountIdMap = territorySheet.getRange(3,26,accounts.totalSize,1);
  accountIdMap.setValues(accountIdMatrix);

  // territorySheet.autoResizeColumns(1,26);

  //   let accountName = accounts.records[i].Name;
  //   let accountId = accounts.records[i].Id;
  //   let accountRevenue = accounts.records[i][e.formInput.revenue];
  //   let accountScore = accounts.records[i][e.formInput.accountScore];
  //   let renewalDate = new Date(accounts.records[i][e.formInput.renewalDate]);
  //   const oneDay = 24 * 60 * 60 * 1000;
  //   var today = new Date();
  //   var daysToRenewal = Math.round((renewalDate - today) / oneDay);
  //   if (daysToRenewal < 0) {
  //     daysToRenewal = ""
  //   }

  //   var accountNameCell = territorySheet.getRange(rowCounter,1); //remove .getRange() calls from foor loop -> https://www.steegle.com/google-products/google-apps-script-faq
    
  //   var accountHyperLink = SpreadsheetApp.newRichTextValue()
  //       .setText(accountName)
  //       .setLinkUrl(userProperties.getProperty(baseURLPropertyName) + "/lightning/r/Account/" + accountId + "/view")
  //       .build();
  //   accountNameCell.setRichTextValue(accountHyperLink);

  //   var accountIdCell = territorySheet.getRange(rowCounter,26); //remove .getRange() calls from foor loop -> https://www.steegle.com/google-products/google-apps-script-faq
  //   accountIdCell.setValue(accountId);

  //   var accountRevenueCell = territorySheet.getRange(rowCounter,2); //remove .getRange() calls from foor loop -> https://www.steegle.com/google-products/google-apps-script-faq
  //   accountRevenueCell.setValue(accountRevenue);

  //   var accountRenewalCell = territorySheet.getRange(rowCounter,3); //remove .getRange() calls from foor loop -> https://www.steegle.com/google-products/google-apps-script-faq
  //   accountRenewalCell.setValue(daysToRenewal);

  //   var accountRenewalCell = territorySheet.getRange(rowCounter,4); //remove .getRange() calls from foor loop -> https://www.steegle.com/google-products/google-apps-script-faq
  //   accountRenewalCell.setValue(accountScore);

  //   if(accounts.records[i].Opportunities) {

  //     var openOppIdCell = territorySheet.getRange(rowCounter,25); //remove .getRange() calls from foor loop -> https://www.steegle.com/google-products/google-apps-script-faq
  //     var openOppId = accounts.records[i].Opportunities.records[0].Id;
  //     openOppIdCell.setValue(openOppId);
      
  //     var openOppCell = territorySheet.getRange(rowCounter,7); //remove .getRange() calls from foor loop -> https://www.steegle.com/google-products/google-apps-script-faq
  //     var openOppName = accounts.records[i].Opportunities.records[0].Name;
  //     var oppHyperLink = SpreadsheetApp.newRichTextValue()
  //       .setText(openOppName)
  //       .setLinkUrl(userProperties.getProperty(baseURLPropertyName) + "/lightning/r/Opportunity/" + accounts.records[i].Opportunities.records[0].Id + "/view")
  //       .build();

  //     openOppCell.setRichTextValue(oppHyperLink);

  //     var openOppNextStepCell = territorySheet.getRange(rowCounter,8); //remove .getRange() calls from foor loop -> https://www.steegle.com/google-products/google-apps-script-faq
  //     var openOppNextStep = accounts.records[i].Opportunities.records[0].NextStep;
  //     openOppNextStepCell.setValue(openOppNextStep);

  //   }

  //   rowCounter++;

  // }
  
  
}

var currentSfdcUser;

function getCurrentSfdcUser() {
  Logger.log("Retrieving currently logged in user");
  var currentSfdcUserRequest = salesforceEntryPoint("https://login.salesforce.com/services/oauth2/userinfo","get","",false);

  var currentSFDCUserResponse = JSON.parse(currentSfdcUserRequest);
  currentSfdcUser = currentSFDCUserResponse.user_id;

  //For testing:
  // Chris' userID
  currentSfdcUser = '0058Y00000CFhvzQAD'

  // Jakes's userID
  // var currentSfdcUser = '0058Y00000CFhwsQAD'

  // return currentSfdcUser;
}

function retrieveAccountsOwnedByCurrentUser(e) {

  if (currentSfdcUser == null) {
    getCurrentSfdcUser();
  }

  var oppQuery = "";
  if(e.formInput.include_open_opp_key) {
    oppQuery = `,+(SELECT+Opportunity.Id,+Opportunity.Name,+Opportunity.NextStep+FROM+Account.Opportunities+WHERE+IsClosed+=+FALSE+LIMIT+1)`;
  }

  var accountQuery = "SELECT+Name,Id,+";
  e.formInputs.sfdc_territory_fields.forEach(element =>
    accountQuery = accountQuery + element + ",+"
  );

  // var soql = `SELECT+Name+,+Id+,+${revenue}+,+${renewalDate}+,+${accountScore}${opp_query}+from+Account+WHERE+OwnerId='${currentSfdcUser}'`;
  accountQuery = accountQuery.substring(0, accountQuery.length-2) + `${oppQuery}` + `+from+Account+WHERE+OwnerId='${currentSfdcUser}'`;
  
  var getDataURL = '/services/data/v57.0/query/?q='+accountQuery;

  var accounts = salesforceEntryPoint(userProperties.getProperty(baseURLPropertyName) + getDataURL,"get","",false);
  Logger.log("Retrieved accounts from SFDC");

  return accounts;

}

function createSheetForNewTerritory(e) {

  var activeSpreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  var accountFields = getAccountFields();

  var territoryMap = activeSpreadsheet.getSheetByName("Territory Map");

  //WE MAY WANT TO REMOVE THIS. THIS DELETES AN EXISTING TERRITORY MAP IF ONE OF THE SAME NAME ALREADY EXISTS
  if (territoryMap != null) {
      activeSpreadsheet.deleteSheet(newSheetTab);
  }
  //Create the new sheet
  territoryMap = activeSpreadsheet.insertSheet();
  territoryMap.setName("Territory Map");
  territoryMap.setFrozenColumns(1); //Freeze the first column (account names)
  territoryMap.setFrozenRows(2);
  territoryMap.getRange("A1:Z1000").setWrapStrategy(SpreadsheetApp.WrapStrategy.WRAP); //Set all cells to use WRAP of text
  territoryMap.hideColumn(territoryMap.getRange("Z1")); //Hide the last column so that we can use it for SFDC Account ID
  territoryMap.hideRow(territoryMap.getRange("2:2"));
  territoryMap.setRowHeights(1, 1000, 30);
  territoryMap.getRange("A1:Z1000").setVerticalAlignment("middle");

  var territoryMatrix = [];
  var numColumns = e.formInputs.sfdc_territory_fields.length;
  territoryMap.setColumnWidths(1, numColumns+1, 200);
  
  var bold = SpreadsheetApp.newTextStyle().setBold(true).build();
  var headerRows = territoryMap.getRange(1,1,2,numColumns+1).setHorizontalAlignment("center").setTextStyle(bold);

  var headerRowLabels = ["Account Name"];
  var headerRowIds = ["Name"];

  for(i=0; i < numColumns; i++) {
    var element = e.formInputs.sfdc_territory_fields[i];

    // Logger.log(accountFields[element]);    
    headerRowLabels.push(accountFields[element].label)
    headerRowIds.push(element)
  };

  if(e.formInput.include_open_opp_key) {
    
    headerRows = territoryMap.getRange(1,1,2,numColumns+3).setHorizontalAlignment("center").setTextStyle(bold);

    headerRowLabels.push(["Opportunity Name"],["Opportunity Next Step"])
    headerRowIds.push(["opp-Name"],["opp-NextStep"])

    territoryMap.setColumnWidths(1, numColumns+1, 200);

  }

  territoryMatrix.push(headerRowLabels);
  territoryMatrix.push(headerRowIds);
  headerRows.setValues(territoryMatrix);

  return territoryMap;
}






















