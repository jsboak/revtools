function createNewTerritoryMap(e) {

  var territorySheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Territory Map");
  if(territorySheet != null) {
    Logger.log("Territory map already exists - adding to existing map");

    var sheetArray = territorySheet.getDataRange().getValues();
    var firstEmptyHeader = sheetArray[0].indexOf("") + 1; //Plus one because array index starts at zero but column numbers start at 1
    var firstEmptyFirstRow = sheetArray[2].indexOf("") + 1;
    if(firstEmptyHeader > firstEmptyFirstRow) {
      var firstEmptyColumn = firstEmptyHeader;
    } else {
      var firstEmptyColumn = firstEmptyFirstRow;
    }

    Logger.log("firstEmptyColumn: " + firstEmptyColumn)
    addDataToNewTerritoryMap(territorySheet, e, true, firstEmptyColumn); //Need to create new method for adding columns to existing sheet
    addColumnsToExistingTerritoryMap(territorySheet, e, firstEmptyColumn);

  } else {

    Logger.log("Creating new territory map");

    var territoryMapSheet = createSheetForNewTerritory(e);

    addDataToNewTerritoryMap(territoryMapSheet, e, false, 2);

    PropertiesService.getUserProperties().setProperty("territoryMapName", territoryMapSheet.getName());

  }

  return CardService.newActionResponseBuilder()
    .setNotification(CardService.newNotification()
        .setText("Retrieved your accounts from Salesforce"))
    .setNavigation(CardService.newNavigation().popCard())

    .build();
}

function addColumnsToExistingTerritoryMap(territorySheet, e, firstEmptyColumn) {

  // var accountFields = getAccountFields();

  var numberOfFields = e.formInputs.sfdc_territory_fields.length;

  var columnMatrix = [];
  territorySheet.setColumnWidths(firstEmptyColumn, numberOfFields+1, 200);
  
  var bold = SpreadsheetApp.newTextStyle().setBold(true).build();
  var headerRows = territorySheet.getRange(1,firstEmptyColumn,2,numberOfFields).setHorizontalAlignment("center").setTextStyle(bold);

  var headerRowLabels = [];
  var headerRowIds = [];

  for(i=0; i < numberOfFields; i++) {
    var element = e.formInputs.sfdc_territory_fields[i].split(":");

    // Logger.log(accountFields[element]);    
    headerRowLabels.push(element[1])
    headerRowIds.push(element[0])
  };

  if(e.formInput.include_open_opp_key) {
    
    headerRows = territorySheet.getRange(1,firstEmptyColumn,2,numberOfFields+2).setHorizontalAlignment("center").setTextStyle(bold);

    headerRowLabels.push(["Opportunity Name"],["Opportunity Next Step"])
    headerRowIds.push(["opp-Name"],["opp-NextStep"])

    territorySheet.setColumnWidths(1, numberOfFields+2, 200);

  }

  columnMatrix.push(headerRowLabels);
  columnMatrix.push(headerRowIds);
  headerRows.setValues(columnMatrix);

  SpreadsheetApp.getActive().toast("Successfully added new fields to sheet.", "Update", "1.5");

}

function addDataToNewTerritoryMap(territorySheet, e, existing, firstEmptyColumn) {

  var territoryMatrix = [];
  var accounts = JSON.parse(retrieveAccountsOwnedByCurrentUser(e));
  var numColumns = e.formInputs.sfdc_territory_fields.length;

  if(existing == false) {
    var territoryMap = territorySheet.getRange(3,2,accounts.totalSize,numColumns);
  } else {

    // var sheetArray = territorySheet.getDataRange().getValues();
    // var firstEmptyColumn = sheetArray[0].indexOf("") + 1; //Plus one because array index starts at zero but column numbers start at 1
    // Logger.log(firstEmptyColumn);
    var territoryMap = territorySheet.getRange(3,firstEmptyColumn,accounts.totalSize,numColumns);
  }
  
  var accountNameMatrix = [];
  var accountIdMatrix = [];

  if (!e.formInput.include_open_opp_key) {

    for(let j=0; j < accounts.totalSize; j++) {

      var accountRow = [];
      var accountId = accounts.records[j].Id;

      var accountNameField = [SpreadsheetApp.newRichTextValue().setText(accounts.records[j].Name).
        setLinkUrl(userProperties.getProperty(baseURLPropertyName) + "/lightning/r/Account/" + accountId + "/view").build()];
      var accountIdField = [accountId];

      for(i=0; i < numColumns; i++) {

        var accountField = e.formInputs.sfdc_territory_fields[i].split(":")[0];
        if(accountField == "null") {
          accountField = "";
        }      
        accountRow.push(accounts.records[j][accountField]);
      }
      accountNameMatrix.push(accountNameField);
      accountIdMatrix.push(accountIdField);
      territoryMatrix.push(accountRow)
      
    }

  } else {

    var oppNameMap = []
    var oppNextStepMap = [];
    var oppIdMap = [];

    for(let j=0; j < accounts.totalSize; j++) {

      var accountRow = [];
      var accountId = accounts.records[j].Id;

      var accountNameField = [SpreadsheetApp.newRichTextValue().setText(accounts.records[j].Name).
        setLinkUrl(userProperties.getProperty(baseURLPropertyName) + "/lightning/r/Account/" + accountId + "/view").build()];
      var accountIdField = [accountId];

      for(i=0; i < numColumns; i++) {

        var accountField = e.formInputs.sfdc_territory_fields[i].split(":")[0];
        if(accountField == "null") {
          accountField = "";
        }      
        accountRow.push(accounts.records[j][accountField]);
      }
      accountNameMatrix.push(accountNameField);
      accountIdMatrix.push(accountIdField);
      territoryMatrix.push(accountRow)
      
      if(accounts.records[j].Opportunities) {
        var openOppName = accounts.records[j].Opportunities.records[0].Name;
        var openOppNextStep = accounts.records[j].Opportunities.records[0].NextStep;
        var openOppId = accounts.records[j].Opportunities.records[0].Id;
        var oppHyperLink = SpreadsheetApp.newRichTextValue()
          .setText(openOppName)
          .setLinkUrl(userProperties.getProperty(baseURLPropertyName) + "/lightning/r/Opportunity/" + accounts.records[j].Opportunities.records[0].Id + "/view")
          .build();
        oppNameMap.push([oppHyperLink]);
        oppNextStepMap.push([openOppNextStep]);
        oppIdMap.push([openOppId]);
      } else {
        oppNameMap.push([SpreadsheetApp.newRichTextValue().setText("").build()]);
        oppNextStepMap.push([""]);
        oppIdMap.push([""]);
      }
    }

    territorySheet.getRange(3,2,accounts.totalSize, numColumns+2).setHorizontalAlignment("center");
    var oppMatrix = territorySheet.getRange(3,firstEmptyColumn+numColumns,accounts.totalSize,1);
    oppMatrix.setRichTextValues(oppNameMap);
    var oppNextStepMatrix = territorySheet.getRange(3,firstEmptyColumn+numColumns+1,accounts.totalSize,1);
    oppNextStepMatrix.setValues(oppNextStepMap);
    var oppIdMatrix = territorySheet.getRange(3,25,accounts.totalSize,1);
    oppIdMatrix.setValues(oppIdMap);
  }

  territorySheet.getRange(3,2,accounts.totalSize, numColumns+1).setHorizontalAlignment("center");
  territoryMap.setValues(territoryMatrix).setHorizontalAlignment("center");
  var accountNameMap = territorySheet.getRange(3,1,accounts.totalSize,1);
  accountNameMap.setRichTextValues(accountNameMatrix);

  var accountIdMap = territorySheet.getRange(3,26,accounts.totalSize,1);
  accountIdMap.setValues(accountIdMatrix);

} 


var currentSfdcUser;

function getCurrentSfdcUser() {
  Logger.log("Retrieving currently logged in user");

  var currentSfdcUserRequest = salesforceEntryPoint("https://login.salesforce.com/services/oauth2/userinfo","get","",false);

  var currentSFDCUserResponse = JSON.parse(currentSfdcUserRequest);
  currentSfdcUser = currentSFDCUserResponse.user_id;

  //For testing:
  // Chris' userID
  // currentSfdcUser = '0058Y00000CFhvzQAD'

  return currentSfdcUser;
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
    accountQuery = accountQuery + element.split(":")[0] + ",+"
  );

  accountQuery = accountQuery.substring(0, accountQuery.length-2) + `${oppQuery}` + `+from+Account+WHERE+OwnerId='${currentSfdcUser}'`;
  
  var getDataURL = '/services/data/v57.0/query/?q='+accountQuery;

  var accounts = salesforceEntryPoint(userProperties.getProperty(baseURLPropertyName) + getDataURL,"get","",false);
  Logger.log("Retrieved accounts from SFDC");

  return accounts;

}

function createSheetForNewTerritory(e) {

  var activeSpreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  // var accountFields = getAccountFields();

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
    var element = e.formInputs.sfdc_territory_fields[i].split(":");

    // Logger.log(accountFields[element]);    
    headerRowLabels.push(element[1])
    headerRowIds.push(element[0])
  };

  if(e.formInput.include_open_opp_key) {
    
    headerRows = territoryMap.getRange(1,1,2,numColumns+3).setHorizontalAlignment("center").setTextStyle(bold);

    headerRowLabels.push(["Opportunity Name"],["Opportunity Next Step"])
    headerRowIds.push(["opp-Name"],["opp-NextStep"])

    territoryMap.setColumnWidths(1, numColumns+3, 200);

  }

  territoryMatrix.push(headerRowLabels);
  territoryMatrix.push(headerRowIds);
  headerRows.setValues(territoryMatrix);

  return territoryMap;
}






















