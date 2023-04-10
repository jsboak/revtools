function createNewTerritoryMap(e) {

  var territoryMapSheet = createSheetForNewTerritory(e);

  addDataToNewTerritoryMap(territoryMapSheet, e);

  PropertiesService.getUserProperties().setProperty("territoryMapName", territoryMapSheet.getName());

  Logger.log(PropertiesService.getUserProperties().getProperty("territoryMapName"))

  return CardService.newActionResponseBuilder()
  .setNotification(CardService.newNotification()
      .setText("Retrieved your accounts from Salesforce"))
  .setNavigation(CardService.newNavigation().popCard())

  .build();

}

function addDataToNewTerritoryMap(territorySheet, e) {
  var accounts = JSON.parse(retrieveAccountsOwnedByCurrentUser(e));

  var rowCounter = 3;
  for(let i=0; i < accounts.totalSize; i++) {

    let accountName = accounts.records[i].Name;
    let accountId = accounts.records[i].Id;
    let accountRevenue = accounts.records[i][e.formInput.revenue];
    let renewalDate = new Date(accounts.records[i][e.formInput.renewalDate]);
    Logger.log("Renewal Date: " + renewalDate);
    const oneDay = 24 * 60 * 60 * 1000;
    var today = new Date();
    var daysToRenewal = Math.round((renewalDate - today) / oneDay);
    if (daysToRenewal < 0) {
      daysToRenewal = ""
    }

    let accountScore = accounts.records[i][e.formInput.accountScore];

    var accountNameCell = territorySheet.getRange(rowCounter,1); //Set account name in first column
    accountNameCell.setValue(accountName);

    var accountIdCell = territorySheet.getRange(rowCounter,26);
    accountIdCell.setValue(accountId);

    var accountRevenueCell = territorySheet.getRange(rowCounter,2);
    accountRevenueCell.setValue(accountRevenue);

    var accountRenewalCell = territorySheet.getRange(rowCounter,3);
    accountRenewalCell.setValue(daysToRenewal);

    var accountRenewalCell = territorySheet.getRange(rowCounter,4);
    accountRenewalCell.setValue(accountScore);

    rowCounter++;

  }
  
  territorySheet.autoResizeColumn(1);
}

function retrieveAccountsOwnedByCurrentUser(e) {

  var currentSfdcUserRequest = salesforceEntryPoint("https://login.salesforce.com/services/oauth2/userinfo","get","",false);

  var currentSFDCUserResponse = JSON.parse(currentSfdcUserRequest);
  var currentSfdcUser = currentSFDCUserResponse.user_id;

  //For testing:
  // Chris' userID
  var currentSfdcUser = '0058Y00000CFhvzQAD'

  // Jakes's userID
  // var currentSfdcUser = '0058Y00000CFhwsQAD'

  var revenue = e.formInput.revenue;
  var renewalDate = e.formInput.renewalDate;
  var accountScore = e.formInput.accountScore;

  var soql = `SELECT+name+,+Id+,+${revenue}+,+${renewalDate}+,+${accountScore}+from+Account+WHERE+OwnerId='${currentSfdcUser}'`;
  var getDataURL = '/services/data/v57.0/query/?q='+soql;

  var accounts = salesforceEntryPoint(userProperties.getProperty(baseURLPropertyName) + getDataURL,"get","",false);
  Logger.log(accounts)

  return accounts;

}

function createSheetForNewTerritory(e) {

  //Get the active spreadsheet (the whole doc, not the individual sheet)
  var activeSpreadsheet = SpreadsheetApp.getActiveSpreadsheet();

  //Get today's date
  var today = new Date();
  var dd = String(today.getDate()).padStart(2, '0');
  var mm = String(today.getMonth() + 1).padStart(2, '0'); //January is 0!
  var yyyy = today.getFullYear();
  
  //Format accordingly
  today = mm + '/' + dd + '/' + yyyy;

  var territoryMapName = "Territory Map";// + today;
  var newSheetTab = activeSpreadsheet.getSheetByName(territoryMapName);

  //WE MAY WANT TO REMOVE THIS. THIS DELETES AN EXISTING TERRITORY MAP IF ONE OF THE SAME NAME ALREADY EXISTS
  if (newSheetTab != null) {
      activeSpreadsheet.deleteSheet(newSheetTab);
  }
  //Create the new sheet
  newSheetTab = activeSpreadsheet.insertSheet();
  newSheetTab.setName(territoryMapName);

  //Formatting
  var territoryMap = activeSpreadsheet.getSheetByName(territoryMapName);
  
  territoryMap.getRange("A1:Z1000").setWrapStrategy(SpreadsheetApp.WrapStrategy.WRAP); //Set all cells to use WRAP of text

  var bold = SpreadsheetApp.newTextStyle().setBold(true).build();
  territoryMap.getRange("A1:1").setHorizontalAlignment("center").setTextStyle(bold); //Center text and bold text in top row
  territoryMap.setFrozenColumns(1); //Freeze the first column (account names)
  territoryMap.hideColumn(territoryMap.getRange("Z1")); //Hide the last column so that we can use it for SFDC Account ID
  territoryMap.setRowHeights(2, 500, 30);
  territoryMap.getRange("A1:Z1000").setVerticalAlignment("middle");

  //Add the columns to header row
  territoryMap.getRange(1,1).setValue("Account Name");
  territoryMap.getRange(2,1).setValue("Name");
  territoryMap.hideRow(territoryMap.getRange("2:2"));

  territoryMap.getRange(1,2).setValue("Revenue");
  territoryMap.getRange(2,2).setValue(e.formInput.revenue);

  territoryMap.getRange(1,3).setValue("Days Until Renewal");
  territoryMap.getRange(2,3).setValue(e.formInput.renewalDate);

  territoryMap.getRange(1,4).setValue("Account Score");
  territoryMap.getRange(2,4).setValue(e.formInput.accountScore);
  territoryMap.getRange("D2:D").setHorizontalAlignment("center");

  territoryMap.getRange(1,5).setValue("Days Since Last Meeting");
  territoryMap.getRange(1,6).setValue("License Count");
  territoryMap.getRange(1,7).setValue("Open Opportunity");
  territoryMap.getRange(1,8).setValue("Opportunity Next Step");
  territoryMap.getRange(1,9).setValue("Notes");

  territoryMap.setColumnWidth(9,500);
  territoryMap.setColumnWidth(8,500);
  territoryMap.setColumnWidth(3, 160);

  territoryMap.setFrozenRows(2); //Freeze the top row (header row)

  return territoryMap;
}






















