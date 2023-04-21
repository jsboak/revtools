function buildThresholdList() {

  var configuredThresholds = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Configured Thresholds");
  var territoryMap = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Territory Map").getDataRange();
  
  var accountRowMap = mapAccountIdRows();
  var fieldIdColumnMap = mapFieldIdColumns();

  var thresholds = configuredThresholds.getDataRange();

  for (let j = 1; j < thresholds.getValues().length; j++) {

    var accountId = thresholds.getCell(j+1,26).getValue();
    var fieldId = thresholds.getCell(j+1,25).getValue();

    var currentValue = territoryMap.getCell(accountRowMap[accountId], fieldIdColumnMap[fieldId]).getValue();
    Logger.log("Current value for " + accountId + " at " + fieldId + " is: " + currentValue);

    thresholds.getCell(j+1,7).setValue(currentValue); //TODO: move away from hardcoded column value

    var inequality = configuredThresholds.getRange(j+1,3).getValue().toString(); //move away from hardcoded column value
    var thresholdValue = configuredThresholds.getRange(j+1,4).getValue(); //move away from hardcoded column value

    if( (inequality == "Less Than" && currentValue < thresholdValue) || 
        (inequality == "Greater Than" && currentValue > thresholdValue) || 
        (inequality == "Equal To" && currentValue == thresholdValue)) {

      territoryMap.getCell(accountRowMap[accountId], fieldIdColumnMap[fieldId]).setBackground("red");

    };

  }
}

function mapAccountIdRows() {

  var territoryMap = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Territory Map");
  var accountIds = territoryMap.getDataRange();

  var accountIdRowMap = {}

  for (let i = 2; i < accountIds.getValues().length; i++) {

    accountId = accountIds.getCell(i+1,26).getValue();

    accountIdRowMap[accountId] = i+1;

  }

  return accountIdRowMap;
}

function mapFieldIdColumns() {

  var territoryMap = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Territory Map");
  var fieldIds = territoryMap.getDataRange();

  var fieldIdColumnMap = {}

  //iterate over columns
  for (let i = 1; i < 27; i++) {

    fieldId = fieldIds.getCell(2,i).getValue();

    fieldIdColumnMap[fieldId] = i;

  }

  return fieldIdColumnMap;
}

