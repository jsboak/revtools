/*
thresholdJson DB Table:
{
  NumberOfEmployees={
    0018Y00002xtfWWQAY={thresholdInequality=Less Than, Account=Rapid API, fieldName=Employees, thresholdValue=1234567890, notificationMethod=E-mail, thresholdCrossOn=}, 
    0018Y00002xtfWTQAY={notificationMethod=E-mail, thresholdInequality=Less Than, Account=Grafana Labs, thresholdValue=1234567890, fieldName=Employees, thresholdCrossOn=}, 
    0018Y00002xtfWVQAY={fieldName=Employees, Account=Snyk.io, notificationMethod=E-mail, thresholdInequality=Less Than, thresholdValue=1234567890, thresholdCrossOn=05/11/2023}, 
    0018Y00002xtfWUQAY={Account=MongoDB, fieldName=Employees, thresholdInequality=Less Than, notificationMethod=E-mail, thresholdValue=1234567890, thresholdCrossOn=}
}, 
  ARR__c={
    0018Y00002xtfWXQAY={thresholdInequality=Greater Than, thresholdValue=123456789, fieldName=ARR, Account=Notion, notificationMethod=E-mail, thresholdCrossOn=}, 
    0018Y00002xtfWYQAY={thresholdValue=123456789, thresholdInequality=Greater Than, fieldName=ARR, notificationMethod=E-mail, Account=Plaid, thresholdCrossOn=}, 
    0018Y00002xtfWZQAY={thresholdValue=123456789, thresholdInequality=Greater Than, fieldName=ARR, notificationMethod=E-mail, Account=Calendly, thresholdCrossOn=05/5/2023}, 
    0018Y00002xtfWWQAY={Account=Rapid API, thresholdValue=123456789, fieldName=ARR, thresholdInequality=Greater Than, notificationMethod=E-mail, thresholdCrossOn=}
  }
}
*/
function getThresholdValuesFromSfdc() {

  //This will use the thresholdJson property to query SFDC to retrieve the current values of the SFDC fields
  var thresholdJson = getThresholdProperty();

  var accountQuery = "SELECT+Id,+";

  //iterate over SFDC Fields from Threshold Table
  for (const[sfdcField,accounts] of Object.entries(thresholdJson)) {

      accountQuery = accountQuery + sfdcField + ",+";

  }

  var currentSfdcUser = getCurrentSfdcUser();

  accountQuery = accountQuery.substring(0, accountQuery.length-2)+ `+from+Account+WHERE+OwnerId='${currentSfdcUser}'`;

  var getDataURL = '/services/data/v57.0/query/?q='+accountQuery;

  if(isTokenValid()) {
    var sfdcData = JSON.parse(salesforceEntryPoint(userProperties.getProperty(baseURLPropertyName) + getDataURL,"get","",false));
  }

  checkThresholdValues(sfdcData, thresholdJson);

}

function findIndexByKey(list, key, value) {
  return list.findIndex(obj => obj[key] === value);
}

function checkThresholdValues(sfdcData, thresholdJson) {

  var today = new Date();
  var dd = String(today.getDate()).padStart(2, '0');
  var mm = String(today.getMonth() + 1).padStart(2, '0'); //January is 0!
  var yyyy = today.getFullYear();
  today = mm + '/' + dd + '/' + yyyy;
  var accountsInEmailBody = '';

  for (const[sfdcField,accounts] of Object.entries(thresholdJson)) {

    for (const[accountId,thresholds] of Object.entries(accounts)) {

      var index = findIndexByKey(sfdcData.records, "Id", accountId);
      var currentValue = sfdcData.records[index][sfdcField];
      var inequality = thresholds.thresholdInequality;
      var thresholdValue = thresholds.thresholdValue;

      if( (inequality == "Less Than" && currentValue < thresholdValue) || 
        (inequality == "Greater Than" && currentValue > thresholdValue) || 
        (inequality == "Equal To" && currentValue == thresholdValue)) {

        thresholdJson[sfdcField][accountId]["thresholdCrossedOn"] = today;

        if(thresholdJson[sfdcField][accountId]["notificationMethod"] == "E-mail") {

          accountsInEmailBody+= "Account: " + thresholds.Account + " Field: " + sfdcField + " Current Value " + currentValue + " Threshold = " + inequality + " " + thresholdValue + "{{newline}}";

        }
      }
    }
  }

  // Logger.log(thresholdJson)
  if (accountsInEmailBody != '') {
    sendEmail(accountsInEmailBody);
  }

}

function sendEmail(accountsInEmailBody) {

  Logger.log("Accounts: " + accountsInEmailBody)

  const template = HtmlService.createTemplateFromFile('emailNotificationTemplate');
  template.name = userProperties.getProperty("userName");
  template.accountsInEmailBody = accountsInEmailBody;
  template.territoryMapUrl = userProperties.getProperty("configThresholdsSheet");
  const htmlBody = template.evaluate().getContent().replace(/{{newline}}/g, '<br>');

  var emailOptions = {name: "RevTools.io", noReply:true, htmlBody: htmlBody}
  MailApp.sendEmail(userProperties.getProperty("userEmail") ,"RevTools: Thresholds Crossed for Accounts", "", emailOptions)
  //This function will be used to send an email to the user when the threshold is crossed

}

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

    thresholds.getCell(j+1,3).setValue(currentValue); //TODO: move away from hardcoded column value

    var inequality = configuredThresholds.getRange(j+1,4).getValue().toString(); //move away from hardcoded column value
    var thresholdValue = configuredThresholds.getRange(j+1,5).getValue(); //move away from hardcoded column value

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

