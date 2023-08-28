function getThresholdValuesFromSfdc(e) {

  var adHocInvocation;
  try {
    adHocInvocation = e.parameters["adHocInvocation"];
  } catch(error) {
    adHocInvocation ="";
  }

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

    Logger.log("Querying SFDC to get current values for checking thresholds.")
    var sfdcData = JSON.parse(salesforceEntryPoint(userProperties.getProperty(baseURLPropertyName) + getDataURL,"get","",false));
    // Logger.log(sfdcData)
  }

  checkThresholdValues(sfdcData, thresholdJson, adHocInvocation);

}

function findIndexByKey(list, key, value) {
  return list.findIndex(obj => obj[key] === value);
}



function checkThresholdValues(sfdcData, thresholdJson, adHocInvocation) {

  var today = new Date();
  var dd = String(today.getDate()).padStart(2, '0');
  var mm = String(today.getMonth() + 1).padStart(2, '0'); //January is 0!
  var yyyy = today.getFullYear();
  today = mm + '/' + dd + '/' + yyyy;
  var accountsInEmailBody = '';

  Logger.log("JSON: " + JSON.stringify(thresholdJson))

  for (const[sfdcField,accounts] of Object.entries(thresholdJson)) {

    for (const[accountId,thresholds] of Object.entries(accounts)) {

      var index = findIndexByKey(sfdcData.records, "Id", accountId);

      if(index == -1) {

        Logger.log("No accounts in thresholds owned by currently logged in user.")
        return 
      }
      
      var currentValue = sfdcData.records[index][sfdcField];
      var inequality = thresholds.thresholdInequality;
      var thresholdValue = thresholds.thresholdValue;
      var fieldName = thresholds.fieldName

      if( (inequality == "Less Than" && currentValue < thresholdValue) || 
        (inequality == "Greater Than" && currentValue > thresholdValue) || 
        (inequality == "Equal To" && currentValue == thresholdValue)) {

        if(!thresholdJson[sfdcField][accountId]["thresholdCrossedOn"] || adHocInvocation === "adHoc") {
          
          thresholdJson[sfdcField][accountId]["thresholdCrossedOn"] = today;
          thresholdJson[sfdcField][accountId]["currentValue"] = currentValue;

          if(thresholdJson[sfdcField][accountId]["notificationMethod"] == "E-mail") {

            accountsInEmailBody+= "Account: " + thresholds.Account + 
              "{{newline}}Field: " + fieldName + 
              "{{newline}}Threshold: " + inequality + " " + thresholdValue +
              "{{newline}}Current Value: " + currentValue + "{{newline}}{{newline}}";

          }
        }
      }
    }
    if(adHocInvocation === "adHoc") {
      updateThresholdSheetFromProperty();
    }
  }

  // Logger.log(thresholdJson)
  userProperties.setProperty("thresholdJson", JSON.stringify(thresholdJson));
  if (accountsInEmailBody != '') {

    Logger.log("Found accounts that crossed thresholds. Sending email.")
    sendEmail(accountsInEmailBody);
  }

}

function getThreshold() {
  Logger.log(userProperties.getProperty("thresholdJson"))
}

function sendEmail(accountsInEmailBody) {

  const template = HtmlService.createTemplateFromFile('emailNotificationTemplate');
  template.name = userProperties.getProperty("userName");
  template.accountsInEmailBody = accountsInEmailBody;
  template.territoryMapUrl = userProperties.getProperty("configThresholdsSheet");
  const htmlBody = template.evaluate().getContent().replace(/{{newline}}/g, '<br>');

  var emailOptions = {name: "SeeGlass.ai", noReply:true, htmlBody: htmlBody}
  MailApp.sendEmail(userProperties.getProperty("userEmail") ,"RevTools: Thresholds Crossed for Accounts", "", emailOptions)

}

function mapAccountIdRows() {

  var territoryMap = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Territory Map");
  var accountIds = territoryMap.getDataRange();

  var accountIdRowMap = {}

  var numColumns = accountIds.getValues()[0].length;

  for (let i = 2; i < accountIds.getValues().length; i++) {

    accountId = accountIds.getCell(i+1,numColumns).getValue();

    accountIdRowMap[accountId] = i+1;

  }

  return accountIdRowMap;
}

function mapFieldIdColumns() {

  var territoryMap = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Territory Map");
  
  var fieldIds = territoryMap.getDataRange();

  var fieldIdColumnMap = {}

  var numColumns = fieldIds.getLastColumn()

  //iterate over columns
  for (let i = 1; i < numColumns+1; i++) {

    fieldId = fieldIds.getCell(2,i).getValue();

    fieldIdColumnMap[fieldId] = i;

  }

  return fieldIdColumnMap;
}

