function goToThresholdBuilder(e) {
    try {       
          let nav = CardService.newNavigation().pushCard(thresholdBuilder(e));
            return CardService.newActionResponseBuilder()
            .setNavigation(nav)
            .build();

    }catch(e){
    Logger.log(e);
    }
}

function thresholdBuilder(e) {

  var title = 'Configure Thresholds';

  var builder = CardService.newCardBuilder()
    .setHeader(CardService.newCardHeader().setTitle(title))

  builder.addSection(CardService.newCardSection()
    .addWidget(CardService.newSelectionInput().setTitle("Threshold Inequality")
    .setFieldName("thresholdInequality")
    .setType(CardService.SelectionInputType.DROPDOWN)
    .addItem("Greater Than","Greater Than",false)
    .addItem("Less Than","Less Than",false)
    .addItem("Equal To","Equal To",false)
    )
    .addWidget(CardService.newTextInput()
      .setValue("10")
      .setHint("Integer Value")
      .setFieldName("thresholdValue")
      .setTitle("Threshold Value")
    )
    .addWidget(CardService.newTextInput()
      .setMultiline(true)
      .setTitle("Threshold Description")
      .setFieldName("thresholdDescription")
      .setHint("Example: Notify me when customer is 90 days from renewal.")
    )
  )

  builder.addSection(CardService.newCardSection()
    .setHeader("Notification Preferences")
    .addWidget(CardService.newSelectionInput().setTitle("Notification Method")
      .setFieldName("notificationMethod")
      .setType(CardService.SelectionInputType.DROPDOWN)
      .addItem("None","none",true)
      .addItem("E-mail","E-mail",false)
      .addItem("Slack (Coming soon!)","slack",false)
      .addItem("Calendar Event (Coming soon!)","calendar",false)
    )

  )

  builder.addSection(CardService.newCardSection()
    .addWidget(CardService.newButtonSet()
      .addButton(CardService.newTextButton()
        .setText('Set Thresholds')
        .setTextButtonStyle(CardService.TextButtonStyle.FILLED)
        .setOnClickAction(CardService.newAction().setFunctionName('addThresholdsFromTerritoryMap'))
        .setDisabled(false))))

  builder.addSection(CardService.newCardSection()
    .setHeader("Test Existing Thresholds")
    .addWidget(CardService.newDecoratedText()
        .setText("Test Thresholds")
        .setOnClickAction(CardService.newAction().setFunctionName('getThresholdValuesFromSfdc').setParameters({adHocInvocation: "adHoc"}))
        .setStartIcon(CardService.newIconImage().setIcon(CardService.Icon.PHONE))
        .setWrapText(true)
    )
  );

  return builder.build();

}

function getThresholdProperty() {

  if(userProperties.getProperty("thresholdJson")) {

    // Logger.log("Property exists: " + userProperties.getProperty("thresholdJson"));
    return JSON.parse(userProperties.getProperty("thresholdJson"));
  } else {

    return {};
  }
}

//Only used for testing!
function deleteThresholdJson() {
  userProperties.deleteProperty("thresholdJson");
}

function modifyThresholdsFromConfiguredThresholds() {

  Logger.log("Preparing to update Thresholds Property from modified Configured Thresholds sheet.")
  //Update thresholdJson property from changes that take place in the ThresholdMap sheet
  var activeSheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();

  var sheetValues = activeSheet.getDataRange().getValues();
  var thresholdJson = {};

  var configuredFieldColumn = sheetValues[0].indexOf("Configured Field");
  var thresholdConditionalColumn = sheetValues[0].indexOf("Threshold Conditional");
  var thresholdMetricColumn = sheetValues[0].indexOf("Threshold Metric");
  var notificationColumn = sheetValues[0].indexOf("Notification Method");
  var thresholdDescriptionColumn = sheetValues[0].indexOf("Threshold Description");
  var fieldIdColumn = sheetValues[0].indexOf("Configured Field ID");
  var accountIdColumn = sheetValues[0].indexOf("Account ID");
  var accountNameColumn = sheetValues[0].indexOf("Account Name");
  var currentValueColumn = sheetValues[0].indexOf("Current Value");
  var thresholdCrossedOnColumn = sheetValues[0].indexOf("Threshold Crossed On Date");

  for (let i = 1; i < sheetValues.length; i++) {

    try {
      var fieldName = sheetValues[i][configuredFieldColumn];
      var thresholdConditional = sheetValues[i][thresholdConditionalColumn];
      var thresholdValue = sheetValues[i][thresholdMetricColumn];
      var notificationMethod = sheetValues[i][notificationColumn];
      var thresholdDescription = sheetValues[i][thresholdDescriptionColumn];
      var fieldId = sheetValues[i][fieldIdColumn];
      var accountId = sheetValues[i][accountIdColumn];
      var accountName = sheetValues[i][accountNameColumn];
      var currentValue = sheetValues[i][currentValueColumn];
      var thresholdCrossedOn = sheetValues[i][thresholdCrossedOnColumn];

      if(thresholdJson[fieldId]) {
        thresholdJson[fieldId][accountId] = {"Account":accountName, 
          "fieldName":fieldName, 
          "thresholdInequality":thresholdConditional,
          "thresholdValue":thresholdValue,
          "notificationMethod":notificationMethod,
          "thresholdDescription":thresholdDescription,
          "currentValue":currentValue,
          "thresholdCrossedOn":thresholdCrossedOn
          };
      } else {
        thresholdJson[fieldId] = {};
        thresholdJson[fieldId][accountId] = {"Account":accountName, 
          "fieldName":fieldName, 
          "thresholdInequality":thresholdConditional,
          "thresholdValue":thresholdValue,
          "notificationMethod":notificationMethod,
          "thresholdDescription":thresholdDescription,
          "currentValue":currentValue,
          "thresholdCrossedOn":thresholdCrossedOn
        }
      }

    } catch(E) {
      Logger.log(E);
    }
  }

  userProperties.setProperty("thresholdJson", JSON.stringify(thresholdJson));

  Logger.log("Updated Threshold DB Table (thresholdJSON property) from Configured Thresholds sheet.");

}

function addThresholdsFromTerritoryMap(e) {

  if(isNaN(e.formInput.thresholdValue)) {
    return CardService.newActionResponseBuilder()
    .setNotification(CardService.newNotification()
    .setText("Threshold value must be a number."))
    .build(); 
  }

  var activeSheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  if(activeSheet.getName() == "Territory Map") {
    var rangeList = activeSheet.getActiveRangeList().getRanges();
  } else {
    return CardService.newActionResponseBuilder()
      .setNotification(CardService.newNotification()
      .setText("Please navigate to Territory Map sheet to configure thresholds."))
      .build();
  }

  var sheetValues = activeSheet.getDataRange().getValues();
  var lastRow = activeSheet.getDataRange().getLastRow();
  var thresholdJson = getThresholdProperty();

  for (let i = 0; i < rangeList.length; i++) {

    var range = rangeList[i]

    for (let j = 0; j < range.getValues().length; j++) {
      
      var row = range.getCell(j+1,1).getRow();

      if(row > lastRow) {
        break;
      } else if (row == 1.0 || row == 2.0) {
        continue;
      }

      var column = range.getCell(j+1,1).getColumn();

      try {
        var accountId = sheetValues[[row-1]][25];
        var accountName = sheetValues[[row-1]][0];
        var fieldName = sheetValues[[0]][column-1];
        var fieldId = sheetValues[[1]][column-1];
        var currentValue = sheetValues[[row-1]][column-1];

        if(thresholdJson[fieldId]) {

          thresholdJson[fieldId][accountId] = {"Account":accountName, 
            "fieldName":fieldName, 
            "thresholdInequality":e.formInput.thresholdInequality,
            "thresholdValue":e.formInput.thresholdValue,
            "notificationMethod":e.formInput.notificationMethod,
            "thresholdDescription":e.formInput.thresholdDescription,
            "currentValue":currentValue
          };

        } else {
          thresholdJson[fieldId] = {};
          thresholdJson[fieldId][accountId] = {"Account":accountName, 
            "fieldName":fieldName, 
            "thresholdInequality":e.formInput.thresholdInequality,
            "thresholdValue":e.formInput.thresholdValue,
            "notificationMethod":e.formInput.notificationMethod,
            "thresholdDescription":e.formInput.thresholdDescription,
            "currentValue":currentValue
          }
        }
      } catch(E) {
        Logger.log(E);
      }
    }
  }

  userProperties.setProperty("thresholdJson", JSON.stringify(thresholdJson));
  userProperties.setProperty("userEmail", SpreadsheetApp.getActiveSpreadsheet().getOwner().getEmail());
  userProperties.setProperty("userName", SpreadsheetApp.getActiveSpreadsheet().getOwner().getUsername())
  userProperties.setProperty("configThresholdsSheet", SpreadsheetApp.getActiveSpreadsheet().getUrl() + "#gid=" + activeSheet.getSheetId());
  userProperties.setProperty("spreadSheetUrl", SpreadsheetApp.getActiveSpreadsheet().getUrl());
  // Logger.log(JSON.parse(userProperties.getProperty("thresholdJson")));

  updateThresholdMap();

  createThresholdCheckTrigger();

  return CardService.newActionResponseBuilder()
    .setNotification(CardService.newNotification()
    .setText("Thresholds have been configured."))
    .setNavigation(CardService.newNavigation().popToRoot())
    .build();
}

function createThresholdCheckTrigger() {
  if(ScriptApp.getProjectTriggers().filter(t => t.getHandlerFunction() == "getThresholdValuesFromSfdc").length == 0) {
    ScriptApp.newTrigger("getThresholdValuesFromSfdc").timeBased().everyHours(1).create();
  }
}

function getUserName() {
  Logger.log(userProperties.getProperty("userName"));
}

function updateThresholdMap() {

  var thresholdSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Configured Thresholds");

  if (thresholdSheet != null) {
      
    Logger.log("Threshold sheet already exists, adding thresholds to existing sheet")

  } else {

    Logger.log("Threshold sheet doesn't yet exist, creating new sheet and adding thresholds");

    thresholdSheet = SpreadsheetApp.getActiveSpreadsheet().insertSheet();
    thresholdSheet.setName("Configured Thresholds");
    //Formatting

    var bold = SpreadsheetApp.newTextStyle().setBold(true).build();
    thresholdSheet.getRange("A1:1").setHorizontalAlignment("center").setTextStyle(bold); //Center text and bold text in top row
     //Freeze the first column (account names)
    thresholdSheet.hideColumn(thresholdSheet.getRange("Z1"));
    thresholdSheet.hideColumn(thresholdSheet.getRange("Y1")); //Hide the last column so that we can use it for SFDC Account ID
    thresholdSheet.setRowHeights(2, 500, 30);
    thresholdSheet.getRange("A1:Z1000").setVerticalAlignment("middle");

    //Add the columns to header row
    thresholdMap = thresholdSheet.getRange(1,1,1,8)
    thresholdMapMatrix = [["Account Name", "Configured Field", "Current Value", "Threshold Conditional","Threshold Metric","Notification Method","Threshold Description","Threshold Crossed On Date"]]
    thresholdMap.setValues(thresholdMapMatrix)
    // thresholdSheet.getRange(1,1).setValue("Account Name");

    thresholdSheet.setFrozenRows(1); //Freeze the top row (header row)

    thresholdSheet.getRange(1,25).setValue("Configured Field ID")
    thresholdSheet.getRange(1,26).setValue("Account ID")

    thresholdSheet.autoResizeColumns(1, 26);
    thresholdSheet.setFrozenColumns(1);
    thresholdSheet.getRange("A2:Z1000").setWrapStrategy(SpreadsheetApp.WrapStrategy.WRAP);

    Logger.log("Created Threshold Sheet");
  }

  updateThresholdSheetFromProperty();

}

function updateThresholdSheetFromProperty() {

  Logger.log("Adding Thresholds to sheet");

  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Configured Thresholds");
  var headerRow = sheet.getRange("A1:1").getValues()[0];
  var thresholdJson = JSON.parse(userProperties.getProperty("thresholdJson"));
  var thresholdMatrix = []
  var fieldIdMatrix = []
  var numRows=0;

  var accountIndex = headerRow.indexOf("Account Name")
  var fieldNameIndex = headerRow.indexOf("Configured Field")
  var currentValueIndex = headerRow.indexOf("Current Value")
  var thresholdInequalityIndex = headerRow.indexOf("Threshold Conditional")
  var thresholdValueIndex = headerRow.indexOf("Threshold Metric")
  var notificationMethodIndex = headerRow.indexOf("Notification Method")
  var thresholdDescriptionIndex = headerRow.indexOf("Threshold Description")
  var thresholdCrossedOnIndex = headerRow.indexOf("Threshold Crossed On Date")

  for (var sfdcField of Object.keys(thresholdJson)) {

    for (var sfdcAccountId of Object.keys(thresholdJson[sfdcField])) {

      var accountRow = Array.apply(null, Array(8))
      accountRow[accountIndex] = thresholdJson[sfdcField][sfdcAccountId]["Account"];
      accountRow[fieldNameIndex] = thresholdJson[sfdcField][sfdcAccountId]["fieldName"]
      accountRow[currentValueIndex] = thresholdJson[sfdcField][sfdcAccountId]["currentValue"]
      accountRow[thresholdInequalityIndex] = thresholdJson[sfdcField][sfdcAccountId]["thresholdInequality"]
      accountRow[thresholdValueIndex] = thresholdJson[sfdcField][sfdcAccountId]["thresholdValue"]
      accountRow[notificationMethodIndex] = thresholdJson[sfdcField][sfdcAccountId]["notificationMethod"]
      accountRow[thresholdDescriptionIndex] = thresholdJson[sfdcField][sfdcAccountId]["thresholdDescription"]
      accountRow[thresholdCrossedOnIndex] = thresholdJson[sfdcField][sfdcAccountId]["thresholdCrossedOn"]

      thresholdMatrix.push(accountRow)

      var fieldRow = [[sfdcField],[sfdcAccountId]]
      fieldIdMatrix.push(fieldRow);
      numRows++;

    }

  }
  Logger.log("Placing threshold values into sheet");
  sheet.getRange(2,1,numRows,8).setValues(thresholdMatrix);
  sheet.getRange(2,25,numRows,2).setValues(fieldIdMatrix);

  SpreadsheetApp.setActiveSheet(SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Configured Thresholds"));

}

function getFirstEmptyRowByColumnArray(sheetName) {

  var spr = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheetName);

  var column = spr.getRange('A:A');
  var values = column.getValues(); // get all data in one call
  var ct = 0;
  while ( values[ct] && values[ct][0] != "" ) {
    ct++;
  }
  return (ct+1);
}



