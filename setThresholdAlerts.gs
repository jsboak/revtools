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
    // .addWidget(CardService.newImage().setImageUrl("https://freeiconshop.com/wp-content/uploads/edd/phone-flat.png"))
    .addWidget(CardService.newSelectionInput().setTitle("Notification Method")
      .setFieldName("notificationMethod")
      .setType(CardService.SelectionInputType.DROPDOWN)
      .addItem("None","none",true)
      .addItem("E-mail","E-mail",false)
      .addItem("Slack (Coming soon!)","slack",false)
      .addItem("Calendar Event (Coming soon!)","calendar",false)
    )
    .addWidget(CardService.newDecoratedText()
    // .setTopLabel("Highlight")
    .setText("Highlight Territory Map Cell")
    .setWrapText(true)
    .setSwitchControl(CardService.newSwitch()
        .setFieldName("highlight-cell")
        .setSelected(true)
        .setValue("true")
      )
    )
  )

  builder.addSection(CardService.newCardSection()
    .addWidget(CardService.newButtonSet()
      .addButton(CardService.newTextButton()
        .setText('Set Thresholds')
        .setTextButtonStyle(CardService.TextButtonStyle.FILLED)
        .setOnClickAction(CardService.newAction().setFunctionName('setThreshold'))
        .setDisabled(false))));

  return builder.build();

}
/*
New approach:
For threshold checking, create JSON:
  {
    
    ARR__c: {0018Y00002xtfWTQAY:1000,0018Y00002xtfWTQAY:1500,0018Y00002xtfWTQAY:1500,0018Y00002xtfWTQAY:2500},
    DaysSinceLastContacted__c:{0018Y00002xtfWTQAY:1000,0018Y00002xtfWTQAY:1500,0018Y00002xtfWTQAY:1500,0018Y00002xtfWTQAY:2500},
    DaysSinceOppUpdated__c: {0018Y00002xtfWTQAY:1000,0018Y00002xtfWTQAY:1500,0018Y00002xtfWTQAY:1500,0018Y00002xtfWTQAY:2500}
  }
Save to UserProperties. This will mean that thresholds set on different sheets will get added to this map.
On schedule, the ScriptApp will query salesforce for the fields (which are the Keys in the JSON) -> still use WHERE accounts owned by current user.
Loop through the JSON keys, loop through the field keys+values (which are the account ids and their respective thresholds), 
then check that value against the one from the Salesforce API response.
*/

var thresholdList = []

function getThresholdProperty() {

  if(userProperties.getProperty("thresholdJson")) {

    // Logger.log("Property exists: " + userProperties.getProperty("thresholdJson"));
    return JSON.parse(userProperties.getProperty("thresholdJson"));
  } else {

    return {};
  }
}

function setThreshold(e) {

  var activeSheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  if(activeSheet.getName() == "Territory Map") {
    var rangeList = activeSheet.getActiveRangeList().getRanges();
  } else {
    return CardService.newActionResponseBuilder()
      .setNotification(CardService.newNotification()
      .setText("Please navigate to Territory Map sheet to configure thresholds."))
      .build();
  }

  var sheet = SpreadsheetApp.getActiveSheet();
  var thresholdJson = getThresholdProperty();

  for (let i = 0; i < rangeList.length; i++) {

    var range = rangeList[i]

    for (let j = 0; j < range.getValues().length; j++) {

      var row = range.getCell(j+1,1).getRow();
      var column = range.getCell(j+1,1).getColumn();
      var accountId = sheet.getRange(row,26).getValue();
      var accountName = sheet.getRange(row,1).getValue();
      var fieldName = sheet.getRange(1,column).getValue();
      var fieldId = sheet.getRange(2,column).getValue();

      if(thresholdJson[fieldId]) {

        thresholdJson[fieldId][accountId] = e.formInput.thresholdValue;

      } else {
        thresholdJson[fieldId] = {};
        thresholdJson[fieldId][accountId] = e.formInput.thresholdValue;
      }
      
      thresholdList.push({
        "Account":accountName, 
        "SFDCID":accountId, 
        "fieldName":fieldName, 
        "fieldId":fieldId, 
        "thresholdInequality":e.formInput.thresholdInequality,
        "thresholdValue":e.formInput.thresholdValue,
        "notificationMethod":e.formInput.notificationMethod,
        "thresholdDescription":e.formInput.thresholdDescription
        })
    }
  }

  userProperties.setProperty("thresholdJson", JSON.stringify(thresholdJson));
  Logger.log(JSON.parse(userProperties.getProperty("thresholdJson")));

  createThresholdMap();

  return CardService.newActionResponseBuilder()
    .setNotification(CardService.newNotification()
    .setText("Thresholds have been configured."))
    .setNavigation(CardService.newNavigation().popToRoot())
    .build();
}

function createThresholdMap() {

  var thresholdSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Configured Thresholds");

  if (thresholdSheet != null) {
      
    Logger.log("Threshold sheet already exists, adding thresholds to existing sheet")

  } else {

    Logger.log("Threshold sheet doesn't yet exist, creating new sheet and adding thresholds");

    thresholdSheet = SpreadsheetApp.getActiveSpreadsheet().insertSheet();
    thresholdSheet.setName("Configured Thresholds");
    //Formatting
     //Set all cells to use WRAP of text

    var bold = SpreadsheetApp.newTextStyle().setBold(true).build();
    thresholdSheet.getRange("A1:1").setHorizontalAlignment("center").setTextStyle(bold); //Center text and bold text in top row
     //Freeze the first column (account names)
    thresholdSheet.hideColumn(thresholdSheet.getRange("Z1"));
    thresholdSheet.hideColumn(thresholdSheet.getRange("Y1")); //Hide the last column so that we can use it for SFDC Account ID
    thresholdSheet.setRowHeights(2, 500, 30);
    thresholdSheet.getRange("A1:Z1000").setVerticalAlignment("middle");

    //Add the columns to header row
    thresholdSheet.getRange(1,1).setValue("Account Name");

    thresholdSheet.setFrozenRows(1); //Freeze the top row (header row)

    thresholdSheet.getRange(1,2).setValue("Configured Field")

    thresholdSheet.getRange(1,3).setValue("Current Value");

    thresholdSheet.getRange(1,4).setValue("Threshold Conditional")

    thresholdSheet.getRange(1,5).setValue("Threshold Metric")

    thresholdSheet.getRange(1,6).setValue("Notification Method")

    thresholdSheet.getRange(1,7).setValue("Threshold Description")

    thresholdSheet.getRange(1,25).setValue("Configured Field ID")
    thresholdSheet.getRange(1,26).setValue("Account ID")

    thresholdSheet.autoResizeColumns(1, 26);
    thresholdSheet.setFrozenColumns(1);
    thresholdSheet.getRange("A2:Z1000").setWrapStrategy(SpreadsheetApp.WrapStrategy.WRAP);

  }

  addThresholdsToSheet(thresholdList);

}

function addThresholdsToSheet(thresholdList) {

  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Configured Thresholds");
  
  const emptyRow = getFirstEmptyRowByColumnArray("Configured Thresholds");

  for(i=0; i < thresholdList.length; i++) {

    sheet.getRange(emptyRow+i, 1).setValue(thresholdList[i].Account);
    sheet.getRange(emptyRow+i,26).setValue(thresholdList[i].SFDCID);
    sheet.getRange(emptyRow+i,2).setValue(thresholdList[i].fieldName);
    sheet.getRange(emptyRow+i,25).setValue(thresholdList[i].fieldId)
    sheet.getRange(emptyRow+i,4).setValue(thresholdList[i].thresholdInequality)
    sheet.getRange(emptyRow+i,5).setValue(thresholdList[i].thresholdValue)
    sheet.getRange(emptyRow+i,6).setValue(thresholdList[i].notificationMethod)
    sheet.getRange(emptyRow+i,7).setValue(thresholdList[i].thresholdDescription)

  }

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



