var thresholdList = []

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
    .addItem("Greater Than",">",false)
    .addItem("Less Than","<",false)
    .addItem("Equal To","=",false)
    ))

  builder.addSection(
    //Add text field for threshold metric (e.g. 90) here -> Example threshold would be: when Days Until Renewal = 90
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

function setThreshold(e) {

  var rangeList = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet().getActiveRangeList().getRanges();
  var sheet = SpreadsheetApp.getActiveSheet();

  for (let i = 0; i < rangeList.length; i++) {

    var range = rangeList[i]

    for (let j = 0; j < range.getValues().length; j++) {

      Logger.log(range.getCell(j+1,1).getA1Notation());

      // thresholdList.push({"Cell":range.getCell(j+1,1).getA1Notation()})
      var row = range.getCell(j+1,1).getRow();
      var column = range.getCell(j+1,1).getColumn();
      var accountId = sheet.getRange(row,26).getValue();
      var accountName = sheet.getRange(row,1).getValue();
      var fieldName = sheet.getRange(1,column).getValue();
      var fieldId = sheet.getRange(2,column).getValue();
      thresholdList.push({"Account":accountName, "SFDCID":accountId, "fieldName":fieldName, "fieldId":fieldId})

    }
  }

  createThresholdMap();

}

function createThresholdMap() {

  var thresholdSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Configured Thresholds");

  //WE MAY WANT TO REMOVE THIS. THIS DELETES AN EXISTING TERRITORY MAP IF ONE OF THE SAME NAME ALREADY EXISTS
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

    thresholdSheet.getRange(1,3).setValue("Threshold")

    thresholdSheet.getRange(1,4).setValue("Notification Method")

    thresholdSheet.getRange(1,5).setValue("Threshold Description")

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
  
  Logger.log("First empty row: " + getFirstEmptyRowByColumnArray())
  const emptyRow = getFirstEmptyRowByColumnArray();

  for(i=0; i < thresholdList.length; i++) {

    sheet.getRange(emptyRow+i, 1).setValue(thresholdList[i].Account);
    sheet.getRange(emptyRow+i,26).setValue(thresholdList[i].SFDCID);
    sheet.getRange(emptyRow+i,2).setValue(thresholdList[i].fieldName);
    sheet.getRange(emptyRow+i,25).setValue(thresholdList[i].fieldId)


  }
}

function getFirstEmptyRowByColumnArray() {
  var spr = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Configured Thresholds");
  var column = spr.getRange('A:A');
  var values = column.getValues(); // get all data in one call
  var ct = 0;
  while ( values[ct] && values[ct][0] != "" ) {
    ct++;
  }
  return (ct+1);
}



