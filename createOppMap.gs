function createNewOppMap(e) {

  var oppSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Opportunities");

  if(oppSheet != null) {
    Logger.log("Opportunities tab already exists - adding to existing map");

    var sheetArray = oppSheet.getDataRange().getValues();
    var firstEmptyHeader = sheetArray[0].indexOf("") + 1; //Plus one because array index starts at zero but column numbers start at 1
    var firstEmptyFirstRow = sheetArray[2].indexOf("") + 1;
    if(firstEmptyHeader > firstEmptyFirstRow) {
      var firstEmptyColumn = firstEmptyHeader;
    } else {
      var firstEmptyColumn = firstEmptyFirstRow;
    }

    addDataToNewOppMap(oppSheet, e, true, firstEmptyColumn); //Need to create new method for adding columns to existing sheet
    addColumnsToExistingOppMap(oppSheet, e, firstEmptyColumn);

  } else {

    Logger.log("Creating new Opp map");

    var oppMapSheet = createSheetForNewOpp(e);

    addDataToNewOppMap(oppMapSheet, e, false);

    PropertiesService.getUserProperties().setProperty("oppMapName", oppMapSheet.getName());

  }

  return CardService.newActionResponseBuilder()
    .setNotification(CardService.newNotification()
        .setText("Retrieved your Opportunities from Salesforce"))
    .setNavigation(CardService.newNavigation().popCard())

    .build();
}

function addColumnsToExistingOppMap(oppSheet, e, firstEmptyColumn) {

  var oppFields = getOppFields();

  var numberOfFields = e.formInputs.sfdc_opp_fields.length;

  var columnMatrix = [];
  oppSheet.setColumnWidths(firstEmptyColumn, numberOfFields+1, 200);
  
  var bold = SpreadsheetApp.newTextStyle().setBold(true).build();
  var headerRows = oppSheet.getRange(1,firstEmptyColumn,2,numberOfFields).setHorizontalAlignment("center").setTextStyle(bold);

  var headerRowLabels = [];
  var headerRowIds = [];

  for(i=0; i < numberOfFields; i++) {
    var element = e.formInputs.sfdc_opp_fields[i];

    // Logger.log(oppFields[element]);    
    headerRowLabels.push(oppFields[element].label)
    headerRowIds.push(element)
  };

  columnMatrix.push(headerRowLabels);
  columnMatrix.push(headerRowIds);
  headerRows.setValues(columnMatrix);

  SpreadsheetApp.getActive().toast("Successfully added new fields to sheet.", "Update", "1.5");

}

function addDataToNewOppMap(oppSheet, e, existing, firstEmptyColumn) {

  var oppMatrix = [];
  var opps = JSON.parse(retrieveOppsOwnedByCurrentUser(e));
  var numColumns = e.formInputs.sfdc_opp_fields.length;

  if(existing == false) {
    var oppMap = oppSheet.getRange(3,2,opps.totalSize,numColumns);
  } else {

    // var sheetArray = oppSheet.getDataRange().getValues();
    // var firstEmptyColumn = sheetArray[0].indexOf("") + 1; //Plus one because array index starts at zero but column numbers start at 1
    // Logger.log(firstEmptyColumn);
    var oppMap = oppSheet.getRange(3,firstEmptyColumn,opps.totalSize,numColumns);
  }
  
  var oppNameMatrix = [];
  var oppIdMatrix = [];

  for(let j=0; j < opps.totalSize; j++) {

    var oppRow = [];
    var oppId = opps.records[j].Id;

    var oppNameField = [SpreadsheetApp.newRichTextValue().setText(opps.records[j].Name).
      setLinkUrl(userProperties.getProperty(baseURLPropertyName) + "/lightning/r/Opportunity/" + oppId + "/view").build()];
    var oppIdField = [oppId];

    for(i=0; i < numColumns; i++) {

      var oppField = e.formInputs.sfdc_opp_fields[i];
      if(oppField == "null") {
        oppField = "";
      }
      if(oppField.includes(".")) {
        var splitElement = oppField.split(".");
        var objectName = splitElement[0];
        var fieldName = splitElement[1];
        if(opps.records[j][objectName] == null) {
          oppRow.push("");
        } else {

          oppRow.push(opps.records[j][objectName][fieldName])
        }
      } else {
        oppRow.push(opps.records[j][oppField]);
      }
      
    }
    oppNameMatrix.push(oppNameField);
    oppIdMatrix.push(oppIdField);
    oppMatrix.push(oppRow)
    
  }

  oppSheet.getRange(3,2,opps.totalSize, numColumns+1).setHorizontalAlignment("center");

  oppMap.setValues(oppMatrix).setHorizontalAlignment("center");
  var oppNameMap = oppSheet.getRange(3,1,opps.totalSize,1);
  oppNameMap.setRichTextValues(oppNameMatrix);

  var oppIdMap = oppSheet.getRange(3,26,opps.totalSize,1);
  oppIdMap.setValues(oppIdMatrix);

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

function retrieveOppsOwnedByCurrentUser(e) {

  if (currentSfdcUser == null) {
    getCurrentSfdcUser();
  }

  var oppQuery = "SELECT+Name,Id,+";
  for(i=0; i < e.formInputs.sfdc_opp_fields.length; i++) {
    // var element = JSON.parse(e.formInputs.sfdc_opp_fields[i]);
    var element = e.formInputs.sfdc_opp_fields[i]
    oppQuery = oppQuery + element + ",+"
  }

  oppQuery = oppQuery.substring(0, oppQuery.length-2) + `+from+Opportunity+WHERE+OwnerId='${currentSfdcUser}'`;
  
  var getDataURL = '/services/data/v57.0/query/?q='+oppQuery;

  var opps = salesforceEntryPoint(userProperties.getProperty(baseURLPropertyName) + getDataURL,"get","",false);

  Logger.log("Retrieved opps from SFDC");
  // Logger.log("opps: " + opps);

  return opps;

}

function createSheetForNewOpp(e) {

  var activeSpreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  var oppFields = getOpportunityFields();

  var oppMap = activeSpreadsheet.getSheetByName("Opp Map");

  //WE MAY WANT TO REMOVE THIS. THIS DELETES AN EXISTING TERRITORY MAP IF ONE OF THE SAME NAME ALREADY EXISTS
  if (oppMap != null) {
      activeSpreadsheet.deleteSheet(newSheetTab);
  }
  //Create the new sheet
  oppMap = activeSpreadsheet.insertSheet();
  oppMap.setName("Opportunities");
  oppMap.setFrozenColumns(1); //Freeze the first column (opp names)
  oppMap.setFrozenRows(2);
  oppMap.getRange("A1:Z1000").setWrapStrategy(SpreadsheetApp.WrapStrategy.WRAP); //Set all cells to use WRAP of text
  oppMap.hideColumn(oppMap.getRange("Z1")); //Hide the last column so that we can use it for SFDC Opp ID
  oppMap.hideRow(oppMap.getRange("2:2"));
  oppMap.setRowHeights(1, 1000, 30);
  oppMap.getRange("A1:Z1000").setVerticalAlignment("middle");

  var oppMatrix = [];
  var numColumns = e.formInputs.sfdc_opp_fields.length;
  oppMap.setColumnWidths(1, numColumns+1, 200);
  
  var bold = SpreadsheetApp.newTextStyle().setBold(true).build();
  var headerRows = oppMap.getRange(1,1,2,numColumns+1).setHorizontalAlignment("center").setTextStyle(bold);

  var headerRowLabels = ["Opportunity Name"];
  var headerRowIds = ["Name"];

  for(i=0; i < numColumns; i++) {

    var element = e.formInputs.sfdc_opp_fields[i];

    if(oppFields[element] == null) {

      headerRowLabels.push(element.replace("."," "));
      headerRowIds.push(element);

    } else {
      
      headerRowLabels.push(oppFields[element].label)
      headerRowIds.push(element)
    }   

  };

  oppMatrix.push(headerRowLabels);
  oppMatrix.push(headerRowIds);
  headerRows.setValues(oppMatrix);

  return oppMap;
}
