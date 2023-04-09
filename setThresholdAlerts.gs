function setThreshold(e) {

  var rangeList = SpreadsheetApp.getActiveSheet().getActiveRangeList().getRanges();

  for (let i = 0; i < rangeList.length; i++) {

    var range = rangeList[i]

    for (let j = 0; j < range.getValues().length; j++) {

      Logger.log(range.getCell(j+1,1).getA1Notation());

    }

      // Logger.log(rangeList[i].get);
  }
  
  
}

function getCellFieldMapping(cellA1Notation) {

  

}