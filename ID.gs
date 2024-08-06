// This works similar to abstract class (template class) since it is not instantiated but inherited by process
class ID {
    //declare as object so that the order that does not matter
    constructor({service, region, month}) {
      this.service = service;
      this.region = region || null;
      this.month = month || null;
    }

    defineCells(numberOfRows, firstIndexOfServiceAndRegion, month) {
      let sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("ALL DH Services");
      let startColumn = 6;    
      let endColumn = 12;

      let definedCellsForServiceAndRegion = sheet.getRange(firstIndexOfServiceAndRegion, startColumn, numberOfRows, endColumn);
      var offsetColumn;

      for (let i = 0; i < monthNames.length; i++) {
          offsetColumn = i;
          if (month === monthNames[i]) {
              var definedCells = definedCellsForServiceAndRegion.offset(0, offsetColumn, numberOfRows, 1);
          }
      }
      
      return definedCells;
    }

}