// nTB main process is done, add-on to count Obselete cases (also done)
// For reports that are similar to nTB
// Its composite key is combination of service, region, month, country and case origin
// Since country and case origin can directly be accessed from the reference sheet, its not included in the constructor
class ProcessOne extends ID {

    countFromReferenceSheet(referenceSpreadsheet) {
      const referenceSheet = referenceSpreadsheet.getSheets()[0];
      let countryColumn = SheetHandler.findColumnHeader(referenceSheet,"Country");
      let caseOriginColumn = SheetHandler.findColumnHeader(referenceSheet, "Case Origin");
      let statusColumn = SheetHandler.findColumnHeader(referenceSheet, "Status");
      let lastRow = referenceSheet.getLastRow() - 1; // total number of rows minus the header row

      let sourceCountry = referenceSheet.getRange(2, countryColumn, lastRow, 1);
      let sourceChannel = referenceSheet.getRange(2, caseOriginColumn, lastRow, 1);
      let status = referenceSheet.getRange(2, statusColumn, lastRow, 1);

      let sourceCountryValues = sourceCountry.getValues();
      let sourceChannelValues = sourceChannel.getValues();
      let statusValues = status.getValues().flat();
      Logger.log(statusValues);

      let countObsoleteCase = {};

      let counter = {};
      for (let i = 0; i < sourceCountryValues.length; i++) {
          counter[sourceCountryValues[i] + "_" + sourceChannelValues[i]] = 0;    
          countObsoleteCase[sourceCountryValues[i] + "_" + sourceChannelValues[i]] = 0; 
      };
      
      const keyword = "Obsolete";

      for (let i = 0; i < sourceCountryValues.length; i++) {
          let countryValue = sourceCountryValues[i][0];
          let channelValue = sourceChannelValues[i][0];
          counter[countryValue + '_' + channelValue]++;

          if (statusValues[i] === keyword) countObsoleteCase[countryValue + "_" + channelValue]++;
      }
      
      Logger.log(counter);
      Logger.log(countObsoleteCase);

      return {
          counter,
          countObsoleteCase
      };
    }     

    //dependent to getCellsRangeInfoForID (to access numberOfRows and firstIndexOfServiceAndRegion)
    // dependent to countFromReferenceSheet
    sortValuesBasedOnCurrentSheetData(values, numberOfRows, firstIndexOfServiceAndRegion) {
        let mainSheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
        const countryColumn = SheetHandler.findColumnHeader(mainSheet, "Country");
        const caseOriginColumn = SheetHandler.findColumnHeader(mainSheet, "Contact Channel");
        const countries = mainSheet.getRange(firstIndexOfServiceAndRegion, countryColumn, numberOfRows, 1).getValues();
        const caseOrigins = mainSheet.getRange(firstIndexOfServiceAndRegion, caseOriginColumn, numberOfRows, 1).getValues();

        let countryAndCaseOriginArray = [];

        for (let i = 0; i < countries.length; i++) {
            countryAndCaseOriginArray.push(countries[i] + "_" + caseOrigins[i]);
        }

        //Logger.log(countryAndCaseOriginArray);
        
        let keys = Object.keys(values);
        Logger.log("values: " + keys);

        let sortedValues = [];

        for (let i = 0; i < countryAndCaseOriginArray.length; i++) {
            if (keys.includes(countryAndCaseOriginArray[i])) {
              sortedValues.push(values[countryAndCaseOriginArray[i]]); 
            }
            else {
              sortedValues.push(0); 
            }
        }

        Logger.log("sorted values: " + sortedValues);
        return sortedValues;
    }

    insertValuesIntoMainSheet(sortedValues, numberOfRows, firstIndexOfServiceAndRegion, month) {
        const definedCells = super.defineCells(numberOfRows, firstIndexOfServiceAndRegion, month);
        for (let i = 0; i < sortedValues.length; i++) {
            definedCells.getCell(i + 1, 1).setValue(sortedValues[i]);
        }
    }

    isThereObseleteCases(obseleteCasesCount, numberOfRows, firstIndexOfServiceAndRegion, month) {
        let keys = Object.keys(obseleteCasesCount);
        let values = Object.values(obseleteCasesCount);

        const mainSheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
        const countryColumn = SheetHandler.findColumnHeader(mainSheet, "Country");
        const caseOriginColumn = SheetHandler.findColumnHeader(mainSheet, "Contact Channel");

        let cellsRange = mainSheet.getRange(firstIndexOfServiceAndRegion, 1, numberOfRows, mainSheet.getLastColumn());
        let cellsValues = cellsRange.getValues();
        
        let obseleteCasesArray = [];
        let countryAndCaseOrigin = [];

        for (let i = 0; i < values.length; i++) {
            if (values[i] !== 0) {
                countryAndCaseOrigin.push(keys[i]);
                obseleteCasesArray.push(values[i]);
            }
        }

        if (obseleteCasesArray.length === 0) {
            return 0;
        }
        else {
            let indexArray = [];
            let countryAndCaseOriginObject = {}
            for (let i = 0; i < obseleteCasesArray.length; i++) {
                countryAndCaseOriginObject[countryAndCaseOrigin[i].split("_")[0]] = countryAndCaseOrigin[i].split("_")[1];
            }

            Logger.log(countryAndCaseOriginObject);
            obseleteCasesArray = [];
            
            for (let i = 0; i < cellsValues.length; i++) {
                for (let key in countryAndCaseOriginObject) {
                    if (cellsValues[i][countryColumn - 1] === key && cellsValues[i][caseOriginColumn - 1] === countryAndCaseOriginObject[key]) {
                        let rowIndex = cellsRange.getCell(i + 1, countryColumn).getRow();
                        indexArray.push(rowIndex);
                    }
                }
                for (let key in obseleteCasesCount) {
                    //Logger.log(cellsValues[i][countryColumn - 1] + "_" + cellsValues[i][caseOriginColumn - 1]);
                    //Logger.log(key);
                    if ((cellsValues[i][countryColumn - 1] + "_" + cellsValues[i][caseOriginColumn - 1]) === key && obseleteCasesCount[key] !== 0) obseleteCasesArray.push(obseleteCasesCount[key]);
                }
            }
            Logger.log("obsolete cases array: " + obseleteCasesArray);
            Logger.log(indexArray);

            let definedCells = super.defineCells(numberOfRows, firstIndexOfServiceAndRegion, month);
        
            for (let i = 0; i < indexArray.length; i++) {
                let commentText = "Obselete cases count: " + obseleteCasesArray[i];
                let cell = definedCells.offset(indexArray[i] - firstIndexOfServiceAndRegion, 0, 1);
                cell.setComment(commentText);
            }
          }
      }

}