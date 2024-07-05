// composite key is combination of service, region and month
// the region in the report is global, needs to classify the city/country into region by ourself
// TODO: Discuss the classification with PO
class ProcessThree extends ID {

    classifyLocations(referenceSheet, currentYear, selectedMonth, region) {
        const locationColumn = SheetHandler.findColumnHeader(referenceSheet, "Location");
        let locationValues = referenceSheet.getRange(2, locationColumn, referenceSheet.getLastRow() - 1, 1).getValues().flat();

        const { monthsFromReferenceSheet, years } = SheetHandler.getDateInfo(referenceSheet, "Opened", currentYear);

        Logger.log("months: " + monthsFromReferenceSheet);

        let locations = [];
        for (let i = 0; i < years.length; i++) {
            if (years[i] === currentYear) {
                locations.push(locationValues[i]);
            }
        }

        Logger.log(locations.length);
        Logger.log(locations);
        
        const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Regions Classification");

        const apacColumn = SheetHandler.findColumnHeader(sheet, "APAC Region");
        const emeaColumn = SheetHandler.findColumnHeader(sheet, "EMEA Region");
        const amerColumn = SheetHandler.findColumnHeader(sheet, "AMER Region");
        const total = sheet.getLastRow() - 1; // minus header column

        let apacLocations = sheet.getRange(2, apacColumn, total, 1).getValues().flat();
        let emeaLocations = sheet.getRange(2, emeaColumn, total, 1).getValues().flat();
        let amerLocations = sheet.getRange(2, amerColumn, total, 1).getValues().flat();

        let countApac = 0;
        let countEmea = 0;
        let countAmer = 0;

        for (let i = 0; i < locations.length; i++) {
            if (apacLocations.includes(locations[i]) && monthsFromReferenceSheet[i] === selectedMonth) {
                countApac += 1;
            } else if (emeaLocations.includes(locations[i]) && monthsFromReferenceSheet[i] === selectedMonth) {
                countEmea += 1;
            } else if (amerLocations.includes(locations[i]) && monthsFromReferenceSheet[i] === selectedMonth) {
                countAmer += 1;
            }
        }
        
        Logger.log("APAC count for " + selectedMonth + ": " + countApac);
        Logger.log("EMEA count for " + selectedMonth + ": " + countEmea);
        Logger.log("AMER count for " + selectedMonth + ": " + countAmer);

        if (region === "APAC") {
            return countApac;
        } else if (region === "EMEA") {
            return countEmea;
        } else if (region === "AMER") {
            return countAmer;
        } else {
            Logger.log("invalid region");
        }
    }

    insertValuesByUpdatedRegion(count, service, selectedMonth, region) {
        let numOfRowsAndFirstIndex = SheetHandler.getCellsRangeInfoForID(service, region);
        let definedCells = super.defineCells(numOfRowsAndFirstIndex[0], numOfRowsAndFirstIndex[1], selectedMonth);
        definedCells.setValue(count);  
    }

}