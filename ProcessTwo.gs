// for Report Structure similar to Navify Analytics APAC
// composite key is combination of service, region and month
// involves current year so if there is transition of year, it may not work (instead of currentYear, try to use reports' year)
class ProcessTwo extends ID {

    insertValuesIntoCells(referenceSpreadsheet, numberOfRows, firstIndexOfServiceAndRegion, currentYear, selectedMonth) {
        const referenceSheet = SheetHandler.getReferenceSheet(referenceSpreadsheet);
        const keyword = "Date/Time Opened";
        const { monthsFromReferenceSheet, years } = SheetHandler.getDateInfo(referenceSheet, keyword, currentYear);

        const statusColumn = SheetHandler.findColumnHeader(referenceSheet, "Status");

        let statusValues = referenceSheet.getRange(2, statusColumn, referenceSheet.getLastRow() - 1, 1).getValues().flat();

        let statusForCurrentYear = [];
        for (let i = 0; i < years.length; i++) {
            if (years[i] === currentYear) statusForCurrentYear.push(statusValues[i]);
        }

        //Logger.log(statusForCurrentYear);

        let count = 0;

        for (let i = 0; i < monthsFromReferenceSheet.length; i++) {
            if (monthsFromReferenceSheet[i] === selectedMonth) count++;
        }

        const obsoleteKeyword = "Obsolete";

        let obseleteCount = 0;

        for (let i = 0; i < statusForCurrentYear.length; i++) {
            if (statusForCurrentYear[i] === obsoleteKeyword && monthsFromReferenceSheet[i] === selectedMonth) obseleteCount++;
        }

        Logger.log("obsolete count: " + obseleteCount);
        const definedCells = super.defineCells(numberOfRows, firstIndexOfServiceAndRegion, selectedMonth);
    
        definedCells.setValue(count);
        Logger.log("defined cells notation: " + definedCells.getA1Notation());

        if (obseleteCount === 0) {
            return 0;
        }
        else {
            const definedCells = super.defineCells(numberOfRows, firstIndexOfServiceAndRegion, selectedMonth);
            const commentText = "Obsolete cases count: " + obseleteCount;
            definedCells.setComment(commentText);
        }
    }

    

}