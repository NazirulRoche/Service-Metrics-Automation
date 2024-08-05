class ProcessFour extends ID {

    insertValuesIntoCells(referenceSpreadsheet, currentYear, numberOfRows, firstIndexOfServiceAndRegion, selectedMonth) {
        const referenceSheet = SheetHandler.getReferenceSheet(referenceSpreadsheet);
        const monthsFromReferenceSheet = SheetHandler.getDateInfo(referenceSheet, "opened_at", currentYear).monthsFromReferenceSheet;

        let count = 0;

        for (let i = 0; i < monthsFromReferenceSheet.length; i++) {
            if (monthsFromReferenceSheet[i] === selectedMonth) count += 1;
        }

        Logger.log(count);
        const definedCells = super.defineCells(numberOfRows, firstIndexOfServiceAndRegion, selectedMonth);
        definedCells.setValue(count);
    }

}