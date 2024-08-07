// This script shows the flow of the automation count

// declared globally
const monthNames = ["Jan", "Feb", "Mar", "Apr", "May", "Jun", "Jul", "Aug", "Sep", "Oct", "Nov", "Dec"];

let selectedMonth;

function showPrompt() {
    var htmlOutput = HtmlService.createHtmlOutputFromFile('dropdown')
        .setWidth(400)
        .setHeight(125);
    SpreadsheetApp.getUi().showModalDialog(htmlOutput, 'Select Month');
}

function returnMonth(selectedMonth) {
    PropertiesService.getScriptProperties().setProperty('selectedMonth', selectedMonth);
    return selectedMonth;
}

function continueAfterSelection() {
    var selectedMonth = PropertiesService.getScriptProperties().getProperty('selectedMonth');
    // Continue your script here
    Logger.log("Script continues with selected month: " + selectedMonth);
}

function waitForUserSelection() {
    while (!selectedMonth) {
        Utilities.sleep(100);  // Pause execution for 1 second
        selectedMonth = PropertiesService.getScriptProperties().getProperty('selectedMonth');
    }
}

function clearSelectedMonth() {
  PropertiesService.getScriptProperties().deleteProperty('selectedMonth');
}

function main(r) { // r is the name of the specific region   

    showPrompt();

    waitForUserSelection();

    const driveHandler = new DriveHandler();
    const { fileNames, fileIds } = driveHandler.listFilesInFolder();
    const services = driveHandler.matchServicesFromFilesWithSheet(fileNames);
    const { region: regionsFromFile, month: monthsFromFile } = driveHandler.getDetailsFromFileName(fileNames);

    const currentYear = SheetHandler.getCurrentTime().currentYear;

    // for SIP report, in which the region is global, change the region based on user selection
    for (let i = 0; i < regionsFromFile.length; i++) {
      if (Array.isArray(regionsFromFile[i])) {
          for (let j = 0; j < regionsFromFile[i].length; j++) {
            if (regionsFromFile[i][j] === r) regionsFromFile[i] = regionsFromFile[i][j];
          }
      }
    }

    // Filter file details based on selected region
    const filteredIndexes = regionsFromFile.reduce((acc, reg, idx) => {
        if (reg === r) {
            acc.push(idx); 
        }
        Logger.log(acc);
        return acc;
    }, []);

    const filteredServices = filteredIndexes.map(idx => services[idx]);
    const filteredMonths = filteredIndexes.map(idx => monthsFromFile[idx]);
    const filteredRegions = filteredIndexes.map(idx => regionsFromFile[idx]);
    const filteredSheetIds = filteredIndexes.map(idx => fileIds[idx]);

    // Iterate over filtered data for the selected region
    for (let i = 0; i < filteredSheetIds.length; i++) {
        const referenceSpreadsheet = SpreadsheetApp.openById(filteredSheetIds[i]);

        const numberOfRowsAndFirstIndex = SheetHandler.getCellsRangeInfoForID(filteredServices[i], filteredRegions[i]);

        let reportType = SheetHandler.determineReportType(filteredServices[i], filteredRegions[i], filteredMonths[i], selectedMonth);

         Logger.log("report type: " + reportType);

        switch(reportType) {
            case "first type":
                let processOne = new ProcessOne({service: filteredServices[i], region: filteredRegions[i], month: filteredMonths[i]});
                let {counter, countObsoleteCase} = processOne.countFromReferenceSheet(referenceSpreadsheet);
                let sortedValues = processOne.sortValuesBasedOnCurrentSheetData(counter, numberOfRowsAndFirstIndex[0], numberOfRowsAndFirstIndex[1]);
                processOne.insertValuesIntoMainSheet(sortedValues, numberOfRowsAndFirstIndex[0], numberOfRowsAndFirstIndex[1], selectedMonth);
                processOne.isThereObseleteCases(countObsoleteCase, numberOfRowsAndFirstIndex[0], numberOfRowsAndFirstIndex[1], selectedMonth);
                break;
            case "second type":
                let processTwo = new ProcessTwo({service: filteredServices[i], region: filteredRegions[i], month: filteredMonths[i]});
                processTwo.insertValuesIntoCells(referenceSpreadsheet, numberOfRowsAndFirstIndex[0], numberOfRowsAndFirstIndex[1], currentYear, selectedMonth);
                break;
            case "third type":
                let processThree = new ProcessThree({service: filteredServices[i], region: filteredRegions[i], month: filteredMonths[i]});
                const count = processThree.countLocations(referenceSpreadsheet, filteredRegions[i]);
                processThree.insertValuesByUpdatedRegion(count, filteredServices[i], selectedMonth, filteredRegions[i]);
                break;
            case "fourth type":
                let processFour = new ProcessFour({service: filteredServices[i], region: filteredRegions[i], month: filteredMonths[i]});
                processFour.insertValuesIntoCells(referenceSpreadsheet, currentYear, numberOfRowsAndFirstIndex[0], numberOfRowsAndFirstIndex[1], selectedMonth);
                break;
            default:
                Logger.log("Process is not yet established or report is not related.");
                break;
        }
    }

    clearSelectedMonth();
}

function mainAPAC() {
    main("APAC");
}

function mainEMEA() {
    main("EMEA");
}

function mainAMER() {
    main("AMER");
}

function onOpen() {
    let ui = SpreadsheetApp.getUi();

    var menu = ui.createMenu("Automatic Count");

    menu.addItem("Count for APAC region", "mainAPAC");
    menu.addItem("Count for EMEA region", "mainEMEA");
    menu.addItem("Count for AMER region", "mainAMER");

    menu.addToUi();
}
