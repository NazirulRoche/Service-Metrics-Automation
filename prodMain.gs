// declared globally
const monthNames = ["Jan", "Feb", "Mar", "Apr", "May", "Jun", "Jul", "Aug", "Sep", "Oct", "Nov", "Dec"];

let selectedMonth;

function showPrompt() {
    var htmlOutput = HtmlService.createHtmlOutputFromFile('dropdown')
        .setWidth(400)
        .setHeight(100);
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
        Utilities.sleep(1000);  // Pause execution for 1 second
        selectedMonth = PropertiesService.getScriptProperties().getProperty('selectedMonth');
    }
}

function clearSelectedMonth() {
  PropertiesService.getScriptProperties().deleteProperty('selectedMonth');
}

function main(r) { // r is the name of the specific region   
    /*
    const message = "Please choose a month (Jan/Feb/Mar/Apr/May/Jun/Jul/Aug/Sep/Oct/Nov/Dec): ";
    let monthPrompt = SpreadsheetApp.getUi().prompt(message);
    let selectedMonth = monthPrompt.getResponseText();
    */

    showPrompt();

    waitForUserSelection();

    const driveHandler = new DriveHandler();
    const { fileNames, fileIds } = driveHandler.listFilesInFolder();
    const services = driveHandler.matchServicesFromFilesWithSheet(fileNames);
    const { region: region, month: monthsFromFile } = driveHandler.getDetailsFromFileName(fileNames);

    const currentYear = SheetHandler.getCurrentTime().currentYear;

    for (let i = 0; i < region.length; i++) {
      if (Array.isArray(region[i])) {
          for (let j = 0; j < region[i].length; j++) {
            if (region[i][j] === r) region[i] = region[i][j];
          }
      }
    }

    // Filter file details based on selected region
    const filteredIndexes = region.reduce((acc, reg, idx) => {
        if (reg === r) {
            acc.push(idx); 
        }
        Logger.log(acc);
        return acc;
    }, []);

    const filteredServices = filteredIndexes.map(idx => services[idx]);
    const filteredMonths = filteredIndexes.map(idx => monthsFromFile[idx]);
    const filteredRegions = filteredIndexes.map(idx => region[idx]);
    const filteredSheetIds = filteredIndexes.map(idx => fileIds[idx]);

    // Iterate over filtered data for the selected region
    for (let i = 0; i < filteredSheetIds.length; i++) {
        const referenceSpreadsheet = SpreadsheetApp.openById(filteredSheetIds[i]);
        const referenceSheet = referenceSpreadsheet.getSheets()[0];

        const numberOfRowsAndFirstIndex = SheetHandler.getCellsRangeInfoForID(filteredServices[i], filteredRegions[i]);;

        let reportType = SheetHandler.determineReportType(filteredServices[i], filteredRegions[i], filteredMonths[i], selectedMonth);

         Logger.log("report type: " + reportType);

        switch(reportType) {
            case "first type":
                let processOne = new ProcessOne({service: filteredServices[i], region: filteredRegions[i], month: filteredMonths[i]});
                let {counter, countObsoleteCase} = processOne.countFromReferenceSheet(referenceSheet);
                let sortedValues = processOne.sortValuesBasedOnCurrentSheetData(counter, numberOfRowsAndFirstIndex[0], numberOfRowsAndFirstIndex[1]);
                processOne.insertValuesIntoMainSheet(sortedValues, numberOfRowsAndFirstIndex[0], numberOfRowsAndFirstIndex[1], selectedMonth);
                processOne.isThereObseleteCases(countObsoleteCase, numberOfRowsAndFirstIndex[0], numberOfRowsAndFirstIndex[1], selectedMonth);
                break;
            case "second type":
                let processTwo = new ProcessTwo({service: filteredServices[i], region: filteredRegions[i], month: filteredMonths[i]});
                processTwo.insertValuesIntoCells(referenceSheet, numberOfRowsAndFirstIndex[0], numberOfRowsAndFirstIndex[1], currentYear, selectedMonth);
                break;
            case "third type":
                let processThree = new ProcessThree({service: filteredServices[i], region: filteredRegions[i], month: filteredMonths[i]});
                const count = processThree.classifyLocations(referenceSheet, currentYear, selectedMonth, filteredRegions[i]);
                processThree.insertValuesByUpdatedRegion(count, filteredServices[i], selectedMonth, filteredRegions[i]);
                break;
            case "fourth type":
                let processFour = new ProcessFour({service: filteredServices[i], region: filteredRegions[i], month: filteredMonths[i]});
                processFour.insertValuesIntoCells(referenceSheet, currentYear, numberOfRowsAndFirstIndex[0], numberOfRowsAndFirstIndex[1], selectedMonth);
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
