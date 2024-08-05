// Utility class: can be accessed freely by other classes
class SheetHandler {

    static findColumnHeader(sheet, keyword) {
        let currentValues = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues();

        for (var i = 0; i < currentValues.length; i++) {
          for (var j = 0; j < currentValues[i].length; j++) {
              if (currentValues[i][j] === keyword) {
                var index = j + 1;
              }
          }
        }
        return index;
    }

    static getCellsRangeInfoForID(service, region) {
        var activeSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("ALL DH Services");
        let dataRange = activeSheet.getDataRange();
        let values = dataRange.getValues();

        let serviceColumn = this.findColumnHeader(activeSheet, "Service") - 1;
        let regionColumn = this.findColumnHeader(activeSheet, "Region") - 1;

        let numberOfEntities = 0;
        let firstIndex = -1;

        for (let i = 0; i < values.length; i++) {
          if (values[i][serviceColumn] === service && values[i][regionColumn] === region) {
            numberOfEntities += 1;
            if (firstIndex === -1) {
              firstIndex = i + 1; // Update firstIndex only if it hasn't been set yet
            }
          }
        }

        let numOfRowsAndFirstIndex = [numberOfEntities, firstIndex];

        return numOfRowsAndFirstIndex;
    }

    static getDateInfo(referenceSheet, keyword, currentYear) {
        let column = SheetHandler.findColumnHeader(referenceSheet, keyword);
        let values = referenceSheet.getRange(2, column, referenceSheet.getLastRow(), 1).getDisplayValues();
        const totalValues = values.length - 1; // column header is not included
        let monthsFromReferenceSheet = [];

        //const currentYear = getCurrentTime().currentYear;

        //Logger.log(values);
        let years = [];

        // data manipulation process
        for (let i = 0; i < totalValues; i++) {
            var parts = values[i][0].split('/');
            var day = parseInt(parts[0], 10);
            var month = parseInt(parts[1], 10) - 1; // Subtract 1 because months are zero-indexed in JavaScript
            var year = parseInt(parts[2], 10);
            var date = new Date(year, month, day);

            years.push(year);
            
            if (date.getFullYear() === currentYear) {
                monthsFromReferenceSheet.push(monthNames[date.getMonth()]);
            };
        }

        //Logger.log(monthsFromReferenceSheet);
        return {
            monthsFromReferenceSheet,
            years
        };
    }

    static getCurrentTime() {
        // Create a new Date object for the current date and time
        var currentDate = new Date();
        
        // Get the numeric month (0 for January, 1 for February, etc.) and add 1 to make it human-readable
        var numericMonth = currentDate.getMonth() + 1;
        var currentYear = currentDate.getFullYear();

        for (let i = 1; i < monthNames.length + 1; i++) {
            if (numericMonth === i) var monthName = monthNames[i - 1]; // minus 1 because monthNames index start with 0
        }

        // Return an object containing both representations
        return {
          monthName: monthName,
          currentYear: currentYear
        };
    }

    static determineReportType(service, region, month, selectedMonth) {
        const reportOne = "first type";
        const reportTwo = "second type";
        const reportThree = "third type";
        const reportFour = "fourth type";

        // for reports that are filtered by month, must include month === selectedMonth condition

        if (service === "Navify Tumorboard" && region === "APAC" && month === selectedMonth) {
            return reportOne;
        } else if (service === "Navify Analytics" && region === "APAC") {
            return reportTwo;
        } else if (service === "SIP" && month === selectedMonth) {
            return reportThree;
        } else if (service === "ISIS") {
            return reportFour;
        } else if (service === "Navify Analytics" && region === "EMEA" && month === selectedMonth) { // report is similar to nTB APAC
            return reportOne;
        } else {
            return 0;
        }
    }

    static getReferenceSheet(referenceSpreadsheet) {
        let referenceSheets = referenceSpreadsheet.getSheets();
        
        let sheetNames = [];
        for (let i = 0; i < referenceSheets.length; i++) {
            sheetNames.push(referenceSheets[i].getName());
        }

        let sheetMonths = [];
        for (let i = 0; i < sheetNames.length; i++) {
            let str = sheetNames[i].split("_");
            sheetMonths.push(str[2]);
        }

        const latestMonthIndex = sheetMonths.reduce((acc, month, index) => {
            if (monthNames.indexOf(sheetMonths[acc]) < monthNames.indexOf(month)) acc = index;
            return acc;
        }, 0);

        Logger.log("latest month index: " + latestMonthIndex);

        let referenceSheet = referenceSpreadsheet.getSheets()[latestMonthIndex];

        let referenceSheetName = referenceSheet.getName();

        Logger.log(referenceSheetName);
        
        referenceSpreadsheet.rename(referenceSheetName);

        return referenceSheet;
    }

    // binary search may not be suitable for searching for non-empty cell (last cell)
    // TODO: research on better searching algorithm
    static getlastRowInSpecificColumn(sheet, columnNumber) {
        let maxRows = sheet.getMaxRows();
        let low = 2;
        let high = maxRows;
        
        // using binary search to find the last row of each region column in classification sheet
        while (low <= high) {
          let mid = Math.floor((low + high) / 2);
          let range = sheet.getRange(mid, columnNumber, high - mid + 1, 1);
          let values = range.getValues();
          
          if (values[0][0] === "" && values[values.length - 1][0] === "") {
            high = mid - 1;
          } else if (values[0][0] !== "") {
            low = mid + 1;
          } else {
            for (let i = values.length - 1; i >= 0; i--) {
              if (values[i][0] !== "") {
                return mid + i;
              }
            }
            high = mid - 1;
          }
        }
        //Logger.log(`last row: ${low - 1}`);
        return low - 1;  // Adjust by -1 to return the correct last row number
    }
    
}




