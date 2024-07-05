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
        var activeSheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
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

        if (service === "Navify Tumorboard" && region === "APAC" && month === selectedMonth) {
            return reportOne;
        } else if (service === "Navify Analytics" && region === "APAC") {
            return reportTwo;
        } else if (service === "SIP") {
            return reportThree;
        } else if (service === "ISIS") {
            return reportFour;
        } else if (service === "Navify Analytics" && region === "EMEA" && month === selectedMonth) { // report is similar to nTB APAC
            return reportOne;
        } else {
            return 0;
        }
    }

    
}