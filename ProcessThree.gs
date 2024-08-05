// composite key is combination of service, region and month
class ProcessThree extends ID {

    // Apps Script does not support the private fields/methods, therefore use naming convention of '_' to indicate that its private
    // Private fields/methods are not accessible outside their container class
    _classifyLocationsFromGeoNamesAPI(location, regionClassificationSheet) {
        const regionCode = {
            'NA': 'AMER',
            'SA': 'AMER',
            'AS': 'APAC',
            'EU': 'EMEA',
            'AF': 'undefined' // africa is out of scope, therefore declared as undefined
        }

        const apacColumn = SheetHandler.findColumnHeader(regionClassificationSheet, "APAC Region");
        const emeaColumn = SheetHandler.findColumnHeader(regionClassificationSheet, "EMEA Region");
        const amerColumn = SheetHandler.findColumnHeader(regionClassificationSheet, "AMER Region");

        const username = 'nazirul_4129';  // need to change if the PIC for this project is leaving

        let totalApac = SheetHandler.getlastRowInSpecificColumn(regionClassificationSheet, apacColumn);
        let totalEmea = SheetHandler.getlastRowInSpecificColumn(regionClassificationSheet, emeaColumn);
        let totalAmer = SheetHandler.getlastRowInSpecificColumn(regionClassificationSheet, amerColumn);

        //Logger.log(`total amer: ${totalAmer}`);
        //Logger.log(`total emea: ${totalEmea}`);
        //Logger.log(`total apac: ${totalApac}`);

        let cityName = location;
        let url = 'http://api.geonames.org/searchJSON?q=' + encodeURIComponent(cityName) + '&maxRows=1&username=' + username;
        let response = UrlFetchApp.fetch(url);
        let cityData = JSON.parse(response.getContentText());

        cityData.geonames.forEach(function(city) {
            try {
                let geonameId = city.geonameId;
                
                let url = 'http://api.geonames.org/getJSON?geonameId=' + geonameId + '&username=' + username;
                let response = UrlFetchApp.fetch(url);
                let data = JSON.parse(response.getContentText());
                
                if (regionCode[data.continentCode] !== 'undefined') {
                  switch(regionCode[data.continentCode]) {
                        case 'AMER':
                            regionClassificationSheet.getRange(totalAmer + 1, amerColumn).setValue(location);
                            break;
                        case 'EMEA':
                            regionClassificationSheet.getRange(totalEmea + 1, emeaColumn).setValue(location);
                            break;
                        case 'APAC':
                            regionClassificationSheet.getRange(totalApac + 1, apacColumn).setValue(location);
                            break;
                  }                  
                } else {
                  //throw new Error('Continent information not available');
                  Logger.log("region is out of scope");
                }
            } catch (error) {
              Logger.log(error.toString());
              return null;
            }
        });

    }

    countLocations(referenceSpreadsheet, region) {
        const referenceSheet = referenceSpreadsheet.getSheets()[0];
        //const referenceSheet = SheetHandler.getReferenceSheet(referenceSpreadsheet);
        const locationColumn = SheetHandler.findColumnHeader(referenceSheet, "Location");
        let locations = referenceSheet.getRange(2, locationColumn, referenceSheet.getLastRow() - 1, 1).getValues().flat();
        
        const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Regions Classification");

        const apacColumn = SheetHandler.findColumnHeader(sheet, "APAC Region");
        const emeaColumn = SheetHandler.findColumnHeader(sheet, "EMEA Region");
        const amerColumn = SheetHandler.findColumnHeader(sheet, "AMER Region");
        let total = sheet.getLastRow() - 1; // minus header column
        
        if (total === 0) total = 1; // the total number of rows cannot be 0 (if it is 0, then there is no location inside the classification, which is unlikely in the main sheet)

        let classifiedLocations = sheet.getRange(2, apacColumn, total, 3).getValues().flat();

        let uniqueLocations = [...new Set(locations)];  // to eliminate duplicates in location column, use Set
        Logger.log(uniqueLocations); 

        // insert locations into classification IF it doesnt exist yet
        for (let i = 0; i < uniqueLocations.length; i++) {
            if (classifiedLocations.includes(uniqueLocations[i]) === false) {
                this._classifyLocationsFromGeoNamesAPI(uniqueLocations[i], sheet);
                total++;
            }; 
        }

        classifiedLocations = sheet.getRange(2, apacColumn, total, 3).getValues().flat();
        Logger.log("Updated Classified Locations: " + classifiedLocations);

        let apacLocations = sheet.getRange(2, apacColumn, total, 1).getValues().flat();
        let emeaLocations = sheet.getRange(2, emeaColumn, total, 1).getValues().flat();
        let amerLocations = sheet.getRange(2, amerColumn, total, 1).getValues().flat();

        let countApac = 0;
        let countEmea = 0;
        let countAmer = 0;

        Logger.log(`APAC Locations: ${apacLocations}`);
        Logger.log(`EMEA Locations: ${emeaLocations}`);
        Logger.log(`AMER Locations: ${amerLocations}`);

        // count based on region classification
        for (let i = 0; i < locations.length; i++) {
            if (apacLocations.includes(locations[i])) {
                countApac += 1;
            } else if (emeaLocations.includes(locations[i])) {
                countEmea += 1;
            } else if (amerLocations.includes(locations[i])) {
                countAmer += 1;
            }
        }

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
        Logger.log(`number of rows and first index: ${numOfRowsAndFirstIndex}`);
        let definedCells = super.defineCells(numOfRowsAndFirstIndex[0], numOfRowsAndFirstIndex[1], selectedMonth);
        definedCells.setValue(count);  
    }

}  

