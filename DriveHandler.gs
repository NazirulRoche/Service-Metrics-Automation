class DriveHandler {

    listFilesInFolder() {
        //folderId = "1-8PKTNqaXiV_GNGK8UHZwa1YX-wlC31v"; // folder name: Service Metrics Reports (by Nazirul-Intern)
        var folderId = "1d4qcRm82KTp2lm-39BA22VH7315xAa9w"; // folder for testing since files cant be deleted in main folder
        var folder = DriveApp.getFolderById(folderId);
        var files = folder.getFiles();

        var filesDetails = {
              fileNames: [],
              fileIds: [],
        }

        while (files.hasNext()) {
          var file = files.next();
          //Logger.log('File Name: ' + file.getName());
          filesDetails.fileNames.push(file.getName());
          //Logger.log('File ID: ' + file.getId());
          filesDetails.fileIds.push(file.getId());
        }

        Logger.log(filesDetails.fileNames);
        Logger.log(filesDetails.fileIds);
        return filesDetails;
    }

    getSheetIdFromFileId(fileIds) {
        let totalFiles = fileIds.length;
        let spreadsheetId = [];
        let file = [];
        for (let i = 0; i < totalFiles; i++) {
            file.push(DriveApp.getFileById(fileIds[i]));
            var url = [file[i].getUrl()];
            var match = /spreadsheets\/d\/([a-zA-Z0-9-_]+)/.exec(url);//This part of the pattern is a capturing group enclosed in parentheses
            if (match && match[1]) {
                spreadsheetId.push(match[1]);
            } else {
                continue; // No spreadsheet ID found
            }
        }
        Logger.log("spreadsheet IDs: " + spreadsheetId);
        return spreadsheetId; // Return the matched spreadsheet ID
    }

    getDetailsFromFileName(fileNames) {
        let totalFiles = fileNames.length;
        const allRegions = ["APAC", "EMEA", "AMER"];
        let service = [];
        let region = [];
        let month = [];
        let year = [];
        for (let i = 0; i < totalFiles; i++) {
            let str = fileNames[i].split("_");
            service.push(str[0]);
            if (str[1] !== "Global") {
              region.push(str[1]);
            } else { // if region is global, insert array of regions into the array
              region.push(allRegions);
            }
            month.push(str[2]);
            year.push(str[3]);
        }
        //Logger.log(service);
        //Logger.log(month);

        return {
          service,
          region,
          month,
          year
        } 
    }

    matchServicesFromFilesWithSheet(fileNames) {
        const servicesFromFilesName = this.getDetailsFromFileName(fileNames).service;

        const referenceSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Services Codes");

        const serviceCodes = referenceSheet.getRange(2, 1, referenceSheet.getLastRow() - 1, 2).getValues();
        
        const serviceObject = {};
        for (let i  = 0; i < serviceCodes.length; i++) {
            serviceObject[serviceCodes[i][1]] = serviceCodes[i][0];
        }

        const serviceKeys = Object.keys(serviceObject);

        let services = [];

        // pushing property of service object into services array IF service is one of the keys of serviceObject
        servicesFromFilesName.forEach(function(service) {
            if (serviceKeys.includes(service)) {
                services.push(serviceObject[service]);
            }
        });

        //Logger.log("services: " + services);
        return services;     
    }
}