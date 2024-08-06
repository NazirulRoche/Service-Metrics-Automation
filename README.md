# Apps Script Automation for QBR Service Metrics
<div align = "center">
  <img src = "images/Apps_Script_Automation.png">
</div>

<p align = "center">
  Automation Count Feature in Google Sheets
</p>

## Project Description
This project is an automation feature for counting process for QBR Service Metrics. It utilizes Google Apps Script (GAS) as the scripting tool, Google Sheets as data storage, and Google Drive as centralized files storage. It also integrates indirectly Google Sheets and Google Drive services via built-in APIs provided in GAS.

## Requirements
To contribute to this project, the following requirements must be met:
- Access to Google Sheets development & testing environment
- Access to Google Drive development & testing environment
- Basics about JavaScript and HTML

No installation or extension is required

## Usage Instructions
To begin developing this project, please follow the setup instructions below:
1. Open the Google Sheets Development & Testing environment
2. Click on the extension menu
3. Choose the Apps Script (a script called Code.gs will be auto created and deployed by default)
4. Have fun developing the script~

## Available Features
In GAS, there are multiple tools that can be utilized to maximize the efficiency of the script.

### 1. Execution Log
GAS offers a feature for the developers to check the variables' values by logging them inside the execution log using `Logger.log()`

Example: `Logger.log('Hello World')`;

<img src = "images/execution_log_example.png">

The result of the logging will appear in the execution log interface

This is useful for debugging process in case you prefer to debug manually without using debugger tool


### 2. Project History
GAS offers a small scale version control to save the script, track the changes, and revert previous version. 

  1. Saving the script via new deployment
     - To save the script, we can deploy it as a new version
       <img src = "images/manage_deployment.png" width = "800" height = "400">



       <img src = "images/create_new_version.png" width = "800" height = "500">

     - When creating new version, it is a good practice to include meaningful description in present tense form
        <img src = "images/deploy_new_version.png" width = "800" height = "600">

     For every big changes made into the script, it is a good practice to create new version
       
  
  1. Track the changes


  
  2. Restore the previous version
   

### 3. Services and Library
<img src = "images/GAS_menu.png" height = "450" width = "250">

GAS offers services as built-in APIs for developers to interact programmatically with the systems of Google products. Instead of using API keys containing the URL, these services can be called with the same syntax as classes

1. SpreadsheetApp service
  
   
2. DriveApp service



## Documentation
<div align = "center">
  <img src = "images/QBR_Service_Metrics_Class_Diagram_latest.png">
</div>
<p align = "center">
  Class diagram of the script
</p>

### DriveHandler Class
This is a utility class to handle the operation in Google Drive for:
  - Fetching files metadata (files' IDs and names)
  - Manipulating files' metadata

The table below shows the overview of the class methods functionality:


|       Class methods       |               Purpose               |           Requirement(s)           |            Outcome            |
| ------------------------- | ----------------------------------- | ---------------------------------- | ----------------------------- |
| listFilesInFolder() | listing all the files' IDs and names into arrays (IDs in one array, and names in another array) | No parameter | return Object called fileDetails containing array of fileIds and array of filenames |
| getDetailsFromFileName() | derive the service, region and month from the file name as a result from the standardized naming convention | needs to pass filenames (from listFilesInFolder) as an argument | return an Object of arrays of service, region, month and year (total 4 arrays) |
| matchServicesFromFilesWithSheet() | match the service abbreviation from filenames with the service full name | needs to pass filenames (from listFilesInFolder) as an argument | return an array of services' full names |

This class method is instantiated as an object in the main function only


### SheetHandler Class
This is a utility class to handle the operations in Google Sheets including:
- Extract the data
- Manipulate the data
- Clean the data
- Insert the data
- Insert comments

The table below shows the overview of the class methods functionality:


|       Class methods       |               Purpose               |           Requirement(s)           |            Outcome            |
| ------------------------- | ----------------------------------- | ---------------------------------- | ----------------------------- |
| findColumnHeader() | Find the column number based on given argument | Needs to pass the whole sheet and the exact column name (case sensitive) as arguments | return the column number (starting from 1)
| getCellsRangeInfoForID()  | Get the total number of rows and the the first index of the row based on the combination of service AND region | Needs to pass service (from matchServicesFromFilesWithSheet) and region (from getDetailsFromFileName) as arguments | return the total number of rows (cannot be 0) and the first index of the row (cannot be -1)
| getDateInfo() | clean the Date column due to inconsistency with the data type (some are string, some are Date objects) | Needs to pass the whole sheet, the exact column name (case sensitive), and current year as arguments | return an object of arrays containing month array and year array
| getCurrentTime() | Get the current time information (date, year, month, etc.) using Date library | No parameter | return an object of arrays containing current month array and current year array |
| determineReportType() | Return report type based on if else condition to determine which process (from 1 to 4) should be enforced to the report | Needs to pass service, region, month, and selectedMonth as the arguments | return a string of report type |


All the methods in this class are declared as static where they can be accessed anywhere throughout the script without having to instantiate the SheetHandler class

   
### ID Class
This class works like an abstract class (template class) in which it cannot be instantiated as an object but can be inherited to another class. Its fields and method can only be accessed by its child classes (process classes)


   
### ProcessOne Class

   
### ProcessTwo Class

    
### ProcessThree Class

    
### ProcessFour Class
    

## Resources
[Google Sheets Development & Testing Environment](https://docs.google.com/spreadsheets/d/1uLNC91rhGPvsknd7DNO8s35U0D_fkjl21RSy2EFNLLs/edit?gid=1866606411#gid=1866606411)

[Google Drive Development & Testing Environment](https://drive.google.com/drive/folders/1d4qcRm82KTp2lm-39BA22VH7315xAa9w)

[Google Sheets API documentation](https://developers.google.com/sheets/api/reference/rest)

[Google Drive API documentation](https://developers.google.com/drive/api/reference/rest/v3)

[GeoNames API documentation](https://www.geonames.org/export/web-services.html)




