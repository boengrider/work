## CZ02_VASI

<img src="images/algorithm(1).png">

### There are 3 scripts in this project
  - Script A (Data processing)
  - Script B (CSV maker)
  - Script C (SAP uploader)
 
 ```diff
+ Data source 
Periodic weekly and monthly reports (xlsx) sent via email. Placed in the sharepoint library by flow
- Workflow
1. MS Flow
  2. Script A
    3. Powerapps
      4. Script B
        5. Script C
1. Flow receives weekly/monthly report (.xlsx) files and places them in the sharepoint library
2. Script A searches through SP library and processes every file it discovers within this SP library. 
    Fileas are then moved to processed subdirectory.
    Processed data is uploaded to sharepoint list which servers as data source for powerapps
    
3. Powerapps is the frontend for users. User makes final decision in the powerapps. Individual 
    list item is marked 'Ready' for further processing
    
4. Script B periodically scans SP list for items marked 'Ready'. It produces output CSV files 
    which are placed in yet another SP library waiting for Script C
    Output file name is structured so that uploading Script C can perform apropriate action
    
5. Script C takes outpu .CSV files from the SP library and uploads them to SAP based on rules
    specified within configuration and encoded within file name
