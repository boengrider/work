## CZ02_VASI

### There are 3 scripts in this project
  - Script A (Data processing)
  - Script B (CSV maker)
  - Script C (SAP uploader)
 
 ```diff
+ Data source 
Periodic weekly and monthly reports (xlsx) sent via email. Placed in the sharepoint library by flow

- Workflow
1. MS Flow
   - Script A
     - Powerapps
       - Script B
         - Script C
