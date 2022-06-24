## Project CZ02_VASI
Series of scripts for determining the contract coverage and presenting data in consistent way to the sharepoint users
In this project we also make use of powerapps as a frontend to the users
Based on user final decision, data is further processed and script(s) finish by uploading .CSV file(s) to SAP

### There are 3 scripts in this project
| Script | Name | Description |
|---|---|---|
| Script A | ProcessInputUploadToSPList | Input data validation + upload to sharepoint |
| Script B | ProcessSPListMakeCSV | Data transormation based on configuration (.CSV creation) |
| Script C | UploadToSAP | Upload of .CSV to SAP (SM35). Verification and info logging |
 
  
 ---
 ### Workflow
 <picture>
  <img alt="Shows an illustrated sun in light color mode and a moon with stars in dark color mode." src="images/algorithm (1).png">
</picture>


### 1. Periodic weekly and monthly reports (.xlsx) arrive by e-mail
Placed by MS flow to the sharepoint source library for script A to discover and process

<picture>
  <img alt="Sharepoint source library" src="images/sp_source_library.PNG">
</picture>


### 2. Script A processes incoming reports
All monthly and weekly reports discovered by Script A in the SOURCE library are processed<br>
Some data transormation occurs at this stage, minly the contract coverage determination plus<br>
some additional data clensing in order to make data more consistent over time since reports<br>
arriving sometimes differ in presentation<br><br>
Finally data is placed in the sharepoint list which acts as a datasource for powerapps that presents<br>
data to users

<picture>
  <img alt="Sharepoint source library" src="images/sp_source_portal1.PNG">
</picture>
<br>
<picture>
  <img alt="Sharepoint source library" src="images/sp_source_portal2.PNG">
</picture>

![image](https://user-images.githubusercontent.com/17108964/175505226-45133ab4-4a98-4c6c-9d82-1b2650db748d.png)
