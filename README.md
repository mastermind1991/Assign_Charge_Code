assign_charge_code
This script doese the following:

Queries ServiceNow for "Assign Charge Code to Item in Cerner PROD" in the SCTASK Short description
Grabs the bill items from the ticket description
Compiles them into a spreadsheet
Adds the spreadsheet to the ticket

Note: This can be ran on a schedule

Spreadsheet Info
The spreadsheet created is store in the root folder labeled "Spreadsheet" for later reference


Naming Convention: charge_services_data_<SCTASK#>.xlsx

Note: The spreadsheet can be deleted out of this directory if you would like. It is only stored for historical reference

Log Info


Location: root\logs

Convention:

Created Spreadsheet for: <SCTASK#>
Spreadsheet already exists on: <SCTASK#>
No bill items found on: <SCTASK#>
No tickets are open or on hold with that description
