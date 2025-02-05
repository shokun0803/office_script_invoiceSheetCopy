# office_script_invoiceSheetCopy
This is an Office Script that duplicates the original invoice tab created in Excel and creates a tab for the next month. It is intended to retrieve and reflect numerical values ​​when executed in Power Automate.

## prerequisite
* You will need to create the original invoice in Excel beforehand.
* I need to send the working hours from Power Automate to a numeric variable "workMinutes".
* The numeric variable "hourlyRate" is hard-coded to an appropriate value. It needs to be changed to an appropriate value.

## Short Description
This is a simple script that simply assigns the billing date of a previously created invoice and a formula that calculates the working hours from the hourly rate into specific cells.
Since there are only two places to be rewritten, you need to fill in the other necessary information on the original invoice. Copy the original sheet and write the calculated results in the cells of the necessary items.
Since the setting is such that working hours are obtained using Power Automate, once you have set up an office script in the Excel file to which this script is applied, you will need to send a variable when calling this script on the Power Automate side.
If there are other items you want to include in the invoice, please add them appropriately. Note that the hourly rate is hard-coded in the script, so if you want to share a Power Automate flow, you will need to send the hourly rate from Power Automate, or some other method.
