# ResFeci for Excel

In it's current state, this package will allow you to quickly split a single Excel sheet into either separate tabs within the same report or into different Excel reports.  For example, if Column 'B' of your Excel report contains "Names" and "Jack" appears in 3 rows and "Jill" appears in 2, then, this will split the report into 2 separate reports (one for only Jack's rows and another for only Jill's rows); alternatively, this will create a new report with tabs entitled "Jack" and "Jill".  

By default, the script will split the first worksheet. However, the user may select a specific sheet to split through an optional parameter (see below examples).

**IMPORTANT: Your sheet must contain a header row to function properly!**


## EXAMPLE OF SPLITTING AN EXCEL SHEET BY UNIQUE VALUES IN COLUMN B INTO DIFFERENT REPORTS

from resfeci import excel_split

excel_split.split_into_separate_reports(input_report_path= 'PATH_TO_REPORT.xlsx' , column_number_to_split_by=2)


## EXAMPLE OF SPLITTING AN EXCEL SHEET BY UNIQUE VALUES IN COLUMN B INTO DIFFERENT TABS OF THE SAME REPORT

from resfeci import excel_split

excel_split.split_into_new_tabs_of_single_report(input_report_path= 'PATH_TO_REPORT.xlsx' , column_number_to_split_by=2)

## EXAMPLE OF SPLITTING AN EXCEL SHEET BY UNIQUE VALUES IN COLUMN 'C' OF A WORKSHEET CALLED 'NAMES' INTO DIFFERENT REPORTS

from resfeci import excel_split

excel_split.split_into_separate_reports(input_report_path= 'PATH_TO_REPORT.xlsx' , column_number_to_split_by=3, sheet='NAMES')



**PARAMETER DESCRIPTIONS**

input_report_path = the path to the report to be split

column_number_to_split_by = the NUMBER of the Excel column to split by (A = 1, B = 2, C =3, etc.)

sheet = OPTIONAL parameter which allows you to enter the name of the specific sheet to split.  By default, the script will split the first worksheet.

**DEPENDENCIES**
Openpyxl, Pandas

**KNOWN ISSUE** 
The current version of this script does not support transfer of formulas!!
