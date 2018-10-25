# Save-excel-sheets-individually & backup cell value
## Update notes
### Current Version: 4.0
### Pass versions:
  - Ver. 1.0 User interface 
  - Ver. 1.1 Call Excel application
  - Ver. 1.3 Create new workbooks with .xlsx worksheets name
  - Ver. 1.5 Create new individual workbooks only with specific .xlsx worksheets name
  - Ver. 2.0 Added supprot to .xls
  - Ver. 2.1 Added new functions (Choose single file or multiple files)
  - Ver. 2.2 Added function to verify Excel application
  - Ver. 2.4 Enhanced exported Excel file sizes
  - Ver. 2.8 Added warning messages and user friendly functions
  - Ver. 3.0 Finalized Excel exporting function and fixed bugs
  - Ver. 3.5 Added ini export function
  - Ver. 4.0 Finalized ini export function and fixed bugs

### *First try of programming in VB.net

## Main function:
- Save every worksheets inside .xls or .xlsx file to new individual workbooks
- Read the lines inside worksheets and extract specific data
- Save the worksheets name and data inside worksheets to a ini file for future usage

## System requirements:
- For windows x86 x64
- Requires .Net framework 4.5
- Excel 2010 or above

### Tested on
- Windows 7 32bit with Excel 2010

## To be added in future:
  - Ini files storing sheet name and company name
  - Add reference of program script used from internet
  
## To-Do (Past and done):

  - Reduce size of new individual workbooks
    - Fixed: Use copy function and build new workbook
  - .xls file cannot save as .xlsx
    - Fixed: Specify second parameter (File Format) instead of directly changing file extension
           Use copy function to copy sheet content to newly created workbook
  - Simplifying the script
  - Adjust the variables (private? public?)
    - Fixed
  - Check exist files and overwrite or not
    - Fixed
  - Signals
    - XXX number of file(s) processed
    - Fixed

## Remainder & Knowledge：
- Workbook should be closed before assiging it to "Nothing" (workbook.close() workbook=Nothing)
- Garbage collect is related to the declaration of variables and COMs
- gc.collect() should be used carefully
- Buttons need to be manully disabled while running
- ToolStripProgressBar = percentage
- Note: I've separated it out into two lines because it's important not use double-dot references with Office interop (e.g., Worksheet.Cell.Value) because you end up with objects you can't release, which will cause issues with Excel not closing properly.
https://stackoverflow.com/questions/23004274/vb-net-excel-worksheet-cells-value

## Reference:
[To be added]
https://stackoverflow.com/questions/20469524/save-a-file-in-xlsx-format-in-vb
https://docs.microsoft.com/zh-hk/dotnet/api/microsoft.office.interop.excel?view=excel-pia
