# XLSX Advanced Options
  
Módulo con opciones avanzadas para XLSX  

*Read this in other languages: [English](Manual_AdvancedXLSX.md), [Español](Manual_AdvancedXLSX_es.md), [Portugues](Manual_AdvancedXLSX_pr.md).* 

![banner](imgs/Banner_AdvancedXLSX.png)

## How to install this module
  
__Download__ and __install__ the content in 'modules' folder in Rocketbot path  



## Description of the commands

### Open xls
  
Open a xls file to work with native command
|Parameters|Description|example|
| --- | --- | --- |
|Path to XLS file |Select the XLS file you want to open|example.xls|
|Column/s as date (optional) ||0|
|Id (optional) |Session identifier|id|
|Assign result to variable||Variable|

### Fromat cells
  
Give format to cells
|Parameters|Description|example|
| --- | --- | --- |
|Sheet ||Sheet1|
|Cells |Cells range|A1:B5|
|ID Formato |Format ID. Check Documentation https//learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.spreadsheet.numberingformat?view=openxml-2.8.1|0|
|Assign result to variable||Variable|

### Create sheet
  
Create a new sheet
|Parameters|Description|example|
| --- | --- | --- |
|Sheet name |Sheet name that will be created|Sheet2|

### Remove sheet
  
Remove a sheet from workbook
|Parameters|Description|example|
| --- | --- | --- |
|Sheet name ||Sheet1|

### Count in range
  
Returns the maximum number of rows and columns from a cell
|Parameters|Description|example|
| --- | --- | --- |
|Sheet name |Sheet name where the range is located|Sheet1|
|Start cell|Start cell of the range|A1|
|Assign result to variable (Row)|Variable name where the row length will be saved|Variable|
|Assign result to variable (Column)|Variable name where the column length will be saved|Variable|

### Column filter
  
Filter by column
|Parameters|Description|example|
| --- | --- | --- |
|Filters |Filters to apply.|["A > 3", "D *ARS", "C == Invoice"]|
|Sheet's name |Sheet's name to filter.|Sheet1|
|Detailed result|Mark to get detailed result.|True|
|Assign result to variable||Variable|

### Delete Row/Column
  
Command to delete rows or columns
|Parameters|Description|example|
| --- | --- | --- |
|Sheet name |Name of the sheet where the row or column will be deleted|Sheet1|
|Row(s)|Range of rows to delete|1:5|
|Column(s)|Range of columns to delete|A:G|

### Insert Row/Column
  
Command to insert rows or columns
|Parameters|Description|example|
| --- | --- | --- |
|Sheet name |Name of the sheet where the row or column will be deleted|Sheet1|
|Row(s)|Range of rows to delete|1:5|
|Column(s)|Range of columns to delete|A:G|
