# XLSX Advanced Options
  
Module with advanced options for XLSX  

*Read this in other languages: [English](Manual_AdvancedXLSX.md), [Español](Manual_AdvancedXLSX.es.md).*
  
![banner](imgs/Banner_advancedxlsx.png)
## How to install this module
  
__Download__ and __install__ the content in 'modules' folder in Rocketbot path  



## Description of the commands

### Open xls
  
Open a xls file to work with native command
|Parameters|Description|example|
| --- | --- | --- |
|Path to XLS file |Select the XLS file you want to open|file.XLS|
|Id (optional) |Session identifier|id|

### Get Fromatted cells
  
Get cells with format
|Parameters|Description|example|
| --- | --- | --- |
|Cells |Cells range|A1:B5|
|Assign result to variable |Variable name where result will be saved|Variable|

### Create sheet
  
Create a new sheet
|Parameters|Description|example|
| --- | --- | --- |
|Sheet name |Sheet name that will be created|Sheet 2|

### Count in range
  
Returns the maximum number of rows and columns from a cell
|Parameters|Description|example|
| --- | --- | --- |
|Sheet name |Sheet name where the range is located|Sheet 2|
|Start cell|Start cell of the range|A1|
|Assign row length to variable |Variable name where the row length will be saved|Variable|
|Assign column length to variable |Variable name where the column length will be saved|Variable|

### Column filter
  
Filter by column
|Parameters|Description|example|
| --- | --- | --- |
|Filters |Filters to apply.|["A > 3", "D *ARS", "C == Invoice"]|
|Sheet's name |Sheet's name to filter.|sheet1|
|Detailed result|Mark to get detailed result.|True|
|Variable where to store the result |Variable where the result will be stored.|result|

### Delete Row/Column
  
Command to delete rows or columns
|Parameters|Description|example|
| --- | --- | --- |
|Sheet name |Name of the sheet where the row or column will be deleted|Sheet 2|
|Row(s)|Range of rows to delete|1:5|
|Column(s)|Range of columns to delete|A:G|
