



# XLSX Advanced Options
  
Format cells, create and remove sheets, filter data, add and delete columns and rows, open xls files and transform them into xlsx format.  

*Read this in other languages: [English](Manual_AdvancedXLSX.md), [Português](Manual_AdvancedXLSX.pr.md), [Español](Manual_AdvancedXLSX.es.md)*
  
![banner](imgs/Banner_AdvancedXLSX.png)
## How to install this module
  
To install the module in Rocketbot Studio, it can be done in two ways:
1. Manual: __Download__ the .zip file and unzip it in the modules folder. The folder name must be the same as the module and inside it must have the following files and folders: \__init__.py, package.json, docs, example and libs. If you have the application open, refresh your browser to be able to use the new module.
2. Automatic: When entering Rocketbot Studio on the right margin you will find the **Addons** section, select **Install Mods**, search for the desired module and press install.  

## How to use this module

Only if you are using the 2023 version of Rocketbot should you follow the following steps to avoid the error:

ImportError: cannot import name 'etree' from 'lxml'

1. You should go to the root folder of Rocketbot and verify that the 'lxml' library exists.
2. If it does not exist, from a terminal, go to the root folder of Rocketbot and type:  pip install lxml -t .
3. Please note that you should install the library with Python 3.10 64-bit.


## Description of the commands

### Open xls
  
Open a xls file to work with native command
|Parameters|Description|example|
| --- | --- | --- |
|Path to XLS file |Select the XLS file you want to open|example.xls|
|Column/s as date (optional) ||0|
|Id (optional) |Session identifier|id|
|Encoding|Type of Encoding to apply. Default Latin-1|latin-1|
|Assign result to variable||Variable|

### Open advanced xlsx
  
Open a xlsx file to work with native command
|Parameters|Description|example|
| --- | --- | --- |
|Path to XLSX file |Select the XLSX file you want to open|example.xlsx|
|Read only|Check if you want to open the xlsx in read only mode, the content cannot be edited.|False|
|Keep vba|Check to keep the possible VBA code that could be in the workbook.|False|
|Data only|Controls if cells with formulas have the formula (default) or the value stored the|False|
|Keep links|Check if links to external workbooks should be kept.|False|
|Id (optional) |Session identifier|id|
|Assign result to variable||Variable|

### Convert xls to xlsx
  
Convert an xls format file to xlsx format
|Parameters|Description|example|
| --- | --- | --- |
|Path to XLS file |Select the XLS file you want to open|path/to/file/example.xls|
|Path to XLSX file |Put the full path where you want to save the XLSX file (including name and '.xlsx' extension)|path/to/file/example.xlsx|
|Encoding|Type of Encoding to apply. Default Latin-1|latin-1|

### Convert sheet to csv
  
Convert a sheet of the opened xlsx file to csv
|Parameters|Description|example|
| --- | --- | --- |
|Path to CSV file |Select the CSV file you want to open|path/to/file/example.csv|
|Delimiter|Delimiter of the csv file|,|
|Date output format|Format with which the dates of the xlsx Sheet will be converted to csv|%d/%m/%Y|
|Assign result to variable |Name of the variable where the result will be stored|Variable|

### Read range
  
Returns the value of the given range. One value if the range is a cell or a list if the range has multiple cells.
|Parameters|Description|example|
| --- | --- | --- |
|Sheet name |Sheet name where the range is located|Sheet1|
|Cell or range|Start cell of the range|A1|
|Assign result to variable (Column)|Variable name where the column length will be saved|Variable|

### Rename sheet
  
Rename a sheet
|Parameters|Description|example|
| --- | --- | --- |
|Sheet name to rename |Name of the sheet to rename|OldSheet|
|New sheet name|Name of the sheet|NewSheet|

### Format cells
  
Give format to cells
|Parameters|Description|example|
| --- | --- | --- |
|Sheet ||Sheet1|
|Cells |Cells range|A1:B5|
|Horizontal Alignment||---- Select ----|
|Vertical Alignment||---- Select ----|
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
|Filters |Filters to apply. For empty filters use == None|["A > 3", "D *ARS", "C == Invoice"]|
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

### Insert image
  
Insert an image into a document
|Parameters|Description|example|
| --- | --- | --- |
|Image path |Select the image file you want to insert into the document|example.png|
|Sheet |Name of the document sheet where to insert the image|Sheet1|
|Cell |Cell where to insert the image|A1|

### Close xlsx
  
Close an open xlsx file
|Parameters|Description|example|
| --- | --- | --- |
|Assign result to variable||Variable|
