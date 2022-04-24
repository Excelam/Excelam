# What is Excelam?

Excelam is a .NET library over OpenXml to use Excel easily.
The library is writen in C# dotnet6.

The only dependency is DocumentFormat.OpenXml (Open XML SDK), the official Microsoft library to work with Word, Excel and PowerPoint documents.
use the last version: 2.16.0.

To have code samples, see the tests project: Excelam.Tests.

A nuget package has been published:
https://www.nuget.org/packages/Excelam/0.0.1

# Start using the library

## Set a cell string value

Set a string value in a cell, the corresponding Excel type is General. 

```
// create the api to work with an Excel file
ExcelApi excelApi = new ExcelApi();

// open an existing file
ExcelWorkbook workbook;
ExcelError error;
bool res = excelApi.ExcelFileApi.OpenExcelFile(fileName, out workbook, out error);

// get the first sheet
ExcelSheet? excelSheet = excelApi.ExcelSheetApi.GetSheet(workbook, 0);

// set string 'hello' in A1 cell, Excel type is General
excelApi.ExcelCellValueApi.SetCellValueGeneral(excelSheet, "A1", "hello");

// get the A1 cell value format
ExcelCellFormat cellFormatA1= excelApi.ExcelCellValueApi.GetCellFormat(excelSheet, "A1");
// the result: cellFormatA1.Code=ExcelCellFormatCode.General

// save and close the excel file
excelApi.ExcelFileApi.CloseExcelFile(workbook, out error);
```
## Get a cell value as a string

Get a cell value as a string even is the type is different.

```
string cellValB1= excelApi.ExcelCellValueApi.GetCellValueAsString(sheet, "B1");
```

## Get a cell as an int and double

```
// get the cell value as an integer (number in Excel)
int cellValB23;
res = excelApi.ExcelCellValueApi.GetCellValueAsNumber(sheet, "B23", out cellValB23);

// get the cell value as a double (decimal in Excel)
double cellValB25;
res = excelApi.ExcelCellValueApi.GetCellValueAsDecimal(sheet, "B25", out cellValB25);
```

## Set an int or a double value in a cell 


```
// set an int in a cell
res = excelApi.ExcelCellValueApi.SetCellValueNumber(sheet, "B23", 890);

// set a double in a cell
res = excelApi.ExcelCellValueApi.SetCellValueDecimal(sheet, "B25", 12.34);
 ```

