# What is Excelam?

Excelam is a .NET library over OpenXml to use Excel easily.
The library is writen in C# dotnet6.

The only dependency is DocumentFormat.OpenXml (Open XML SDK), the official Microsoft library to work with Word, Excel and PowerPoint documents.
use the last version: 2.16.0.

To have code samples, see the tests project: Excelam.Tests.

A nuget package has been published:
https://www.nuget.org/packages/Excelam/0.0.1

Next stages will be to manage others types: date, time, fraction, percentage, scientific and currencies (Euro, Dollar and others).

# Start using the library

## create a new excel file

Create an Excel file, provide the filename and the name of the first sheet.

```csharp
ExcelApi excelApi = new ExcelApi();
string fileName = @"Files\NewExcel.xlsx";

ExcelWorkbook excelWorkbook;
ExcelError error;
bool res = excelApi.ExcelFileApi.CreateExcelFile(fileName, excelApi.ExcelFileApi.DefaultFirstSheetName, out excelWorkbook, out error);

// save and close the file
res= excelApi.ExcelFileApi.CloseExcelFile(excelWorkbook, out error);
```

## Open an existing excel file

```csharp
ExcelApi excelApi = new ExcelApi();
string fileName = @"Files\MyExcel.xlsx";

ExcelWorkbook excelWorkbook;
ExcelError error;
bool res = excelApi.ExcelFileApi.OpenExcelFile(fileName, out excelWorkbook, out error);

// save and close the file
res= excelApi.ExcelFileApi.CloseExcelFile(excelWorkbook, out error);
```

## Set a cell string value

Set a string value in a cell, the corresponding Excel type is General. 

```csharp
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

```csharp
string cellValB1= excelApi.ExcelCellValueApi.GetCellValueAsString(sheet, "B1");
```

## Get a cell as an int and double

```csharp
// get the cell value as an integer (number in Excel)
int cellValB23;
bool res = excelApi.ExcelCellValueApi.GetCellValueAsNumber(sheet, "B23", out cellValB23);

// get the cell value as a double (decimal in Excel)
double cellValB25;
res = excelApi.ExcelCellValueApi.GetCellValueAsDecimal(sheet, "B25", out cellValB25);
```

## Set an int or a double value in a cell 

```csharp
// set an int in a cell
bool res = excelApi.ExcelCellValueApi.SetCellValueNumber(sheet, "B23", 890);

// set a double in a cell
res = excelApi.ExcelCellValueApi.SetCellValueDecimal(sheet, "B25", 12.34);
```

## Convert and split an Excel cell address 

```csharp
//--convert a col and a row int values into an excel address
int col=2;
int row=12;
string cellAddress = ExcelCellAddressApi.ConvertAddress(col, row);
// result: cellAddress: B12

//--decode, split an excel cell address
string colRowName="B12";
string colName;
bool res= ExcelCellAddressApi.SplitCellAddress(colRowName, out colName, out col, out row);
// result: colName: "B", col:2, row:12
```

