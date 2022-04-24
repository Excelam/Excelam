# What is Excelam?

Excelam is a .NET library over OpenXml to use Excel easily.
The library is writen in C# .NET6.

To have code samples, see the tests project: Excelam.Tests.

A nuget package has been published:
https://www.nuget.org/packages/Excelam/0.0.1

# Start using the library

```
// create the api to work with an Excel file
ExcelApi excelApi = new ExcelApi();

// open an existing file
ExcelWorkbook workbook;
ExcelError error;
bool res = excelApi.ExcelFileApi.OpenExcelFile(fileName, out workbook, out error);

// get the first sheet
ExcelSheet? excelSheet = excelApi.ExcelSheetApi.GetSheet(workbook, 0);

// set 'hello' - general in A1 cell
excelApi.ExcelCellValueApi.SetCellValueGeneral(excelSheet, "A1", "hello");

// get the A1 cell value format
ExcelCellFormat cellFormatA1= excelApi.ExcelCellValueApi.GetCellFormat(excelSheet, "A1");
// the result: cellFormatA1.Code=ExcelCellFormatCode.General

// save and close the excel file
excelApi.ExcelFileApi.CloseExcelFile(workbook, out error);

```
