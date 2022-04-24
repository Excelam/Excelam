# What is Excelam?

Excelam is a .NET library over OpenXml to use Excel easily.
The library is writen in C# .NET6.

To have code samples, see the tests project: Excelam.Tests.

A nuget package has been published:
https://www.nuget.org/packages/Excelam/0.0.1

# Start using the library

ExcelApi excelApi = new ExcelApi();

ExcelWorkbook workbook;
ExcelError error;
bool res = excelApi.ExcelFileApi.OpenExcelFile(fileName, out workbook, out error);
