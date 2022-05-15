# 1/ What is Excelam?

Excelam is a open-source .NET library over OpenXml to use Excel easily.
The library is writen in C#/net6.

The goal is to get cell value format, get cell value and set cell value, in different formats: general, number, decimal, date, currency,...

The only dependency is DocumentFormat.OpenXml (Open XML SDK), the official Microsoft library to work with Word, Excel and PowerPoint documents.
use the last version: 2.16.0.

To have code samples, see the tests project: Excelam.Tests.

A nuget package has been published:
https://www.nuget.org/packages/Excelam/0.0.1

Next stages will be to manage others types: date, time, fraction, percentage, scientific and currencies (Euro, Dollar and others).

# 2. Create or open an Excel file

## 2.1. Create a new excel file

Create an Excel file, provide the filename and the name of the first sheet.

```csharp
ExcelApi excelApi = new ExcelApi();
string fileName = @"Files\NewExcel.xlsx";

ExcelWorkbook excelWorkbook;
ExcelError error;
bool res = excelApi.ExcelFileApi.CreateExcelFile(fileName, excelApi.ExcelFileApi.DefaultFirstSheetName, out excelWorkbook, out error);

// get the first sheet
var sheet = excelApi.ExcelSheetApi.GetSheet(excelWorkbook, 0);

// do something in the excel file...

// save and close the file
res= excelApi.ExcelFileApi.CloseExcelFile(excelWorkbook, out error);
```

## 2.2. Open an existing excel file

```csharp
ExcelApi excelApi = new ExcelApi();
string fileName = @"Files\MyExcel.xlsx";

ExcelWorkbook excelWorkbook;
ExcelError error;
bool res = excelApi.ExcelFileApi.OpenExcelFile(fileName, out excelWorkbook, out error);

// get the first sheet
var sheet = excelApi.ExcelSheetApi.GetSheet(excelWorkbook, 0);

// do something in the excel file...

// save and close the file
res= excelApi.ExcelFileApi.CloseExcelFile(excelWorkbook, out error);
```

## 2.3. get a sheet par index or by name

```csharp

// 1. get the first sheet by index
var sheet = excelApi.ExcelSheetApi.GetSheet(excelWorkbook, 0);

// 2. get the sheet by name
var sheet2 = excelApi.ExcelSheetApi.GetSheetByName(excelWorkbook, "MySheet");
```

# 3. Get cell value format

The GetCellFormat() function read the cell value format. It return an object ExcelCellFormat containing an enum value for the format: ExcelCellFormatCode.
If the format is not reconized (not implemented), the enum value is: Undefined.

## 3.1. Get the cell value format

```csharp
// get the A1 cell value format
ExcelCellFormat cellFormatA1= excelApi.ExcelCellValueApi.GetCellFormat(excelSheet, "A1");
// the result: cellFormatA1.Code=ExcelCellFormatCode.General
```

## 3.2. List of managed cell value format 

Get the cell value format, can be: </br>
General (string), Number (integer), Decimal (double), DateShort (DateTime), Currency (double),...

For now, only General, Number and Decimal format are managed by the library.

format | 
:-- | 
foo | 
bar | 
baz | 

## 3.3. Is the cell a formula?

The GetCellFormat() function return also is the cell contains a formula.

```csharp
// set an int in a cell
ExcelCellFormat cellFormatE7 = excelApi.ExcelCellValueApi.GetCellFormat(sheet, "E7");
if(cellFormatE7.IsFormula)
    // SUM(E5:E6)
    Console.WriteLine("E7 Formula: " + excelApi.ExcelCellValueApi.GetCellFormula(sheet, "E7"));
```

# 4. Get cell value

## 4.1. Get a cell value as a string

Get a cell value as a string even is the type is different.

```csharp
string cellValB1= excelApi.ExcelCellValueApi.GetCellValueAsString(sheet, "B1");
```

## 4.2. Get a cell as an int and double

```csharp
// get the cell value as an integer (number in Excel)
int cellValB23;
bool res = excelApi.ExcelCellValueApi.GetCellValueAsNumber(sheet, "B23", out cellValB23);

// get the cell value as a double (decimal in Excel)
double cellValB25;
res = excelApi.ExcelCellValueApi.GetCellValueAsDecimal(sheet, "B25", out cellValB25);
```

## 4.2. Get a cell value currency 

To get a cell value which is currency, get it as decimal (double).

```csharp
//--$91,25 - currency-dollarUS - get the cell value currency as a double
double cellVal34;
res = excelApi.ExcelCellValueApi.GetCellValueAsDecimal(sheet, "B34", out cellValB34);
```


# 5. Set a cell value

## 5.1. Set a cell string value

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

// save and close the excel file
excelApi.ExcelFileApi.CloseExcelFile(workbook, out error);
```


## 5.2. Set an int or a double value in a cell 

```csharp
// set an int in a cell
bool res = excelApi.ExcelCellValueApi.SetCellValueNumber(sheet, "B23", 890);

// set a double in a cell
res = excelApi.ExcelCellValueApi.SetCellValueDecimal(sheet, "B25", 12.34);
```


# 6. Convert and split an Excel cell address 

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

