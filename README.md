# 1. What is Excelam?

Excelam is an open-source .NET library over OpenXml to use Excel easily.
The library is writen in C#/net6.

The goal of the libray is focused on working on cell value in different formats available in Excel:

- <b>Get Cell Format</b>

- <b>Get Cell Value</b>

- <b>Set Cell Value</b>

The only dependency is DocumentFormat.OpenXml (Open XML SDK), the official Microsoft library to work with Word, Excel and PowerPoint documents.
The library use the last version: 2.16.0.

To have code samples, see the tests project: Excelam.Tests.

A nuget package has been published:
https://www.nuget.org/packages/Excelam

Next stages will be to manage more types and more cases: Accounting, Fraction, Percentage, and Scientific.

# 2. Excel Cell value format

Excel cell modeling is rich but also complex.

## 2.1. Cell modeling

Cell modeling is composed of four main parts:

- <b>Value format, Fill, Border and Font.</b>
    
And also these two specific parts:

- Alignment, Protection.

The library concerns currently the value format of cell. Others parts are not managed.

## 2.2. Cell value format

Excel provide a rich list of categories for cell value format:

- <b>General, Text, Number, Decimal, Date and Time, Currency, Accounting, Percentage, Fraction and scientific.</b>

General, Text and Number categories are very basic. For others categories, it's possible to configure the format with more details.

All cell value format are identified by an integer named NumberFormatId.
Some cell value format are built-in, NumberFormatId is set to a value lesser than 164.
For format that are not built-in, a string containing the format must be set and the NumberFormatId is greater than 163.

To have more details on the excel cell value format categries, go [here](https://github.com/Excelam/Excelam/wiki/Excel-cell-value-format).

Be carreful, cell value format concerns only the display of the value but the inner value in the cell can be formatted in another format. 


# 3. Create or open an Excel file

## 3.1. Create a new excel file

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

## 3.2. Open an existing excel file

How to open an existing Excel file:

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

## 3.3. Save and close an excel file

It's possible to save the modification of the current excel file.

On closing, the modifications are automatically saved.


```csharp
ExcelApi excelApi = new ExcelApi();
string fileName = @"Files\MyExcel.xlsx";

ExcelWorkbook excelWorkbook;
ExcelError error;
bool res = excelApi.ExcelFileApi.OpenExcelFile(fileName, out excelWorkbook, out error);

// get the first sheet
var sheet = excelApi.ExcelSheetApi.GetSheet(excelWorkbook, 0);

// do something in the excel file...

// save the modifications
excelApi.ExcelFileApi.SaveExcelFile(excelWorkbook, out error);

// do again modifications...

// save and close the file
res= excelApi.ExcelFileApi.CloseExcelFile(excelWorkbook, out error);
```

## 3.4. Get a sheet by index or by name

It's possible to get a sheet from Excel by index, starting from 0, or by the name.

```csharp

// 1. get the first sheet by index
var sheet = excelApi.ExcelSheetApi.GetSheet(excelWorkbook, 0);

// 2. get the sheet by name
var sheet2 = excelApi.ExcelSheetApi.GetSheetByName(excelWorkbook, "MySheet");
```


# 3. Get cell value format : current cases

TODO: rework it!!

The GetCellFormat() function read the cell value format. It return an object ExcelCellFormat containing an enum value for the format: ExcelCellFormatCode.
If the format is not reconized (not implemented), the enum value is: Undefined.

## 3.1. Get the cell value format: general case

```csharp
// get the A1 cell value format
ExcelCellFormat cellFormatA1= excelApi.ExcelCellValueApi.GetCellFormat(excelSheet, "A1");
// the result: cellFormatA1.Code=ExcelCellFormatCode.General
```

## 3.2. List of managed cell value format 

Get the cell value format.

For now, only these cell value format are managed by the library.

Excel     | C# 
:--       | :--: |
General   | string
Number    | int 
Decimal   | double
DateShort | DateTime
DateLarge | DateTime
Time      | DateTime
Currency  | double

# 4. Get cell value format : currency case

## 4.1. Is the cell a currency?

Is a cell value format is currency, then ExcelCellFormat.Code (type is ExcelCellFormatCode) will be set to currency.
The other property ExcelCellFormat.CurrencyCode will be set by the exact currency code.

```csharp
//--B23: $91,25 - currency-dollarUS
ExcelCellFormat cellFormatB23 = excelApi.ExcelCellValueApi.GetCellFormat(sheet, "B23");

// will display:  Currency
Console.WriteLine("cellFormatB23.Code: " + excelApicellFormatB23.Code);

// will display:  UnitedStatesDollar
Console.WriteLine("cellFormatB23.CurrencyCode: " + excelApicellFormatB23.CurrencyCode);
```

## 4.2. Is the cell a accounting?

Excel have a special cell format: accounting which is a in fact a formatted currency value.

```csharp
//--B21: 45,21 €  - accounting
ExcelCellFormat cellFormatB21 = excelApi.ExcelCellValueApi.GetCellFormat(sheet, "B21");

// will display:  Accounting
Console.WriteLine("cellFormatB21.Code: " + excelApicellFormatB21.Code);

// will display:  Euro
Console.WriteLine("cellFormatB21.CurrencyCode: " + excelApicellFormatB21.CurrencyCode);
```


## 4.3. List of managed cell value currency format  

For now, only these cell value formats are managed by the library.

Code      | Currency     | country 
:--       | :--: | :--: 
Euro      | Euro | -
UnitedStatesDollar    | $-Dollar |  USA



# 5. Get cell value format : formula case

## 5.1. Is the cell a formula?

The GetCellFormat() function return also is the cell contains a formula.

```csharp
// set an int in a cell
ExcelCellFormat cellFormatE7 = excelApi.ExcelCellValueApi.GetCellFormat(sheet, "E7");
if(cellFormatE7.IsFormula)
    // SUM(E5:E6)
    Console.WriteLine("E7 Formula: " + excelApi.ExcelCellValueApi.GetCellFormula(sheet, "E7"));
```

# 6. Get cell value

## 6.1. Get a cell value as a string

Get a cell value as a string even is the type is different.

```csharp
string cellValB1= excelApi.ExcelCellValueApi.GetCellValueAsString(sheet, "B1");
```

## 6.2. Get a cell as an int and double

```csharp
// get the cell value as an integer (number in Excel)
int cellValB23;
bool res = excelApi.ExcelCellValueApi.GetCellValueAsNumber(sheet, "B23", out cellValB23);

// get the cell value as a double (decimal in Excel)
double cellValB25;
res = excelApi.ExcelCellValueApi.GetCellValueAsDecimal(sheet, "B25", out cellValB25);
```

## 6.3. Get a cell value dateShort, dateLarge and time

Is a cell value format is DateShort, DateLarge or time, you have to use GetCellValueAsDateTime() to read the value.
For these excel types, the value will be a DateTime.

```csharp
DateTime valB17;        
res = excelApi.ExcelCellValueApi.GetCellValueAsDateTime(sheet, "B17", out valB17);
```


## 6.3. Get a cell value currency 

To get a cell value which is currency, get it as decimal (double).

```csharp
//--$91,25 - currency-dollarUS - get the cell value currency as a double
double cellVal34;
res = excelApi.ExcelCellValueApi.GetCellValueAsDecimal(sheet, "B34", out cellValB34);
```


# 7. Set a cell value

## 7.1. Set a cell string value

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


## 7.2. Set an int or a double value in a cell 

```csharp
// set an int in a cell
bool res = excelApi.ExcelCellValueApi.SetCellValueNumber(sheet, "B23", 890);

// set a double in a cell
res = excelApi.ExcelCellValueApi.SetCellValueDecimal(sheet, "B25", 12.34);
```


# 8. Convert and split an Excel cell address 

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

