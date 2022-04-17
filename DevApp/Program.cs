﻿// See https://aka.ms/new-console-template for more information
using Excelam;
using Excelam.System;

void DisplayRes(string msg, bool res)
{
    if (!res)
    {
        Console.WriteLine("==> " + msg + " Error occurs!");
    }
    else
        Console.WriteLine("==> " + msg + " Ok.");
}

/// <summary>
/// Create a new Excel file.
/// </summary>
void CreateExcelFile()
{
    ExcelApi excelApi = new ExcelApi();

    string fileName = @"Files\NewExcel.xlsx";
    ExcelWorkbook excelDoc;
    ExcelError error;

    // if file exists, remove it
    if (File.Exists(fileName))
        File.Delete(fileName);

    bool res = excelApi.ExcelFileApi.CreateExcelFile(fileName, excelApi.ExcelFileApi.DefaultFirstSheetName, out excelDoc, out error);
    DisplayRes("Create NewFile:", res);

    res= excelApi.ExcelFileApi.CloseExcelFile(excelDoc, out error);
    DisplayRes("Close File:", res);
}

/// <summary>
/// Open an empty existing excel file.
/// </summary>
void OpenEmptyExcelFile()
{
    ExcelApi excelApi = new ExcelApi();

    string fileName = @"Files\ExcelFromTempl.xlsx";
    ExcelWorkbook excelDoc;
    ExcelError error;

    // if file exists, remove it
    if (File.Exists(fileName))
        File.Delete(fileName);

    string fileNameTempl = @"Files\ExcelEmpty.xlsx";
    File.Copy(fileNameTempl, fileName);

    bool res = excelApi.ExcelFileApi.OpenExcelFile(fileName, out excelDoc, out error);
    DisplayRes("Open File:", res);

    res = excelApi.ExcelFileApi.CloseExcelFile(excelDoc, out error);
    DisplayRes("Close File:", res);
}

/// <summary>
/// Open an existing excel file containing basic built-in cell format.
/// </summary>
void OpenExcelFileBasicBuiltInCellFormat()
{
    ExcelApi excelApi = new ExcelApi();

    string fileName = @"Files\FromManyCellTypes.xlsx";
    ExcelWorkbook workbook;
    ExcelError error;

    // if file exists, remove it
    if (File.Exists(fileName))
        File.Delete(fileName);

    string fileNameTempl = @"Files\ManyCellTypes.xlsx";
    File.Copy(fileNameTempl, fileName);

    bool res = excelApi.ExcelFileApi.OpenExcelFile(fileName, out workbook, out error);
    DisplayRes("Open File:", res);

    // get the first sheet
    var sheet = excelApi.ExcelSheetApi.GetSheet(workbook, 0);

    // XXX: définir méthode get cell et format


    //--B1: null
    ExcelCellFormat cellFormatB1 = excelApi.ExcelCellValueApi.GetCellFormat(sheet, "B1");

    //--B2: bonjour - standard/general
    ExcelCellFormat cellFormatB2 = excelApi.ExcelCellValueApi.GetCellFormat(sheet, "B2");
    string valB2 = excelApi.ExcelCellValueApi.GetCellValueAsString(sheet, "B2");

    //XXXXXXXXXXXXXXXXXXX-reflexion, GetCell:

    //--recup cell avec seulement son format value (et pas le reste, pour aller plus vite)
    //ExcelCell cellB4 = excelApi.ExcelCellContentApi.GetCellOnlyFormatValue(sheet, "B4");


    // get the value of the cell:
    // string valB4=excelApi.ExcelCellContentApi.GetCellValueAsString(sheet, cellB4);

    // get the value of the cell:
    // passer un param pour indiquer quoi récuperer sur la cell
    // enum CellFormatSelection: Value, Border, Fill, Font, Alignement, Protection
    // ExcelCell cell= GetCell(sheet, "B1", CellFormatSelection)

    //--recup cell avec tout: format value, border, fill, font, Alignement and Protection
    // ExcelCell cell= GetCell(sheet, "B1")

    //XXXXXXXXXXXXXXXXXXX:




    // write cell value in format
    //excelApi.ExcelCellContentApi.SetCellDateLarge(excelSheet, "D4", DateTime value);
    // TODO: penser a plus+tard, set autre format: border, fill, font...
    // passer par ExcelCell

    excelApi.ExcelCellValueApi.SetCellValueNumber(sheet, "G2", 12);
   
    //--close the file
    res = excelApi.ExcelFileApi.CloseExcelFile(workbook, out error);
    DisplayRes("Close File:", res);
}

/// <summary>
/// delete cells.
/// </summary>
void DeleteCell()
{

    ExcelApi excelApi = new ExcelApi();

    string fileName = @"Files\FromManyCellTypes.xlsx";
    ExcelWorkbook workbook;
    ExcelError error;

    // if file exists, remove it
    if (File.Exists(fileName))
        File.Delete(fileName);

    string fileNameTempl = @"Files\ManyCellTypes.xlsx";
    File.Copy(fileNameTempl, fileName);

    bool res = excelApi.ExcelFileApi.OpenExcelFile(fileName, out workbook, out error);
    DisplayRes("Open File:", res);

    // get the first sheet
    var sheet = excelApi.ExcelSheetApi.GetSheet(workbook, 0);

    excelApi.ExcelCellValueApi.DeleteCell(sheet, "B1");
    excelApi.ExcelCellValueApi.DeleteCell(sheet, "B2");
    excelApi.ExcelCellValueApi.DeleteCell(sheet, "B3");
    excelApi.ExcelCellValueApi.DeleteCell(sheet, "B4");

    //--close the file
    res = excelApi.ExcelFileApi.CloseExcelFile(workbook, out error);
    DisplayRes("Close File:", res);

}

/// <summary>
/// set cell value.
/// </summary>
void SetCellValues()
{

    ExcelApi excelApi = new ExcelApi();

    string fileName = @"Files\SetCellValues.xlsx";
    ExcelWorkbook workbook;
    ExcelError error;

    // if file exists, remove it
    if (File.Exists(fileName))
        File.Delete(fileName);

    string fileNameTempl = @"Files\ExcelEmpty.xlsx";
    File.Copy(fileNameTempl, fileName);

    bool res = excelApi.ExcelFileApi.OpenExcelFile(fileName, out workbook, out error);
    DisplayRes("Open File:", res);

    // get the first sheet
    var sheet = excelApi.ExcelSheetApi.GetSheet(workbook, 0);

    excelApi.ExcelCellValueApi.SetCellValueNumber(sheet, "B1",12);

    //--close the file
    res = excelApi.ExcelFileApi.CloseExcelFile(workbook, out error);
    DisplayRes("Close File:", res);

}

/// <summary>
/// open existing excel file to check the styles.
/// </summary>
void ReadExcelFile()
{

    ExcelApi excelApi = new ExcelApi();

    string fileName = @"Files\OneCellNumberFillBorder.xlsx";
    ExcelWorkbook workbook;
    ExcelError error;

    bool res = excelApi.ExcelFileApi.OpenExcelFile(fileName, out workbook, out error);
    DisplayRes("Open File:", res);

    // get the first sheet
    var sheet = excelApi.ExcelSheetApi.GetSheet(workbook, 0);

    //--C4: bonjour - standard/general
    ExcelCellFormat cellFormatC4 = excelApi.ExcelCellValueApi.GetCellFormat(sheet, "C4");
    string valC4 = excelApi.ExcelCellValueApi.GetCellValueAsString(sheet, "C4");

    //--close the file
    res = excelApi.ExcelFileApi.CloseExcelFile(workbook, out error);
    DisplayRes("Close File:", res);

}

void ReadExcelFileSetManyCellType()
{

    ExcelApi excelApi = new ExcelApi();

    string fileName = @"Files\SetManyCellType.xlsx";
    ExcelWorkbook workbook;
    ExcelError error;

    bool res = excelApi.ExcelFileApi.OpenExcelFile(fileName, out workbook, out error);
    DisplayRes("Open File:", res);

    // get the first sheet
    var sheet = excelApi.ExcelSheetApi.GetSheet(workbook, 0);

    //--B9: blue text - standard/general, fill=blue
    ExcelCellFormat cellFormatB9= excelApi.ExcelCellValueApi.GetCellFormat(sheet, "B9");
    string valB9 = excelApi.ExcelCellValueApi.GetCellValueAsString(sheet, "B9");

    //--close the file
    res = excelApi.ExcelFileApi.CloseExcelFile(workbook, out error);
    DisplayRes("Close File:", res);

}

/// <summary>
/// Main
/// </summary>
Console.WriteLine("==>dev Excelam lib:");

//CreateExcelFile();

//OpenEmptyExcelFile();

//OpenExcelFileBasicBuiltInCellFormat();

//DeleteCell();

//SetCellValues();

//ReadExcelFile();

ReadExcelFileSetManyCellType();