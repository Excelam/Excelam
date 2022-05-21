﻿using Excelam.System;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Excelam.Tests;

/// <summary>
/// Set cell value tests.
/// </summary>
[TestClass]
public class ExcelCellValueApiSetValueTests
{
    /// <summary>
    /// Start from an empty excel file.
    /// set cell values, test all formats.
    /// </summary>
    [TestMethod]
    public void SetCellValuesEmpty()
    {
        string fileName = @"Files\Cells\SetCellValuesEmpty.xlsx";

        ExcelApi excelApi = new ExcelApi();
        ExcelWorkbook workbook;
        ExcelError error;
        bool res = excelApi.ExcelFileApi.OpenExcelFile(fileName, out workbook, out error);
        Assert.IsTrue(res);

        // get the first sheet
        var sheet = excelApi.ExcelSheetApi.GetSheet(workbook, 0);
        Assert.IsNotNull(sheet);

        //--B1 set 'bonjour' - general
        res = excelApi.ExcelCellValueApi.SetCellValueGeneral(sheet, "B1", "bonjour");
        Assert.IsTrue(res);

        // check
        string cellValB1 = excelApi.ExcelCellValueApi.GetCellValueAsString(sheet, "B1");
        Assert.AreEqual("bonjour", cellValB1);

        //--B3, set 12 - number 
        res = excelApi.ExcelCellValueApi.SetCellValueNumber(sheet, "B3", 12);
        Assert.IsTrue(res);

        int cellValB3;
        res = excelApi.ExcelCellValueApi.GetCellValueAsNumber(sheet, "B3", out cellValB3);
        Assert.IsTrue(res);
        Assert.AreEqual(12, cellValB3);

        //--B5, set 34.56 - decimal
        res = excelApi.ExcelCellValueApi.SetCellValueDecimal(sheet, "B5", 34.56);
        Assert.IsTrue(res);

        double cellValB5;
        res = excelApi.ExcelCellValueApi.GetCellValueAsDecimal(sheet, "B5", out cellValB5);
        Assert.IsTrue(res);
        Assert.AreEqual(34.56, cellValB5);


        //--B7, set 19/05/2022 - DateShort
        res = excelApi.ExcelCellValueApi.SetCellValueDateShort(sheet, "B7", new DateTime(2022,05,19));
        Assert.IsTrue(res);

        DateTime cellValB7;
        res = excelApi.ExcelCellValueApi.GetCellValueAsDateTime(sheet, "B7", out cellValB7);
        Assert.IsTrue(res);
        Assert.AreEqual(2022, cellValB7.Year);
        Assert.AreEqual(5, cellValB7.Month);
        Assert.AreEqual(19, cellValB7.Day);

        //--B9, set xxx - Time
        //res = excelApi.ExcelCellValueApi.SetCellValueTime(sheet, "B7", new DateTime(2022, 05, 19));
        //Assert.IsTrue(res);

        // TODO: DateLarge, Time, currency,...
        //ici();

        //--close the file
        res = excelApi.ExcelFileApi.CloseExcelFile(workbook, out error);
        Assert.IsTrue(res);
    }

    /// <summary>
    /// reuse the excel file SetCellValuesEmpty.xlsx
    /// generated by the previous test.
    /// </summary>
    [TestMethod]
    public void SetCellValuesSameFormat()
    {
        string fileName = @"Files\Cells\SetCellValuesSameFormat.xlsx";

        ExcelApi excelApi = new ExcelApi();
        ExcelWorkbook workbook;
        ExcelError error;
        bool res = excelApi.ExcelFileApi.OpenExcelFile(fileName, out workbook, out error);
        Assert.IsTrue(res);

        // get the first sheet
        var sheet = excelApi.ExcelSheetApi.GetSheet(workbook, 0);
        Assert.IsNotNull(sheet);

        //--B1 set 'salute' - general
        res = excelApi.ExcelCellValueApi.SetCellValueGeneral(sheet, "B1", "salute");
        Assert.IsTrue(res);

        // check
        string cellValB1 = excelApi.ExcelCellValueApi.GetCellValueAsString(sheet, "B1");
        Assert.AreEqual("salute", cellValB1);

        //--B3, set 14 - number 
        res = excelApi.ExcelCellValueApi.SetCellValueNumber(sheet, "B3", 14);
        Assert.IsTrue(res);

        int cellValB3;
        res = excelApi.ExcelCellValueApi.GetCellValueAsNumber(sheet, "B3", out cellValB3);
        Assert.IsTrue(res);
        Assert.AreEqual(14, cellValB3);

        //--B5, set 45.88 - decimal
        res = excelApi.ExcelCellValueApi.SetCellValueDecimal(sheet, "B5", 45.88);
        Assert.IsTrue(res);

        double cellValB5;
        res = excelApi.ExcelCellValueApi.GetCellValueAsDecimal(sheet, "B5", out cellValB5);
        Assert.IsTrue(res);
        Assert.AreEqual(45.88, cellValB5);


        //--B7, set 21/09/1989 - DateShort
        res = excelApi.ExcelCellValueApi.SetCellValueDateShort(sheet, "B7", new DateTime(1989, 09, 21));
        Assert.IsTrue(res);

        DateTime cellValB7;
        res = excelApi.ExcelCellValueApi.GetCellValueAsDateTime(sheet, "B7", out cellValB7);
        Assert.IsTrue(res);
        Assert.AreEqual(1989, cellValB7.Year);
        Assert.AreEqual(9, cellValB7.Month);
        Assert.AreEqual(21, cellValB7.Day);

        //--B9, set xxx - Time
        //res = excelApi.ExcelCellValueApi.SetCellValueTime(sheet, "B7", new DateTime(2022, 05, 19));
        //Assert.IsTrue(res);

        // TODO: DateLarge, Time, currency,...
        //ici();

        //--close the file
        res = excelApi.ExcelFileApi.CloseExcelFile(workbook, out error);
        Assert.IsTrue(res);
    }

    [TestMethod]
    public void SetCellValuesOtherFormatNoStyle()
    {
        string fileName = @"Files\Cells\SetCellValuesOtherFormatNoStyle.xlsx";

        ExcelApi excelApi = new ExcelApi();
        ExcelWorkbook workbook;
        ExcelError error;
        bool res = excelApi.ExcelFileApi.OpenExcelFile(fileName, out workbook, out error);
        Assert.IsTrue(res);

        // get the first sheet
        var sheet = excelApi.ExcelSheetApi.GetSheet(workbook, 0);
        Assert.IsNotNull(sheet);

        //--B1 set 'salute' - general
        res = excelApi.ExcelCellValueApi.SetCellValueGeneral(sheet, "B1", "salute");
        Assert.IsTrue(res);

        // check
        string cellValB1 = excelApi.ExcelCellValueApi.GetCellValueAsString(sheet, "B1");
        Assert.AreEqual("salute", cellValB1);

        //--B3, set 14 - number 
        res = excelApi.ExcelCellValueApi.SetCellValueNumber(sheet, "B3", 14);
        Assert.IsTrue(res);

        int cellValB3;
        res = excelApi.ExcelCellValueApi.GetCellValueAsNumber(sheet, "B3", out cellValB3);
        Assert.IsTrue(res);
        Assert.AreEqual(14, cellValB3);

        //--B5, set 45.88 - decimal
        res = excelApi.ExcelCellValueApi.SetCellValueDecimal(sheet, "B5", 45.88);
        Assert.IsTrue(res);

        double cellValB5;
        res = excelApi.ExcelCellValueApi.GetCellValueAsDecimal(sheet, "B5", out cellValB5);
        Assert.IsTrue(res);
        Assert.AreEqual(45.88, cellValB5);


        //--B7, set 21/09/1989 - DateShort
        res = excelApi.ExcelCellValueApi.SetCellValueDateShort(sheet, "B7", new DateTime(1989, 09, 21));
        Assert.IsTrue(res);

        DateTime cellValB7;
        res = excelApi.ExcelCellValueApi.GetCellValueAsDateTime(sheet, "B7", out cellValB7);
        Assert.IsTrue(res);
        Assert.AreEqual(1989, cellValB7.Year);
        Assert.AreEqual(9, cellValB7.Month);
        Assert.AreEqual(21, cellValB7.Day);

        //--B9, set xxx - Time
        //res = excelApi.ExcelCellValueApi.SetCellValueTime(sheet, "B7", new DateTime(2022, 05, 19));
        //Assert.IsTrue(res);

        // TODO: DateLarge, Time, currency,...
        //ici();

        //--close the file
        res = excelApi.ExcelFileApi.CloseExcelFile(workbook, out error);
        Assert.IsTrue(res);
    }

    [TestMethod]
    public void SetCellValuesOtherFormatWithStyle()
    {
        Assert.Fail("todo:");
    }

    /// <summary>
    /// TODO: rework it!!
    /// Set cell value, many times with different type/value each time.
    /// TODO: SetCellValuesSameFormat
    /// </summary>
    [TestMethod]
    public void SetCellValuesReplace()
    {
        // the template file
        //string fileNameTempl = @"Files\Cells\InitSetManyCellType.xlsx";

        string fileName = @"Files\Cells\SetCellValuesReplace.xlsx";

        ExcelApi excelApi = new ExcelApi();
        ExcelWorkbook workbook;
        ExcelError error;
        bool res = excelApi.ExcelFileApi.OpenExcelFile(fileName, out workbook, out error);
        Assert.IsTrue(res);

        // get the first sheet
        var sheet = excelApi.ExcelSheetApi.GetSheet(workbook, 0);
        Assert.IsNotNull(sheet);

        //--B1 set 'bonjour' - general, the cell is empty/null
        res= excelApi.ExcelCellValueApi.SetCellValueGeneral(sheet, "B1", "bonjour");
        Assert.IsTrue(res);

        //--B3, replace 'salut' by 'coucou' - general
        res = excelApi.ExcelCellValueApi.SetCellValueGeneral(sheet, "B3", "coucou");
        Assert.IsTrue(res);

        string cellValB3 = excelApi.ExcelCellValueApi.GetCellValueAsString(sheet, "B3");
        Assert.AreEqual("coucou", cellValB3);

        //--B5, replace the number '12' by the text 'douze', no others format
        res = excelApi.ExcelCellValueApi.SetCellValueGeneral(sheet, "B5", "douze");
        Assert.IsTrue(res);

        string cellValB5 = excelApi.ExcelCellValueApi.GetCellValueAsString(sheet, "B5");
        Assert.AreEqual("douze", cellValB5);

        //--B7, replace the date short '19/09/1969' by the text 'heure', style already exists
        res = excelApi.ExcelCellValueApi.SetCellValueGeneral(sheet, "B7", "heure");
        Assert.IsTrue(res);

        string cellValB7 = excelApi.ExcelCellValueApi.GetCellValueAsString(sheet, "B7");
        Assert.AreEqual("heure", cellValB7);

        //--B9, replace the dateshortwith styles '10/03/1986' by the text 'bluetext', style doesn't exist in another cell
        res = excelApi.ExcelCellValueApi.SetCellValueGeneral(sheet, "B9", "bluetext");
        Assert.IsTrue(res);

        string cellValB9 = excelApi.ExcelCellValueApi.GetCellValueAsString(sheet, "B9");
        Assert.AreEqual("bluetext", cellValB9);

        //--B11, replace the formula, by the text 'formula', no style
        res = excelApi.ExcelCellValueApi.SetCellValueGeneral(sheet, "B11", "formula");
        Assert.IsTrue(res);

        string cellValB11 = excelApi.ExcelCellValueApi.GetCellValueAsString(sheet, "B11");
        Assert.AreEqual("formula", cellValB11);

        //--B13, set number=12, cell empty
        res = excelApi.ExcelCellValueApi.SetCellValueNumber(sheet, "B13", 12);
        Assert.IsTrue(res);

        int cellValB13; 
        res=excelApi.ExcelCellValueApi.GetCellValueAsNumber(sheet, "B13", out cellValB13);
        Assert.AreEqual(12, cellValB13);

        //--B15, replace number, set number=567, same cell format 
        res = excelApi.ExcelCellValueApi.SetCellValueNumber(sheet, "B15", 567);
        Assert.IsTrue(res);

        int cellValB15;
        res = excelApi.ExcelCellValueApi.GetCellValueAsNumber(sheet, "B15", out cellValB15);
        Assert.AreEqual(567, cellValB15);

        //--B17, replace general, set number=55
        res = excelApi.ExcelCellValueApi.SetCellValueNumber(sheet, "B17", 55);
        Assert.IsTrue(res);

        int cellValB17;
        res = excelApi.ExcelCellValueApi.GetCellValueAsNumber(sheet, "B17", out cellValB17);
        Assert.AreEqual(55, cellValB17);

        //--B19, replace dateShort, set number =67
        res = excelApi.ExcelCellValueApi.SetCellValueNumber(sheet, "B19", 67);
        Assert.IsTrue(res);

        int cellValB19;
        res = excelApi.ExcelCellValueApi.GetCellValueAsNumber(sheet, "B19", out cellValB19);
        Assert.AreEqual(67, cellValB19);

        //--B21, replace general with fill, style exists, by number=754
        res = excelApi.ExcelCellValueApi.SetCellValueNumber(sheet, "B21", 754);
        Assert.IsTrue(res);

        int cellValB21;
        res = excelApi.ExcelCellValueApi.GetCellValueAsNumber(sheet, "B21", out cellValB21);
        Assert.AreEqual(754, cellValB21);

        //--B23, replace accounting with fill and border, style doesn't exists, by number=890
        res = excelApi.ExcelCellValueApi.SetCellValueNumber(sheet, "B23", 890);
        Assert.IsTrue(res);

        int cellValB23;
        res = excelApi.ExcelCellValueApi.GetCellValueAsNumber(sheet, "B23", out cellValB23);
        Assert.AreEqual(890, cellValB23);

        //--B25, set cell decimal, cell doesn't exists, by decimal=12.34
        res = excelApi.ExcelCellValueApi.SetCellValueDecimal(sheet, "B25", 12.34);
        Assert.IsTrue(res);

        double cellValB25;
        res = excelApi.ExcelCellValueApi.GetCellValueAsDecimal(sheet, "B25", out cellValB25);
        Assert.AreEqual(12.34, cellValB25);

        //--B27: cell empty, set dateShort : 16/05/2022
        res = excelApi.ExcelCellValueApi.SetCellValueDateShort(sheet, "B27", new DateTime(2022,05,16));
        Assert.IsTrue(res);

        // check
        ExcelCellFormat cellFormatB27 = excelApi.ExcelCellValueApi.GetCellFormat(sheet, "B27");
        Assert.AreEqual(ExcelCellFormatCode.DateShort, cellFormatB27.Code);
        DateTime dtRes;
        excelApi.ExcelCellValueApi.GetCellValueAsDateTime(sheet, "B27", out dtRes);
        DateTime dtExpected = new DateTime(2022, 05, 16);
        Assert.AreEqual(dtExpected, dtRes);

        //XXXXXXTODO:
        // SetCellValueDateShort()
        // SetCellValueDateLarge()
        // SetCellValueTime()

        //--TODO: 

        //--TODO: faire SetCellValueText()

        //--close the file
        res = excelApi.ExcelFileApi.CloseExcelFile(workbook, out error);
        Assert.IsTrue(res);

    }
}
