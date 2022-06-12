using Excelam.System;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using System;

namespace Excelam.Tests.GetCellValue;

/// <summary>
/// Get cell value tests.
/// </summary>
[TestClass]
public class GetCellValueTests
{
    [TestMethod]
    public void GetCellValuesGeneralText()
    {
        ExcelApi excelApi = new ExcelApi();

        string fileName = @"Files\GetCellValues\GetCellValuesGeneralText.xlsx";
        ExcelWorkbook workbook;
        ExcelError error;

        bool res = excelApi.ExcelFileApi.OpenExcelFile(fileName, out workbook, out error);
        Assert.IsTrue(res);
        Assert.IsNotNull(workbook);
        Assert.IsNull(error);

        // get the first sheet
        var sheet = excelApi.ExcelSheetApi.GetSheet(workbook, 0);
        Assert.IsNotNull(sheet);

        //--B1: null 
        string val = excelApi.ExcelCellValueApi.GetCellValueAsString(sheet, "B1");
        Assert.IsNull(val);

        //--B3: bonjour - general
        val = excelApi.ExcelCellValueApi.GetCellValueAsString(sheet, "B3");
        Assert.AreEqual("bonjour", val);

        //--B5: it's a text - text
        val = excelApi.ExcelCellValueApi.GetCellValueAsString(sheet, "B5");
        Assert.AreEqual("it's a text", val);

        //--close the file
        res = excelApi.ExcelFileApi.CloseExcelFile(workbook, out error);
        Assert.IsTrue(res);
    }

    [TestMethod]
    public void GetCellValuesDecimal()
    {
        ExcelApi excelApi = new ExcelApi();

        string fileName = @"Files\GetCellValues\GetCellValuesDecimal.xlsx";
        ExcelWorkbook workbook;
        ExcelError error;

        bool res = excelApi.ExcelFileApi.OpenExcelFile(fileName, out workbook, out error);
        Assert.IsTrue(res);
        Assert.IsNotNull(workbook);
        Assert.IsNull(error);

        // get the first sheet
        var sheet = excelApi.ExcelSheetApi.GetSheet(workbook, 0);
        Assert.IsNotNull(sheet);

        //--B1: 12 - number
        int valInt;
        res = excelApi.ExcelCellValueApi.GetCellValueAsInt(sheet, "B1", out valInt);
        Assert.IsTrue(res);
        Assert.AreEqual(12, valInt);

        //--B3: 22,56 - decimal, a built-in format
        double valDbl;
        res = excelApi.ExcelCellValueApi.GetCellValueAsDouble(sheet, "B3", out valDbl);
        Assert.IsTrue(res);
        Assert.AreEqual(22.56, valDbl);

        //--B5: 63,456 - decimal - 3dec
        res = excelApi.ExcelCellValueApi.GetCellValueAsDouble(sheet, "B5", out valDbl);
        Assert.IsTrue(res);
        Assert.AreEqual(63.456, valDbl);

        //--B7: 5,6 - decimal - 1dec
        res = excelApi.ExcelCellValueApi.GetCellValueAsDouble(sheet, "B7", out valDbl);
        Assert.IsTrue(res);
        Assert.AreEqual(5.6, valDbl);

        //--B9: 123 - decimal - neg, red, no sign, format: "0.00;[Red]0.00"
        res = excelApi.ExcelCellValueApi.GetCellValueAsDouble(sheet, "B9", out valDbl);
        Assert.IsTrue(res);
        Assert.AreEqual(-123, valDbl);

        //--B11: -123 - decimal - neg, red, sign, format: "0.00_ ;[Red]\\-0.00\\ "
        res = excelApi.ExcelCellValueApi.GetCellValueAsDouble(sheet, "B11", out valDbl);
        Assert.IsTrue(res);
        Assert.AreEqual(-123, valDbl);

        //--B13: 123 000,50 -decimal, 2 dec. thousand sep, format: ?
        res = excelApi.ExcelCellValueApi.GetCellValueAsDouble(sheet, "B13", out valDbl);
        Assert.IsTrue(res);
        Assert.AreEqual(123000.5, valDbl);

        //--close the file
        res = excelApi.ExcelFileApi.CloseExcelFile(workbook, out error);
        Assert.IsTrue(res);
    }


    [TestMethod]
    public void GetCellValuesDateTime()
    {
        ExcelApi excelApi = new ExcelApi();

        string fileName = @"Files\GetCellValues\GetCellValuesDateTime.xlsx";
        ExcelWorkbook workbook;
        ExcelError error;

        bool res = excelApi.ExcelFileApi.OpenExcelFile(fileName, out workbook, out error);
        Assert.IsTrue(res);
        Assert.IsNotNull(workbook);
        Assert.IsNull(error);

        // get the first sheet
        var sheet = excelApi.ExcelSheetApi.GetSheet(workbook, 0);
        Assert.IsNotNull(sheet);

        //--B1: 15/02/2022 - DateShort14
        DateTime valdt;
        res = excelApi.ExcelCellValueApi.GetCellValueAsDateTime(sheet, "B1", out valdt);
        Assert.IsTrue(res);
        Assert.AreEqual(15, valdt.Day);
        Assert.AreEqual(2, valdt.Month);
        Assert.AreEqual(2022, valdt.Year);

        //--B3: 14:10:25 - Time21_hh_mm_ss
        res = excelApi.ExcelCellValueApi.GetCellValueAsDateTime(sheet, "B3", out valdt);
        Assert.IsTrue(res);
        Assert.AreEqual(14, valdt.Hour);
        Assert.AreEqual(10, valdt.Minute);
        Assert.AreEqual(25, valdt.Second);

        //--B5: lundi 18 juin 1945 - DateLarge
        res = excelApi.ExcelCellValueApi.GetCellValueAsDateTime(sheet, "B5", out valdt);
        Assert.IsTrue(res);
        Assert.AreEqual(18, valdt.Day);
        Assert.AreEqual(6, valdt.Month);
        Assert.AreEqual(1945, valdt.Year);

        //--close the file
        res = excelApi.ExcelFileApi.CloseExcelFile(workbook, out error);
        Assert.IsTrue(res);
    }


    [TestMethod]
    public void GetCellValuesCurrency()
    {
        ExcelApi excelApi = new ExcelApi();

        string fileName = @"Files\GetCellValues\GetCellValuesCurrency.xlsx";
        ExcelWorkbook workbook;
        ExcelError error;

        bool res = excelApi.ExcelFileApi.OpenExcelFile(fileName, out workbook, out error);
        Assert.IsTrue(res);
        Assert.IsNotNull(workbook);
        Assert.IsNull(error);

        // get the first sheet
        var sheet = excelApi.ExcelSheetApi.GetSheet(workbook, 0);
        Assert.IsNotNull(sheet);

        //--B1: 88,22 € - currency-euro
        double valB22;
        res = excelApi.ExcelCellValueApi.GetCellValueAsDouble(sheet, "B1", out valB22);
        Assert.IsTrue(res);
        Assert.AreEqual(88.22, valB22);

        //--B3: $91,25 - currency-dollarUS
        res = excelApi.ExcelCellValueApi.GetCellValueAsDouble(sheet, "B1", out valB22);
        Assert.IsTrue(res);
        Assert.AreEqual(88.22, valB22);


        //--TODO: check others

        //--close the file
        res = excelApi.ExcelFileApi.CloseExcelFile(workbook, out error);
        Assert.IsTrue(res);
    }

    // GetCellValuesAccounting()
    // TODO:

    // GetCellValuesFraction()
    // TODO:

    // GetCellValuesPercentage()
    // TODO:

    // GetCellValuesScientific()
    // TODO:
}
