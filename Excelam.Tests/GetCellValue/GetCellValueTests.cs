using Excelam.System;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using System;

namespace Excelam.Tests.GetCellValue;

/// <summary>
/// Get cell value tests.
/// TODO: refactor! only GetCellValue
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
        res = excelApi.ExcelCellValueApi.GetCellValueAsNumber(sheet, "B1", out valInt);
        Assert.IsTrue(res);
        Assert.AreEqual(12, valInt);

        //--B3: 22,56 - decimal, a built-in format
        double valDbl;
        res = excelApi.ExcelCellValueApi.GetCellValueAsDecimal(sheet, "B3", out valDbl);
        Assert.IsTrue(res);
        Assert.AreEqual(22.56, valDbl);

        //--B5: 63,456 - decimal - 3dec
        res = excelApi.ExcelCellValueApi.GetCellValueAsDecimal(sheet, "B5", out valDbl);
        Assert.IsTrue(res);
        Assert.AreEqual(63.456, valDbl);

        //--B7: 5,6 - decimal - 1dec
        res = excelApi.ExcelCellValueApi.GetCellValueAsDecimal(sheet, "B7", out valDbl);
        Assert.IsTrue(res);
        Assert.AreEqual(5.6, valDbl);

        //--B9: 123 - decimal - neg, red, no sign, format: "0.00;[Red]0.00"
        res = excelApi.ExcelCellValueApi.GetCellValueAsDecimal(sheet, "B9", out valDbl);
        Assert.IsTrue(res);
        Assert.AreEqual(-123, valDbl);

        //--B11: -123 - decimal - neg, red, sign, format: "0.00_ ;[Red]\\-0.00\\ "
        res = excelApi.ExcelCellValueApi.GetCellValueAsDecimal(sheet, "B11", out valDbl);
        Assert.IsTrue(res);
        Assert.AreEqual(-123, valDbl);

        //--B13: 123 000,50 -decimal, 2 dec. thousand sep, format: ?
        res = excelApi.ExcelCellValueApi.GetCellValueAsDecimal(sheet, "B13", out valDbl);
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

        // TODO: fichier n'existe pas!!!
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
        res = excelApi.ExcelCellValueApi.GetCellValueAsDecimal(sheet, "B1", out valB22);
        Assert.IsTrue(res);
        Assert.AreEqual(88.22, valB22);

        //--B3: $91,25 - currency-dollarUS
        res = excelApi.ExcelCellValueApi.GetCellValueAsDecimal(sheet, "B1", out valB22);
        Assert.IsTrue(res);
        Assert.AreEqual(88.22, valB22);

        //--close the file
        res = excelApi.ExcelFileApi.CloseExcelFile(workbook, out error);
        Assert.IsTrue(res);
    }

    // GetCellValuesAccounting()
    // TODO:

    /// <summary>
    /// TODO: refactor it, split in several methods.
    /// GetCellValuesBuiltIn, ...
    /// </summary>
    [TestMethod]
    public void GetCellValues()
    {
        ExcelApi excelApi = new ExcelApi();

        string fileName = @"Files\Cells\GetManyCellTypes.xlsx";
        ExcelWorkbook workbook;
        ExcelError error;

        bool res = excelApi.ExcelFileApi.OpenExcelFile(fileName, out workbook, out error);
        Assert.IsTrue(res);
        Assert.IsNotNull(workbook);
        Assert.IsNull(error);

        // get the first sheet
        var sheet = excelApi.ExcelSheetApi.GetSheet(workbook, 0);
        Assert.IsNotNull(sheet);

        //--B1: null (ou empty??)
        string valB1 = excelApi.ExcelCellValueApi.GetCellValueAsString(sheet, "B1");
        Assert.IsNull(valB1);

        //--B2: bonjour - standard/general
        string valB2 = excelApi.ExcelCellValueApi.GetCellValueAsString(sheet, "B2");
        Assert.AreEqual("bonjour", valB2);

        //--B2: bonjour - standard/general  - col, row
        string valB2b = excelApi.ExcelCellValueApi.GetCellValueAsString(sheet, 2, 2);
        Assert.AreEqual("bonjour", valB2b);

        //--B2: bonjour - standard/general  
        int valB2int;
        res = excelApi.ExcelCellValueApi.GetCellValueAsNumber(sheet, "B2", out valB2int);
        Assert.IsFalse(res);

        //--B3: 12 - number
        int valB3;
        res = excelApi.ExcelCellValueApi.GetCellValueAsNumber(sheet, "B3", out valB3);
        Assert.IsTrue(res);
        Assert.AreEqual(12, valB3);

        //--B15: 15/02/2021 - Short date
        DateTime valB15;
        res = excelApi.ExcelCellValueApi.GetCellValueAsDateTime(sheet, "B15", out valB15);
        Assert.IsTrue(res);
        DateTime dtB15 = new DateTime(2021, 02, 15);
        Assert.AreEqual(dtB15, valB15);

        //--B16: vendredi 19 septembre 1969 - date large
        DateTime valB16;
        res = excelApi.ExcelCellValueApi.GetCellValueAsDateTime(sheet, "B16", out valB16);
        Assert.IsTrue(res);
        DateTime dtB16 = new DateTime(1969, 09, 19);
        Assert.AreEqual(dtB16, valB16);

        //--B17: 13:45:00 - time
        DateTime valB17;
        res = excelApi.ExcelCellValueApi.GetCellValueAsDateTime(sheet, "B17", out valB17);
        Assert.IsTrue(res);
        // ! for an excel cell time value, check only hour, min and sec.
        DateTime dtB17 = new DateTime(valB17.Year, valB17.Month, valB17.Day, 13, 45, 0);
        Assert.AreEqual(dtB17, valB17);

        //--B21: 45,21 €  - accounting
        double valB21;
        res = excelApi.ExcelCellValueApi.GetCellValueAsDecimal(sheet, "B21", out valB21);
        Assert.IsTrue(res);
        Assert.AreEqual(45.21, valB21);

        //--B22: 88,22 € - currency-euro
        double valB22;
        res = excelApi.ExcelCellValueApi.GetCellValueAsDecimal(sheet, "B22", out valB22);
        Assert.IsTrue(res);
        Assert.AreEqual(88.22, valB22);

        //--B23: $91,25 - currency-dollarUS



        //--E7: 45.60 - decimal - Formula
        ExcelCellFormat cellFormatE7 = excelApi.ExcelCellValueApi.GetCellFormat(sheet, "E7");
        Assert.IsTrue(cellFormatE7.IsFormula);
        Assert.AreEqual("SUM(E5:E6)", excelApi.ExcelCellValueApi.GetCellFormula(sheet, "E7"));

        string valE7Str = excelApi.ExcelCellValueApi.GetCellValueAsString(sheet, "E7");
        Assert.AreEqual("45.6", valE7Str);
        double valE7;
        excelApi.ExcelCellValueApi.GetCellValueAsDecimal(sheet, "E7", out valE7);
        Assert.AreEqual(45.6, valE7);




        //--todo: faire autres cas: built-in: fraction, percetange, scientific,...

        //--todo: faire autres cas speciaux: date longue, accounting, currency....


        //--close the file
        res = excelApi.ExcelFileApi.CloseExcelFile(workbook, out error);
        Assert.IsTrue(res);
    }
}
