using Excelam.System;
using Microsoft.VisualStudio.TestTools.UnitTesting;

namespace Excelam.Tests;

/// <summary>
/// Get cell value tests.
/// </summary>
[TestClass]
public class ExcelCellValueApiGetValueTests
{
    [TestMethod]
    public void GetManyCellFormatValue()
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

        //--B1: null
        ExcelCellFormat cellFormatB1 = excelApi.ExcelCellValueApi.GetCellFormat(sheet, "B1");
        Assert.IsNull(cellFormatB1);

        //--B2: bonjour - standard/general
        ExcelCellFormat cellFormatB2 = excelApi.ExcelCellValueApi.GetCellFormat(sheet, "B2");
        Assert.AreEqual(ExcelCellFormatCode.General, cellFormatB2.Code);
        string valB2 = excelApi.ExcelCellValueApi.GetCellValueAsString(sheet, "B2");
        Assert.AreEqual("bonjour", valB2);

        //--B2: bonjour - standard/general  - col, row
        ExcelCellFormat cellFormatB2b = excelApi.ExcelCellValueApi.GetCellFormat(sheet, 2,2);
        Assert.AreEqual(ExcelCellFormatCode.General, cellFormatB2b.Code);
        string valB2b = excelApi.ExcelCellValueApi.GetCellValueAsString(sheet, 2,2);
        Assert.AreEqual("bonjour", valB2b);

        //--B2: bonjour - standard/general  
        int valB2int;
        res = excelApi.ExcelCellValueApi.GetCellValueAsNumber(sheet, "B2", out valB2int);
        Assert.IsFalse(res);

        //--B3: 12 - number
        ExcelCellFormat cellFormatB3 = excelApi.ExcelCellValueApi.GetCellFormat(sheet, "B3");
        Assert.AreEqual(ExcelCellFormatCode.Number, cellFormatB3.Code);
        int valB3;
        res = excelApi.ExcelCellValueApi.GetCellValueAsNumber(sheet, "B3", out valB3);
        Assert.IsTrue(res);
        Assert.AreEqual(12, valB3);

        //--E7: 45.60 - decimal - Formula
        ExcelCellFormat cellFormatE7 = excelApi.ExcelCellValueApi.GetCellFormat(sheet, "E7");
        Assert.IsTrue(cellFormatE7.IsFormula);
        Assert.AreEqual("SUM(E5:E6)", excelApi.ExcelCellValueApi.GetCellFormula(sheet, "E7"));

        Assert.AreEqual(ExcelCellFormatCode.Decimal, cellFormatE7.Code);
        string valE7Str = excelApi.ExcelCellValueApi.GetCellValueAsString(sheet, "E7");
        Assert.AreEqual("45.6", valE7Str);
        double valE7;
        excelApi.ExcelCellValueApi.GetCellValueAsDecimal(sheet, "E7", out valE7);
        Assert.AreEqual(45.6, valE7);



        //--todo: GetCellValueAsDateShort()


        //--todo: faire autres cas: built-in: fraction, percetange, scientific,...

        //--todo: faire autres cas speciaux: date longue, accounting, currency....


        //--close the file
        res = excelApi.ExcelFileApi.CloseExcelFile(workbook, out error);
        Assert.IsTrue(res);
    }
}
