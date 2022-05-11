using Excelam.System;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using System;

namespace Excelam.Tests;

/// <summary>
/// Get cell value tests.
/// TODO: refactor! only GetCellValue
/// </summary>
[TestClass]
public class ExcelCellValueApiGetValueTests
{
    [TestMethod]
    public void GetManyCellValue()
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
        string valB2b = excelApi.ExcelCellValueApi.GetCellValueAsString(sheet, 2,2);
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
        DateTime dtB17 = new DateTime(valB17.Year, valB17.Month, valB17.Day, 13,45,0);
        Assert.AreEqual(dtB17, valB17);

        //--B21: 45,21 €  - accounting

        //--B22: 88,22 € - currency-euro

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
