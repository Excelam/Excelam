using Excelam.System;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Excelam.Tests;

/// <summary>
/// Test ExcelCellValueApi.GetCellFormat cases.
/// </summary>
[TestClass]
public class ExcelGetCellFormatTests
{
    [TestMethod]
    public void GetManyCellFormatValue()
    {
        ExcelApi excelApi = new ExcelApi();

        string fileName = @"Files\Cells\GetManyCellFormatValues.xlsx";
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

        //--B2: bonjour - standard/general  - col, row
        ExcelCellFormat cellFormatB2b = excelApi.ExcelCellValueApi.GetCellFormat(sheet, 2, 2);

        //--B3: 12 - number
        ExcelCellFormat cellFormatB3 = excelApi.ExcelCellValueApi.GetCellFormat(sheet, "B3");
        Assert.AreEqual(ExcelCellFormatCode.Number, cellFormatB3.Code);

        //--B15: 15/02/2021 - Short date
        ExcelCellFormat cellFormatB15 = excelApi.ExcelCellValueApi.GetCellFormat(sheet, "B15");
        Assert.AreEqual(ExcelCellFormatCode.DateShort, cellFormatB15.Code);

        //--B16: vendredi 19 septembre 1969 - date large
        ExcelCellFormat cellFormatB16 = excelApi.ExcelCellValueApi.GetCellFormat(sheet, "B16");
        Assert.AreEqual(ExcelCellFormatCode.DateLarge, cellFormatB16.Code);

        //--B17: 13:45:00 - time
        ExcelCellFormat cellFormatB17 = excelApi.ExcelCellValueApi.GetCellFormat(sheet, "B17");
        Assert.AreEqual(ExcelCellFormatCode.Time, cellFormatB17.Code);

        //--B21: 45,21 €  - accounting

        //--B22: 88,22 € - currency-euro

        //--B23: $91,25 - currency-dollarUS



        //--E7: 45.60 - decimal - Formula
        ExcelCellFormat cellFormatE7 = excelApi.ExcelCellValueApi.GetCellFormat(sheet, "E7");
        Assert.AreEqual(ExcelCellFormatCode.Decimal, cellFormatE7.Code);
        Assert.IsTrue(cellFormatE7.IsFormula);

        //--todo: faire autres cas: built-in: fraction, percetange, scientific,...

        //--todo: faire autres cas speciaux: date longue, accounting, currency....


        //--close the file
        res = excelApi.ExcelFileApi.CloseExcelFile(workbook, out error);
        Assert.IsTrue(res);

    }
}
