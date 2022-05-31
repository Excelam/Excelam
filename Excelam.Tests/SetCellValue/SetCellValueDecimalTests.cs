using Excelam.System;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Excelam.Tests.SetCellValue;

[TestClass]
public class SetCellValueDecimalTests
{
    /// <summary>
    /// Start from an empty excel file.
    /// set cell values, test all formats.
    /// </summary>
    [TestMethod]
    public void SetCellValuesEmpty()
    {
        string fileName = @"Files\SetCellValues\SetCellValuesDecimalEmpty.xlsx";

        ExcelApi excelApi = new ExcelApi();
        ExcelWorkbook workbook;
        ExcelError error;
        bool res = excelApi.ExcelFileApi.OpenExcelFile(fileName, out workbook, out error);
        Assert.IsTrue(res);

        // get the first sheet
        var sheet = excelApi.ExcelSheetApi.GetSheet(workbook, 0);
        Assert.IsNotNull(sheet);

        //--B1: 22,56 - decimal, a built-in format
        res = excelApi.ExcelCellValueApi.SetCellValueDecimal(sheet, "B1", ExcelCellDecimalCode.Decimal, 2, 22.56); 
        Assert.IsTrue(res);

        // check
        double cellVal;
        res = excelApi.ExcelCellValueApi.GetCellValueAsDecimal(sheet, "B1", out cellVal);
        Assert.IsTrue(res);
        Assert.AreEqual(22.56, cellVal);


        //--B3: 63,456 - decimal - 3dec
        res = excelApi.ExcelCellValueApi.SetCellValueDecimal(sheet, "B3", ExcelCellDecimalCode.Decimal, 3, 63.456);
        Assert.IsTrue(res);

        // check
        var cellFormat = excelApi.ExcelCellValueApi.GetCellFormat(sheet, "B3");
        Assert.AreEqual(ExcelCellFormatValueCode.Decimal, cellFormat.FormatValue.Code);
        Assert.AreEqual("0.000", cellFormat.FormatValue.StringFormat);
        var cellFormatValueDecimal = cellFormat.GetFormatValueAsDecimal();
        Assert.IsNotNull(cellFormatValueDecimal);
        Assert.AreEqual(ExcelCellDecimalCode.Decimal, cellFormatValueDecimal.SubCode);
        Assert.AreEqual(3, cellFormatValueDecimal.NumberOfDecimal);

        res = excelApi.ExcelCellValueApi.GetCellValueAsDecimal(sheet, "B3", out cellVal);
        Assert.IsTrue(res);
        Assert.AreEqual(63.456, cellVal);


        //--B5: 5,6 - decimal - 1dec
        res = excelApi.ExcelCellValueApi.SetCellValueDecimal(sheet, "B5", ExcelCellDecimalCode.Decimal, 1, 5.6);
        Assert.IsTrue(res);

        // check
        cellFormat = excelApi.ExcelCellValueApi.GetCellFormat(sheet, "B5");
        Assert.AreEqual(ExcelCellFormatValueCode.Decimal, cellFormat.FormatValue.Code);
        Assert.AreEqual("0.0", cellFormat.FormatValue.StringFormat);
        cellFormatValueDecimal = cellFormat.GetFormatValueAsDecimal();
        Assert.IsNotNull(cellFormatValueDecimal);
        Assert.AreEqual(ExcelCellDecimalCode.Decimal, cellFormatValueDecimal.SubCode);
        Assert.AreEqual(1, cellFormatValueDecimal.NumberOfDecimal);

        res = excelApi.ExcelCellValueApi.GetCellValueAsDecimal(sheet, "B5", out cellVal);
        Assert.IsTrue(res);
        Assert.AreEqual(5.6, cellVal);

        //--B7: 123 - decimal - neg, red, no sign, format: "0.00;[Red]0.00"
        res = excelApi.ExcelCellValueApi.SetCellValueDecimal(sheet, "B7", ExcelCellDecimalCode.DecimalNegRedNoSign, 2, -123);
        Assert.IsTrue(res);

        // check
        cellFormat = excelApi.ExcelCellValueApi.GetCellFormat(sheet, "B7");
        Assert.AreEqual(ExcelCellFormatValueCode.Decimal, cellFormat.FormatValue.Code);
        Assert.AreEqual("0.00;[Red]0.00", cellFormat.FormatValue.StringFormat);
        cellFormatValueDecimal = cellFormat.GetFormatValueAsDecimal();
        Assert.IsNotNull(cellFormatValueDecimal);
        Assert.AreEqual(ExcelCellDecimalCode.DecimalNegRedNoSign, cellFormatValueDecimal.SubCode);
        Assert.AreEqual(2, cellFormatValueDecimal.NumberOfDecimal);

        res = excelApi.ExcelCellValueApi.GetCellValueAsDecimal(sheet, "B7", out cellVal);
        Assert.IsTrue(res);
        Assert.AreEqual(-123, cellVal);

        //--B9: -123 - decimal - neg, red, sign, format: "0.00_ ;[Red]\\-0.00\\ "
        res = excelApi.ExcelCellValueApi.SetCellValueDecimal(sheet, "B9", ExcelCellDecimalCode.DecimalNegRed, 2, -123);
        Assert.IsTrue(res);

        // check
        cellFormat = excelApi.ExcelCellValueApi.GetCellFormat(sheet, "B9");
        Assert.AreEqual(ExcelCellFormatValueCode.Decimal, cellFormat.FormatValue.Code);
        Assert.AreEqual(@"0.00_ ;[Red]\\-0.00\\ ", cellFormat.FormatValue.StringFormat);
        cellFormatValueDecimal = cellFormat.GetFormatValueAsDecimal();
        Assert.IsNotNull(cellFormatValueDecimal);
        Assert.AreEqual(ExcelCellDecimalCode.DecimalNegRed, cellFormatValueDecimal.SubCode);
        Assert.AreEqual(2, cellFormatValueDecimal.NumberOfDecimal);

        res = excelApi.ExcelCellValueApi.GetCellValueAsDecimal(sheet, "B9", out cellVal);
        Assert.IsTrue(res);
        Assert.AreEqual(-123, cellVal);

        //--B11: 123 000,50 -decimal, 2 dec. thousand sep, format: null
        res = excelApi.ExcelCellValueApi.SetCellValueDecimal(sheet, "B11", ExcelCellDecimalCode.DecimalBlankThousandSep, 2, 123000.50);
        Assert.IsTrue(res);

        // check
        cellFormat = excelApi.ExcelCellValueApi.GetCellFormat(sheet, "B11");
        Assert.AreEqual(ExcelCellFormatValueCode.Decimal, cellFormat.FormatValue.Code);
        Assert.AreEqual(String.Empty, cellFormat.FormatValue.StringFormat);
        cellFormatValueDecimal = cellFormat.GetFormatValueAsDecimal();
        Assert.IsNotNull(cellFormatValueDecimal);
        Assert.AreEqual(ExcelCellDecimalCode.DecimalBlankThousandSep, cellFormatValueDecimal.SubCode);
        Assert.AreEqual(2, cellFormatValueDecimal.NumberOfDecimal);

        res = excelApi.ExcelCellValueApi.GetCellValueAsDecimal(sheet, "B11", out cellVal);
        Assert.IsTrue(res);
        Assert.AreEqual(123000.50, cellVal);

        //--close the file
        res = excelApi.ExcelFileApi.CloseExcelFile(workbook, out error);
        Assert.IsTrue(res);
    }

    /// <summary>
    /// reuse the excel file SetCellValuesEmpty.xlsx
    /// generated by the previous test.
    /// </summary>
    [TestMethod]
    public void SetCellValuesSameFormatNoStyle()
    {
        string fileName = @"Files\SetCellValues\SetCellValuesDecimalSameFormatNoStyle.xlsx";
        Assert.Fail("todo");
    }

    /// <summary>
    /// reuse the excel file SetCellValuesEmpty.xlsx
    /// generated by the previous test.
    /// </summary>
    [TestMethod]
    public void SetCellValuesSameFormatWithStyle()
    {
        string fileName = @"Files\SetCellValues\SetCellValuesDecimalSameFormatWithStyle.xlsx";
        Assert.Fail("todo");
    }


    [TestMethod]
    public void SetCellValuesOtherFormatNoStyle()
    {
        string fileName = @"Files\SetCellValues\SetCellValuesDecimalOtherFormatNoStyle.xlsx";
        Assert.Fail("todo");
    }

    [TestMethod]
    public void SetCellValuesOtherFormatWithStyle()
    {
        string fileName = @"Files\SetCellValues\SetCellValuesDecimalOtherFormatWithStyle.xlsx";
        Assert.Fail("todo");
    }

}
