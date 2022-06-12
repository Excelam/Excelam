using Excelam.System;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using System;
using System.Collections.Generic;
using System.IO;
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
    public void SetCellValuesNew()
    {
        string fileName = @"Files\SetCellValues\SetCellValuesDecimalNew.xlsx";

        if (File.Exists(fileName))
            File.Delete(fileName);

        ExcelApi excelApi = new ExcelApi();
        ExcelWorkbook workbook;
        ExcelError error;
        bool res = excelApi.ExcelFileApi.CreateExcelFile(fileName, excelApi.ExcelFileApi.DefaultFirstSheetName, out workbook, out error);
        Assert.IsTrue(res);

        // get the first sheet
        var sheet = excelApi.ExcelSheetApi.GetSheet(workbook, 0);
        Assert.IsNotNull(sheet);

        //--B1: 22,56 - decimal, a built-in format
        res = excelApi.ExcelCellValueApi.SetCellValueDecimal(sheet, "B1", 2, false, ExcelCellValueNegativeOption.Default, 22.56);
        Assert.IsTrue(res);

        // check
        double cellVal;
        res = excelApi.ExcelCellValueApi.GetCellValueAsDouble(sheet, "B1", out cellVal);
        Assert.IsTrue(res);
        Assert.AreEqual(22.56, cellVal);


        //--B3: 63,456 - decimal - 3dec
        res = excelApi.ExcelCellValueApi.SetCellValueDecimal(sheet, "B3", 3, false, ExcelCellValueNegativeOption.Default, 63.456);
        Assert.IsTrue(res);

        // check
        var cellFormat = excelApi.ExcelCellValueApi.GetCellFormat(sheet, "B3");
        Assert.AreEqual(ExcelCellFormatValueCode.Decimal, cellFormat.FormatValue.Code);
        Assert.AreEqual("0.000", cellFormat.FormatValue.StringFormat);
        var cellFormatValueDecimal = cellFormat.GetFormatValueAsDecimal();
        Assert.IsNotNull(cellFormatValueDecimal);
        Assert.AreEqual(ExcelCellDecimalCode.DecimalN, cellFormatValueDecimal.DecimalCode);
        Assert.AreEqual(3, cellFormatValueDecimal.NumberOfDecimal);
        Assert.AreEqual(ExcelCellValueNegativeOption.Default, cellFormatValueDecimal.NegativeOption);

        res = excelApi.ExcelCellValueApi.GetCellValueAsDouble(sheet, "B3", out cellVal);
        Assert.IsTrue(res);
        Assert.AreEqual(63.456, cellVal);


        //--B5: 5,6 - decimal - 1dec
        res = excelApi.ExcelCellValueApi.SetCellValueDecimal(sheet, "B5", 1, false, ExcelCellValueNegativeOption.Default, 5.6);
        Assert.IsTrue(res);

        // check
        cellFormat = excelApi.ExcelCellValueApi.GetCellFormat(sheet, "B5");
        Assert.AreEqual(ExcelCellFormatValueCode.Decimal, cellFormat.FormatValue.Code);
        Assert.AreEqual("0.0", cellFormat.FormatValue.StringFormat);
        cellFormatValueDecimal = cellFormat.GetFormatValueAsDecimal();
        Assert.IsNotNull(cellFormatValueDecimal);
        Assert.AreEqual(ExcelCellDecimalCode.DecimalN, cellFormatValueDecimal.DecimalCode);
        Assert.AreEqual(1, cellFormatValueDecimal.NumberOfDecimal);
        Assert.AreEqual(ExcelCellValueNegativeOption.Default, cellFormatValueDecimal.NegativeOption);

        res = excelApi.ExcelCellValueApi.GetCellValueAsDouble(sheet, "B5", out cellVal);
        Assert.IsTrue(res);
        Assert.AreEqual(5.6, cellVal);

        //--B7: 123 - decimal - neg, red, no sign, format: "0.00;[Red]0.00"
        res = excelApi.ExcelCellValueApi.SetCellValueDecimal(sheet, "B7", 2, false, ExcelCellValueNegativeOption.RedWithoutSign, -123);
        Assert.IsTrue(res);

        // check
        cellFormat = excelApi.ExcelCellValueApi.GetCellFormat(sheet, "B7");
        Assert.AreEqual(ExcelCellFormatValueCode.Decimal, cellFormat.FormatValue.Code);
        Assert.AreEqual("0.00;[Red]0.00", cellFormat.FormatValue.StringFormat);
        cellFormatValueDecimal = cellFormat.GetFormatValueAsDecimal();
        Assert.IsNotNull(cellFormatValueDecimal);
        Assert.AreEqual(ExcelCellDecimalCode.DecimalN, cellFormatValueDecimal.DecimalCode);
        Assert.AreEqual(2, cellFormatValueDecimal.NumberOfDecimal);
        Assert.AreEqual(ExcelCellValueNegativeOption.RedWithoutSign, cellFormatValueDecimal.NegativeOption);

        res = excelApi.ExcelCellValueApi.GetCellValueAsDouble(sheet, "B7", out cellVal);
        Assert.IsTrue(res);
        Assert.AreEqual(-123, cellVal);

        //--B9: -123 - decimal - neg, red, sign, format: "0.00_ ;[Red]\\-0.00\\ "
        res = excelApi.ExcelCellValueApi.SetCellValueDecimal(sheet, "B9", 2, false, ExcelCellValueNegativeOption.RedWithSign , -123);
        Assert.IsTrue(res);

        // check
        cellFormat = excelApi.ExcelCellValueApi.GetCellFormat(sheet, "B9");
        Assert.AreEqual(ExcelCellFormatValueCode.Decimal, cellFormat.FormatValue.Code);
        Assert.AreEqual("0.00_ ;[Red]\\-0.00\\ ", cellFormat.FormatValue.StringFormat);
        cellFormatValueDecimal = cellFormat.GetFormatValueAsDecimal();
        Assert.IsNotNull(cellFormatValueDecimal);
        Assert.AreEqual(ExcelCellDecimalCode.DecimalN, cellFormatValueDecimal.DecimalCode);
        Assert.AreEqual(2, cellFormatValueDecimal.NumberOfDecimal);
        Assert.AreEqual(ExcelCellValueNegativeOption.RedWithSign, cellFormatValueDecimal.NegativeOption);

        res = excelApi.ExcelCellValueApi.GetCellValueAsDouble(sheet, "B9", out cellVal);
        Assert.IsTrue(res);
        Assert.AreEqual(-123, cellVal);

        //--B11: 123 000,50 -decimal, 2 dec. thousand sep, format: null
        res = excelApi.ExcelCellValueApi.SetCellValueDecimal(sheet, "B11", 2, true, ExcelCellValueNegativeOption.Default, 123000.50);
        Assert.IsTrue(res);

        // check
        cellFormat = excelApi.ExcelCellValueApi.GetCellFormat(sheet, "B11");
        Assert.AreEqual(ExcelCellFormatValueCode.Decimal, cellFormat.FormatValue.Code);
        Assert.AreEqual(String.Empty, cellFormat.FormatValue.StringFormat);
        cellFormatValueDecimal = cellFormat.GetFormatValueAsDecimal();
        Assert.IsNotNull(cellFormatValueDecimal);
        Assert.AreEqual(ExcelCellDecimalCode.Decimal4BlankThousandSep, cellFormatValueDecimal.DecimalCode);
        Assert.AreEqual(2, cellFormatValueDecimal.NumberOfDecimal);
        Assert.AreEqual(ExcelCellValueNegativeOption.Default, cellFormatValueDecimal.NegativeOption);

        res = excelApi.ExcelCellValueApi.GetCellValueAsDouble(sheet, "B11", out cellVal);
        Assert.IsTrue(res);
        Assert.AreEqual(123000.50, cellVal);

        //--close the file
        res = excelApi.ExcelFileApi.CloseExcelFile(workbook, out error);
        Assert.IsTrue(res);
    }


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
        res = excelApi.ExcelCellValueApi.SetCellValueDecimal(sheet, "B1", 2, false, ExcelCellValueNegativeOption.Default, 22.56);

        Assert.IsTrue(res);

        // check
        double cellVal;
        res = excelApi.ExcelCellValueApi.GetCellValueAsDouble(sheet, "B1", out cellVal);
        Assert.IsTrue(res);
        Assert.AreEqual(22.56, cellVal);


        //--B3: 63,456 - decimal - 3dec
        res = excelApi.ExcelCellValueApi.SetCellValueDecimal(sheet, "B3", 3, false, ExcelCellValueNegativeOption.Default, 63.456);
        Assert.IsTrue(res);

        // check
        var cellFormat = excelApi.ExcelCellValueApi.GetCellFormat(sheet, "B3");
        Assert.AreEqual(ExcelCellFormatValueCode.Decimal, cellFormat.FormatValue.Code);
        Assert.AreEqual("0.000", cellFormat.FormatValue.StringFormat);
        var cellFormatValueDecimal = cellFormat.GetFormatValueAsDecimal();
        Assert.IsNotNull(cellFormatValueDecimal);
        Assert.AreEqual(ExcelCellDecimalCode.DecimalN, cellFormatValueDecimal.DecimalCode);
        Assert.AreEqual(3, cellFormatValueDecimal.NumberOfDecimal);
        Assert.AreEqual(ExcelCellValueNegativeOption.Default, cellFormatValueDecimal.NegativeOption);

        res = excelApi.ExcelCellValueApi.GetCellValueAsDouble(sheet, "B3", out cellVal);
        Assert.IsTrue(res);
        Assert.AreEqual(63.456, cellVal);


        //--B5: 5,6 - decimal - 1dec
        res = excelApi.ExcelCellValueApi.SetCellValueDecimal(sheet, "B5", 1, false, ExcelCellValueNegativeOption.Default, 5.6);
        Assert.IsTrue(res);

        // check
        cellFormat = excelApi.ExcelCellValueApi.GetCellFormat(sheet, "B5");
        Assert.AreEqual(ExcelCellFormatValueCode.Decimal, cellFormat.FormatValue.Code);
        Assert.AreEqual("0.0", cellFormat.FormatValue.StringFormat);
        cellFormatValueDecimal = cellFormat.GetFormatValueAsDecimal();
        Assert.IsNotNull(cellFormatValueDecimal);
        Assert.AreEqual(ExcelCellDecimalCode.DecimalN, cellFormatValueDecimal.DecimalCode);
        Assert.AreEqual(1, cellFormatValueDecimal.NumberOfDecimal);
        Assert.AreEqual(ExcelCellValueNegativeOption.Default, cellFormatValueDecimal.NegativeOption);

        res = excelApi.ExcelCellValueApi.GetCellValueAsDouble(sheet, "B5", out cellVal);
        Assert.IsTrue(res);
        Assert.AreEqual(5.6, cellVal);

        //--B7: 123 - decimal - neg, red, no sign, format: "0.00;[Red]0.00"
        res = excelApi.ExcelCellValueApi.SetCellValueDecimal(sheet, "B7", 2, false, ExcelCellValueNegativeOption.RedWithoutSign, -123);
        Assert.IsTrue(res);

        // check
        cellFormat = excelApi.ExcelCellValueApi.GetCellFormat(sheet, "B7");
        Assert.AreEqual(ExcelCellFormatValueCode.Decimal, cellFormat.FormatValue.Code);
        Assert.AreEqual("0.00;[Red]0.00", cellFormat.FormatValue.StringFormat);
        cellFormatValueDecimal = cellFormat.GetFormatValueAsDecimal();
        Assert.IsNotNull(cellFormatValueDecimal);
        Assert.AreEqual(ExcelCellDecimalCode.DecimalN, cellFormatValueDecimal.DecimalCode);
        Assert.AreEqual(2, cellFormatValueDecimal.NumberOfDecimal);
        Assert.AreEqual(ExcelCellValueNegativeOption.RedWithoutSign, cellFormatValueDecimal.NegativeOption);

        res = excelApi.ExcelCellValueApi.GetCellValueAsDouble(sheet, "B7", out cellVal);
        Assert.IsTrue(res);
        Assert.AreEqual(-123, cellVal);

        //--B9: -123 - decimal - neg, red, sign, format: "0.00_ ;[Red]\\-0.00\\ "
        res = excelApi.ExcelCellValueApi.SetCellValueDecimal(sheet, "B9", 2, false, ExcelCellValueNegativeOption.RedWithSign, -123);
        Assert.IsTrue(res);

        // check
        cellFormat = excelApi.ExcelCellValueApi.GetCellFormat(sheet, "B9");
        Assert.AreEqual(ExcelCellFormatValueCode.Decimal, cellFormat.FormatValue.Code);
        Assert.AreEqual("0.00_ ;[Red]\\-0.00\\ ", cellFormat.FormatValue.StringFormat);
        cellFormatValueDecimal = cellFormat.GetFormatValueAsDecimal();
        Assert.IsNotNull(cellFormatValueDecimal);
        Assert.AreEqual(ExcelCellDecimalCode.DecimalN, cellFormatValueDecimal.DecimalCode);
        Assert.AreEqual(2, cellFormatValueDecimal.NumberOfDecimal);
        Assert.AreEqual(ExcelCellValueNegativeOption.RedWithSign, cellFormatValueDecimal.NegativeOption);

        res = excelApi.ExcelCellValueApi.GetCellValueAsDouble(sheet, "B9", out cellVal);
        Assert.IsTrue(res);
        Assert.AreEqual(-123, cellVal);

        //--B11: 123 000,50 -decimal, 2 dec. thousand sep, format: null
        res = excelApi.ExcelCellValueApi.SetCellValueDecimal(sheet, "B11", 2, true, ExcelCellValueNegativeOption.Default, 123000.50);
        Assert.IsTrue(res);

        // check
        cellFormat = excelApi.ExcelCellValueApi.GetCellFormat(sheet, "B11");
        Assert.AreEqual(ExcelCellFormatValueCode.Decimal, cellFormat.FormatValue.Code);
        Assert.AreEqual(String.Empty, cellFormat.FormatValue.StringFormat);
        cellFormatValueDecimal = cellFormat.GetFormatValueAsDecimal();
        Assert.IsNotNull(cellFormatValueDecimal);
        Assert.AreEqual(ExcelCellDecimalCode.Decimal4BlankThousandSep, cellFormatValueDecimal.DecimalCode);
        Assert.AreEqual(2, cellFormatValueDecimal.NumberOfDecimal);
        Assert.AreEqual(ExcelCellValueNegativeOption.Default, cellFormatValueDecimal.NegativeOption);

        res = excelApi.ExcelCellValueApi.GetCellValueAsDouble(sheet, "B11", out cellVal);
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

        ExcelApi excelApi = new ExcelApi();
        ExcelWorkbook workbook;
        ExcelError error;
        bool res = excelApi.ExcelFileApi.OpenExcelFile(fileName, out workbook, out error);
        Assert.IsTrue(res);

        // get the first sheet
        var sheet = excelApi.ExcelSheetApi.GetSheet(workbook, 0);
        Assert.IsNotNull(sheet);

        //--B1: 22,56 - decimal, a built-in format
        res = excelApi.ExcelCellValueApi.SetCellValueDecimal(sheet, "B1", 2, false, ExcelCellValueNegativeOption.Default, 22.56);
        Assert.IsTrue(res);

        // check
        double cellVal;
        res = excelApi.ExcelCellValueApi.GetCellValueAsDouble(sheet, "B1", out cellVal);
        Assert.IsTrue(res);
        Assert.AreEqual(22.56, cellVal);


        //--B3: 63,456 - decimal - 3dec
        res = excelApi.ExcelCellValueApi.SetCellValueDecimal(sheet, "B3", 3, false, ExcelCellValueNegativeOption.Default, 63.456);
        Assert.IsTrue(res);

        // check
        var cellFormat = excelApi.ExcelCellValueApi.GetCellFormat(sheet, "B3");
        Assert.AreEqual(ExcelCellFormatValueCode.Decimal, cellFormat.FormatValue.Code);
        Assert.AreEqual("0.000", cellFormat.FormatValue.StringFormat);
        var cellFormatValueDecimal = cellFormat.GetFormatValueAsDecimal();
        Assert.IsNotNull(cellFormatValueDecimal);
        Assert.AreEqual(ExcelCellDecimalCode.DecimalN, cellFormatValueDecimal.DecimalCode);
        Assert.AreEqual(3, cellFormatValueDecimal.NumberOfDecimal);
        Assert.AreEqual(ExcelCellValueNegativeOption.Default, cellFormatValueDecimal.NegativeOption);

        res = excelApi.ExcelCellValueApi.GetCellValueAsDouble(sheet, "B3", out cellVal);
        Assert.IsTrue(res);
        Assert.AreEqual(63.456, cellVal);


        //--B5: 5,6 - decimal - 1dec
        res = excelApi.ExcelCellValueApi.SetCellValueDecimal(sheet, "B5", 1, false, ExcelCellValueNegativeOption.Default, 5.6);
        Assert.IsTrue(res);

        // check
        cellFormat = excelApi.ExcelCellValueApi.GetCellFormat(sheet, "B5");
        Assert.AreEqual(ExcelCellFormatValueCode.Decimal, cellFormat.FormatValue.Code);
        Assert.AreEqual("0.0", cellFormat.FormatValue.StringFormat);
        cellFormatValueDecimal = cellFormat.GetFormatValueAsDecimal();
        Assert.IsNotNull(cellFormatValueDecimal);
        Assert.AreEqual(ExcelCellDecimalCode.DecimalN, cellFormatValueDecimal.DecimalCode);
        Assert.AreEqual(1, cellFormatValueDecimal.NumberOfDecimal);
        Assert.AreEqual(ExcelCellValueNegativeOption.Default, cellFormatValueDecimal.NegativeOption);

        res = excelApi.ExcelCellValueApi.GetCellValueAsDouble(sheet, "B5", out cellVal);
        Assert.IsTrue(res);
        Assert.AreEqual(5.6, cellVal);

        //--B7: 123 - decimal - neg, red, no sign, format: "0.00;[Red]0.00"
        res = excelApi.ExcelCellValueApi.SetCellValueDecimal(sheet, "B7", 2, false, ExcelCellValueNegativeOption.RedWithoutSign, -123);
        Assert.IsTrue(res);

        // check
        cellFormat = excelApi.ExcelCellValueApi.GetCellFormat(sheet, "B7");
        Assert.AreEqual(ExcelCellFormatValueCode.Decimal, cellFormat.FormatValue.Code);
        Assert.AreEqual("0.00;[Red]0.00", cellFormat.FormatValue.StringFormat);
        cellFormatValueDecimal = cellFormat.GetFormatValueAsDecimal();
        Assert.IsNotNull(cellFormatValueDecimal);
        Assert.AreEqual(ExcelCellDecimalCode.DecimalN, cellFormatValueDecimal.DecimalCode);
        Assert.AreEqual(2, cellFormatValueDecimal.NumberOfDecimal);
        Assert.AreEqual(ExcelCellValueNegativeOption.RedWithoutSign, cellFormatValueDecimal.NegativeOption);

        res = excelApi.ExcelCellValueApi.GetCellValueAsDouble(sheet, "B7", out cellVal);
        Assert.IsTrue(res);
        Assert.AreEqual(-123, cellVal);

        //--B9: -123 - decimal - neg, red, sign, format: "0.00_ ;[Red]\\-0.00\\ "
        res = excelApi.ExcelCellValueApi.SetCellValueDecimal(sheet, "B9", 2, false, ExcelCellValueNegativeOption.RedWithSign, -123);
        Assert.IsTrue(res);

        // check
        cellFormat = excelApi.ExcelCellValueApi.GetCellFormat(sheet, "B9");
        Assert.AreEqual(ExcelCellFormatValueCode.Decimal, cellFormat.FormatValue.Code);
        Assert.AreEqual("0.00_ ;[Red]\\-0.00\\ ", cellFormat.FormatValue.StringFormat);
        cellFormatValueDecimal = cellFormat.GetFormatValueAsDecimal();
        Assert.IsNotNull(cellFormatValueDecimal);
        Assert.AreEqual(ExcelCellDecimalCode.DecimalN, cellFormatValueDecimal.DecimalCode);
        Assert.AreEqual(2, cellFormatValueDecimal.NumberOfDecimal);
        Assert.AreEqual(ExcelCellValueNegativeOption.RedWithSign, cellFormatValueDecimal.NegativeOption);

        res = excelApi.ExcelCellValueApi.GetCellValueAsDouble(sheet, "B9", out cellVal);
        Assert.IsTrue(res);
        Assert.AreEqual(-123, cellVal);

        //--B11: 123 000,50 -decimal, 2 dec. thousand sep, format: null
        res = excelApi.ExcelCellValueApi.SetCellValueDecimal(sheet, "B11", 2, true, ExcelCellValueNegativeOption.Default, 123000.50);
        Assert.IsTrue(res);

        // check
        cellFormat = excelApi.ExcelCellValueApi.GetCellFormat(sheet, "B11");
        Assert.AreEqual(ExcelCellFormatValueCode.Decimal, cellFormat.FormatValue.Code);
        Assert.AreEqual(String.Empty, cellFormat.FormatValue.StringFormat);
        cellFormatValueDecimal = cellFormat.GetFormatValueAsDecimal();
        Assert.IsNotNull(cellFormatValueDecimal);
        Assert.AreEqual(ExcelCellDecimalCode.Decimal4BlankThousandSep, cellFormatValueDecimal.DecimalCode);
        Assert.AreEqual(2, cellFormatValueDecimal.NumberOfDecimal);
        Assert.AreEqual(ExcelCellValueNegativeOption.Default, cellFormatValueDecimal.NegativeOption);

        res = excelApi.ExcelCellValueApi.GetCellValueAsDouble(sheet, "B11", out cellVal);
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
    public void SetCellValuesSameFormatWithStyle()
    {
        string fileName = @"Files\SetCellValues\SetCellValuesDecimalSameFormatWithStyle.xlsx";

        ExcelApi excelApi = new ExcelApi();
        ExcelWorkbook workbook;
        ExcelError error;
        bool res = excelApi.ExcelFileApi.OpenExcelFile(fileName, out workbook, out error);
        Assert.IsTrue(res);

        // get the first sheet
        var sheet = excelApi.ExcelSheetApi.GetSheet(workbook, 0);
        Assert.IsNotNull(sheet);

        //--B1: 22,56 - decimal, a built-in format
        res = excelApi.ExcelCellValueApi.SetCellValueDecimal(sheet, "B1", 2, false, ExcelCellValueNegativeOption.Default, 22.56);
        Assert.IsTrue(res);

        // check
        double cellVal;
        res = excelApi.ExcelCellValueApi.GetCellValueAsDouble(sheet, "B1", out cellVal);
        Assert.IsTrue(res);
        Assert.AreEqual(22.56, cellVal);


        //--B3: 63,456 - decimal - 3dec
        res = excelApi.ExcelCellValueApi.SetCellValueDecimal(sheet, "B3", 3, false, ExcelCellValueNegativeOption.Default, 63.456);
        Assert.IsTrue(res);

        // check
        var cellFormat = excelApi.ExcelCellValueApi.GetCellFormat(sheet, "B3");
        Assert.AreEqual(ExcelCellFormatValueCode.Decimal, cellFormat.FormatValue.Code);
        Assert.AreEqual("0.000", cellFormat.FormatValue.StringFormat);
        var cellFormatValueDecimal = cellFormat.GetFormatValueAsDecimal();
        Assert.IsNotNull(cellFormatValueDecimal);
        Assert.AreEqual(ExcelCellDecimalCode.DecimalN, cellFormatValueDecimal.DecimalCode);
        Assert.AreEqual(3, cellFormatValueDecimal.NumberOfDecimal);
        Assert.AreEqual(ExcelCellValueNegativeOption.Default, cellFormatValueDecimal.NegativeOption);

        res = excelApi.ExcelCellValueApi.GetCellValueAsDouble(sheet, "B3", out cellVal);
        Assert.IsTrue(res);
        Assert.AreEqual(63.456, cellVal);


        //--B5: 5,6 - decimal - 1dec
        res = excelApi.ExcelCellValueApi.SetCellValueDecimal(sheet, "B5", 1, false, ExcelCellValueNegativeOption.Default, 5.6);
        Assert.IsTrue(res);

        // check
        cellFormat = excelApi.ExcelCellValueApi.GetCellFormat(sheet, "B5");
        Assert.AreEqual(ExcelCellFormatValueCode.Decimal, cellFormat.FormatValue.Code);
        Assert.AreEqual("0.0", cellFormat.FormatValue.StringFormat);
        cellFormatValueDecimal = cellFormat.GetFormatValueAsDecimal();
        Assert.IsNotNull(cellFormatValueDecimal);
        Assert.AreEqual(ExcelCellDecimalCode.DecimalN, cellFormatValueDecimal.DecimalCode);
        Assert.AreEqual(1, cellFormatValueDecimal.NumberOfDecimal);
        Assert.AreEqual(ExcelCellValueNegativeOption.Default, cellFormatValueDecimal.NegativeOption);

        res = excelApi.ExcelCellValueApi.GetCellValueAsDouble(sheet, "B5", out cellVal);
        Assert.IsTrue(res);
        Assert.AreEqual(5.6, cellVal);

        //--B7: 123 - decimal - neg, red, no sign, format: "0.00;[Red]0.00"
        res = excelApi.ExcelCellValueApi.SetCellValueDecimal(sheet, "B7", 2, false, ExcelCellValueNegativeOption.RedWithoutSign, -123);
        Assert.IsTrue(res);

        // check
        cellFormat = excelApi.ExcelCellValueApi.GetCellFormat(sheet, "B7");
        Assert.AreEqual(ExcelCellFormatValueCode.Decimal, cellFormat.FormatValue.Code);
        Assert.AreEqual("0.00;[Red]0.00", cellFormat.FormatValue.StringFormat);
        cellFormatValueDecimal = cellFormat.GetFormatValueAsDecimal();
        Assert.IsNotNull(cellFormatValueDecimal);
        Assert.AreEqual(ExcelCellDecimalCode.DecimalN, cellFormatValueDecimal.DecimalCode);
        Assert.AreEqual(2, cellFormatValueDecimal.NumberOfDecimal);
        Assert.AreEqual(ExcelCellValueNegativeOption.RedWithoutSign, cellFormatValueDecimal.NegativeOption);

        res = excelApi.ExcelCellValueApi.GetCellValueAsDouble(sheet, "B7", out cellVal);
        Assert.IsTrue(res);
        Assert.AreEqual(-123, cellVal);

        //--B9: -123 - decimal - neg, red, sign, format: "0.00_ ;[Red]\\-0.00\\ "
        res = excelApi.ExcelCellValueApi.SetCellValueDecimal(sheet, "B9", 2, false, ExcelCellValueNegativeOption.RedWithSign, -123);
        Assert.IsTrue(res);

        // check
        cellFormat = excelApi.ExcelCellValueApi.GetCellFormat(sheet, "B9");
        Assert.AreEqual(ExcelCellFormatValueCode.Decimal, cellFormat.FormatValue.Code);
        Assert.AreEqual("0.00_ ;[Red]\\-0.00\\ ", cellFormat.FormatValue.StringFormat);
        cellFormatValueDecimal = cellFormat.GetFormatValueAsDecimal();
        Assert.IsNotNull(cellFormatValueDecimal);
        Assert.AreEqual(ExcelCellDecimalCode.DecimalN, cellFormatValueDecimal.DecimalCode);
        Assert.AreEqual(2, cellFormatValueDecimal.NumberOfDecimal);
        Assert.AreEqual(ExcelCellValueNegativeOption.RedWithSign, cellFormatValueDecimal.NegativeOption);

        res = excelApi.ExcelCellValueApi.GetCellValueAsDouble(sheet, "B9", out cellVal);
        Assert.IsTrue(res);
        Assert.AreEqual(-123, cellVal);

        //--B11: 123 000,50 -decimal, 2 dec. thousand sep, format: null
        res = excelApi.ExcelCellValueApi.SetCellValueDecimal(sheet, "B11", 2, true, ExcelCellValueNegativeOption.Default, 123000.50);
        Assert.IsTrue(res);

        // check
        cellFormat = excelApi.ExcelCellValueApi.GetCellFormat(sheet, "B11");
        Assert.AreEqual(ExcelCellFormatValueCode.Decimal, cellFormat.FormatValue.Code);
        Assert.AreEqual(String.Empty, cellFormat.FormatValue.StringFormat);
        cellFormatValueDecimal = cellFormat.GetFormatValueAsDecimal();
        Assert.IsNotNull(cellFormatValueDecimal);
        Assert.AreEqual(ExcelCellDecimalCode.Decimal4BlankThousandSep, cellFormatValueDecimal.DecimalCode);
        Assert.AreEqual(2, cellFormatValueDecimal.NumberOfDecimal);
        Assert.AreEqual(ExcelCellValueNegativeOption.Default, cellFormatValueDecimal.NegativeOption);

        res = excelApi.ExcelCellValueApi.GetCellValueAsDouble(sheet, "B11", out cellVal);
        Assert.IsTrue(res);
        Assert.AreEqual(123000.50, cellVal);

        //--close the file
        res = excelApi.ExcelFileApi.CloseExcelFile(workbook, out error);
        Assert.IsTrue(res);
    }


    [TestMethod]
    public void SetCellValuesOtherFormatNoStyle()
    {
        string fileName = @"Files\SetCellValues\SetCellValuesDecimalOtherFormatNoStyle.xlsx";

        ExcelApi excelApi = new ExcelApi();
        ExcelWorkbook workbook;
        ExcelError error;
        bool res = excelApi.ExcelFileApi.OpenExcelFile(fileName, out workbook, out error);
        Assert.IsTrue(res);

        // get the first sheet
        var sheet = excelApi.ExcelSheetApi.GetSheet(workbook, 0);
        Assert.IsNotNull(sheet);

        //--B1: 22,56 - decimal, a built-in format
        res = excelApi.ExcelCellValueApi.SetCellValueDecimal(sheet, "B1", 2, false, ExcelCellValueNegativeOption.Default, 22.56);
        Assert.IsTrue(res);

        // check
        double cellVal;
        res = excelApi.ExcelCellValueApi.GetCellValueAsDouble(sheet, "B1", out cellVal);
        Assert.IsTrue(res);
        Assert.AreEqual(22.56, cellVal);


        //--B3: 63,456 - decimal - 3dec
        res = excelApi.ExcelCellValueApi.SetCellValueDecimal(sheet, "B3", 3, false, ExcelCellValueNegativeOption.Default, 63.456);
        Assert.IsTrue(res);

        // check
        var cellFormat = excelApi.ExcelCellValueApi.GetCellFormat(sheet, "B3");
        Assert.AreEqual(ExcelCellFormatValueCode.Decimal, cellFormat.FormatValue.Code);
        Assert.AreEqual("0.000", cellFormat.FormatValue.StringFormat);
        var cellFormatValueDecimal = cellFormat.GetFormatValueAsDecimal();
        Assert.IsNotNull(cellFormatValueDecimal);
        Assert.AreEqual(ExcelCellDecimalCode.DecimalN, cellFormatValueDecimal.DecimalCode);
        Assert.AreEqual(3, cellFormatValueDecimal.NumberOfDecimal);
        Assert.AreEqual(ExcelCellValueNegativeOption.Default, cellFormatValueDecimal.NegativeOption);

        res = excelApi.ExcelCellValueApi.GetCellValueAsDouble(sheet, "B3", out cellVal);
        Assert.IsTrue(res);
        Assert.AreEqual(63.456, cellVal);


        //--B5: 5,6 - decimal - 1dec
        res = excelApi.ExcelCellValueApi.SetCellValueDecimal(sheet, "B5", 1, false, ExcelCellValueNegativeOption.Default, 5.6);
        Assert.IsTrue(res);

        // check
        cellFormat = excelApi.ExcelCellValueApi.GetCellFormat(sheet, "B5");
        Assert.AreEqual(ExcelCellFormatValueCode.Decimal, cellFormat.FormatValue.Code);
        Assert.AreEqual("0.0", cellFormat.FormatValue.StringFormat);
        cellFormatValueDecimal = cellFormat.GetFormatValueAsDecimal();
        Assert.IsNotNull(cellFormatValueDecimal);
        Assert.AreEqual(ExcelCellDecimalCode.DecimalN, cellFormatValueDecimal.DecimalCode);
        Assert.AreEqual(1, cellFormatValueDecimal.NumberOfDecimal);

        res = excelApi.ExcelCellValueApi.GetCellValueAsDouble(sheet, "B5", out cellVal);
        Assert.IsTrue(res);
        Assert.AreEqual(5.6, cellVal);

        //--B7: 123 - decimal - neg, red, no sign, format: "0.00;[Red]0.00"
        res = excelApi.ExcelCellValueApi.SetCellValueDecimal(sheet, "B7", 2, false, ExcelCellValueNegativeOption.RedWithoutSign, -123);
        Assert.IsTrue(res);

        // check
        cellFormat = excelApi.ExcelCellValueApi.GetCellFormat(sheet, "B7");
        Assert.AreEqual(ExcelCellFormatValueCode.Decimal, cellFormat.FormatValue.Code);
        Assert.AreEqual("0.00;[Red]0.00", cellFormat.FormatValue.StringFormat);
        cellFormatValueDecimal = cellFormat.GetFormatValueAsDecimal();
        Assert.IsNotNull(cellFormatValueDecimal);
        Assert.AreEqual(ExcelCellDecimalCode.DecimalN, cellFormatValueDecimal.DecimalCode);
        Assert.AreEqual(2, cellFormatValueDecimal.NumberOfDecimal);
        Assert.AreEqual(ExcelCellValueNegativeOption.RedWithoutSign, cellFormatValueDecimal.NegativeOption);

        res = excelApi.ExcelCellValueApi.GetCellValueAsDouble(sheet, "B7", out cellVal);
        Assert.IsTrue(res);
        Assert.AreEqual(-123, cellVal);

        //--B9: -123 - decimal - neg, red, sign, format: "0.00_ ;[Red]\\-0.00\\ "
        res = excelApi.ExcelCellValueApi.SetCellValueDecimal(sheet, "B9", 2, false, ExcelCellValueNegativeOption.RedWithSign, -123);
        Assert.IsTrue(res);

        // check
        cellFormat = excelApi.ExcelCellValueApi.GetCellFormat(sheet, "B9");
        Assert.AreEqual(ExcelCellFormatValueCode.Decimal, cellFormat.FormatValue.Code);
        Assert.AreEqual("0.00_ ;[Red]\\-0.00\\ ", cellFormat.FormatValue.StringFormat);
        cellFormatValueDecimal = cellFormat.GetFormatValueAsDecimal();
        Assert.IsNotNull(cellFormatValueDecimal);
        Assert.AreEqual(ExcelCellDecimalCode.DecimalN, cellFormatValueDecimal.DecimalCode);
        Assert.AreEqual(2, cellFormatValueDecimal.NumberOfDecimal);
        Assert.AreEqual(ExcelCellValueNegativeOption.RedWithSign, cellFormatValueDecimal.NegativeOption);

        res = excelApi.ExcelCellValueApi.GetCellValueAsDouble(sheet, "B9", out cellVal);
        Assert.IsTrue(res);
        Assert.AreEqual(-123, cellVal);

        //--B11: 123 000,50 -decimal, 2 dec. thousand sep, format: null
        res = excelApi.ExcelCellValueApi.SetCellValueDecimal(sheet, "B11", 2, true, ExcelCellValueNegativeOption.Default, 123000.50);
        Assert.IsTrue(res);

        // check
        cellFormat = excelApi.ExcelCellValueApi.GetCellFormat(sheet, "B11");
        Assert.AreEqual(ExcelCellFormatValueCode.Decimal, cellFormat.FormatValue.Code);
        Assert.AreEqual(String.Empty, cellFormat.FormatValue.StringFormat);
        cellFormatValueDecimal = cellFormat.GetFormatValueAsDecimal();
        Assert.IsNotNull(cellFormatValueDecimal);
        Assert.AreEqual(ExcelCellDecimalCode.Decimal4BlankThousandSep, cellFormatValueDecimal.DecimalCode);
        Assert.AreEqual(2, cellFormatValueDecimal.NumberOfDecimal);
        Assert.AreEqual(ExcelCellValueNegativeOption.Default, cellFormatValueDecimal.NegativeOption);

        res = excelApi.ExcelCellValueApi.GetCellValueAsDouble(sheet, "B11", out cellVal);
        Assert.IsTrue(res);
        Assert.AreEqual(123000.50, cellVal);

        //--close the file
        res = excelApi.ExcelFileApi.CloseExcelFile(workbook, out error);
        Assert.IsTrue(res);
    }

    [TestMethod]
    public void SetCellValuesOtherFormatWithStyle()
    {
        string fileName = @"Files\SetCellValues\SetCellValuesDecimalOtherFormatWithStyle.xlsx";

        ExcelApi excelApi = new ExcelApi();
        ExcelWorkbook workbook;
        ExcelError error;
        bool res = excelApi.ExcelFileApi.OpenExcelFile(fileName, out workbook, out error);
        Assert.IsTrue(res);

        // get the first sheet
        var sheet = excelApi.ExcelSheetApi.GetSheet(workbook, 0);
        Assert.IsNotNull(sheet);

        //--B1: 22,56 - decimal, a built-in format
        res = excelApi.ExcelCellValueApi.SetCellValueDecimal(sheet, "B1", 2, false, ExcelCellValueNegativeOption.Default, 22.56);
        Assert.IsTrue(res);

        // check
        double cellVal;
        res = excelApi.ExcelCellValueApi.GetCellValueAsDouble(sheet, "B1", out cellVal);
        Assert.IsTrue(res);
        Assert.AreEqual(22.56, cellVal);


        //--B3: 63,456 - decimal - 3dec
        res = excelApi.ExcelCellValueApi.SetCellValueDecimal(sheet, "B3", 3, false, ExcelCellValueNegativeOption.Default, 63.456);
        Assert.IsTrue(res);

        // check
        var cellFormat = excelApi.ExcelCellValueApi.GetCellFormat(sheet, "B3");
        Assert.AreEqual(ExcelCellFormatValueCode.Decimal, cellFormat.FormatValue.Code);
        Assert.AreEqual("0.000", cellFormat.FormatValue.StringFormat);
        var cellFormatValueDecimal = cellFormat.GetFormatValueAsDecimal();
        Assert.IsNotNull(cellFormatValueDecimal);
        Assert.AreEqual(ExcelCellDecimalCode.DecimalN, cellFormatValueDecimal.DecimalCode);
        Assert.AreEqual(3, cellFormatValueDecimal.NumberOfDecimal);
        Assert.AreEqual(ExcelCellValueNegativeOption.Default, cellFormatValueDecimal.NegativeOption);

        res = excelApi.ExcelCellValueApi.GetCellValueAsDouble(sheet, "B3", out cellVal);
        Assert.IsTrue(res);
        Assert.AreEqual(63.456, cellVal);


        //--B5: 5,6 - decimal - 1dec
        res = excelApi.ExcelCellValueApi.SetCellValueDecimal(sheet, "B5", 1, false, ExcelCellValueNegativeOption.Default, 5.6);
        Assert.IsTrue(res);

        // check
        cellFormat = excelApi.ExcelCellValueApi.GetCellFormat(sheet, "B5");
        Assert.AreEqual(ExcelCellFormatValueCode.Decimal, cellFormat.FormatValue.Code);
        Assert.AreEqual("0.0", cellFormat.FormatValue.StringFormat);
        cellFormatValueDecimal = cellFormat.GetFormatValueAsDecimal();
        Assert.IsNotNull(cellFormatValueDecimal);
        Assert.AreEqual(ExcelCellDecimalCode.DecimalN, cellFormatValueDecimal.DecimalCode);
        Assert.AreEqual(1, cellFormatValueDecimal.NumberOfDecimal);
        Assert.AreEqual(ExcelCellValueNegativeOption.Default, cellFormatValueDecimal.NegativeOption);

        res = excelApi.ExcelCellValueApi.GetCellValueAsDouble(sheet, "B5", out cellVal);
        Assert.IsTrue(res);
        Assert.AreEqual(5.6, cellVal);

        //--B7: 123 - decimal - neg, red, no sign, format: "0.00;[Red]0.00"
        res = excelApi.ExcelCellValueApi.SetCellValueDecimal(sheet, "B7", 2, false, ExcelCellValueNegativeOption.RedWithoutSign, -123);
        Assert.IsTrue(res);

        // check
        cellFormat = excelApi.ExcelCellValueApi.GetCellFormat(sheet, "B7");
        Assert.AreEqual(ExcelCellFormatValueCode.Decimal, cellFormat.FormatValue.Code);
        Assert.AreEqual("0.00;[Red]0.00", cellFormat.FormatValue.StringFormat);
        cellFormatValueDecimal = cellFormat.GetFormatValueAsDecimal();
        Assert.IsNotNull(cellFormatValueDecimal);
        Assert.AreEqual(ExcelCellDecimalCode.DecimalN, cellFormatValueDecimal.DecimalCode);
        Assert.AreEqual(2, cellFormatValueDecimal.NumberOfDecimal);
        Assert.AreEqual(ExcelCellValueNegativeOption.RedWithoutSign, cellFormatValueDecimal.NegativeOption);

        res = excelApi.ExcelCellValueApi.GetCellValueAsDouble(sheet, "B7", out cellVal);
        Assert.IsTrue(res);
        Assert.AreEqual(-123, cellVal);

        //--B9: -123 - decimal - neg, red, sign, format: "0.00_ ;[Red]\\-0.00\\ "
        res = excelApi.ExcelCellValueApi.SetCellValueDecimal(sheet, "B9", 2, false, ExcelCellValueNegativeOption.RedWithSign, -123);
        Assert.IsTrue(res);

        // check
        cellFormat = excelApi.ExcelCellValueApi.GetCellFormat(sheet, "B9");
        Assert.AreEqual(ExcelCellFormatValueCode.Decimal, cellFormat.FormatValue.Code);
        Assert.AreEqual("0.00_ ;[Red]\\-0.00\\ ", cellFormat.FormatValue.StringFormat);
        cellFormatValueDecimal = cellFormat.GetFormatValueAsDecimal();
        Assert.IsNotNull(cellFormatValueDecimal);
        Assert.AreEqual(ExcelCellDecimalCode.DecimalN, cellFormatValueDecimal.DecimalCode);
        Assert.AreEqual(2, cellFormatValueDecimal.NumberOfDecimal);
        Assert.AreEqual(ExcelCellValueNegativeOption.RedWithSign, cellFormatValueDecimal.NegativeOption);

        res = excelApi.ExcelCellValueApi.GetCellValueAsDouble(sheet, "B9", out cellVal);
        Assert.IsTrue(res);
        Assert.AreEqual(-123, cellVal);

        //--B11: 123 000,50 -decimal, 2 dec. thousand sep, format: null
        res = excelApi.ExcelCellValueApi.SetCellValueDecimal(sheet, "B11", 2, true, ExcelCellValueNegativeOption.Default, 123000.50);
        Assert.IsTrue(res);

        // check
        cellFormat = excelApi.ExcelCellValueApi.GetCellFormat(sheet, "B11");
        Assert.AreEqual(ExcelCellFormatValueCode.Decimal, cellFormat.FormatValue.Code);
        Assert.AreEqual(String.Empty, cellFormat.FormatValue.StringFormat);
        cellFormatValueDecimal = cellFormat.GetFormatValueAsDecimal();
        Assert.IsNotNull(cellFormatValueDecimal);
        Assert.AreEqual(ExcelCellDecimalCode.Decimal4BlankThousandSep, cellFormatValueDecimal.DecimalCode);
        Assert.AreEqual(2, cellFormatValueDecimal.NumberOfDecimal);
        Assert.AreEqual(ExcelCellValueNegativeOption.Default, cellFormatValueDecimal.NegativeOption);

        res = excelApi.ExcelCellValueApi.GetCellValueAsDouble(sheet, "B11", out cellVal);
        Assert.IsTrue(res);
        Assert.AreEqual(123000.50, cellVal);

        //--close the file
        res = excelApi.ExcelFileApi.CloseExcelFile(workbook, out error);
        Assert.IsTrue(res);
    }

}
