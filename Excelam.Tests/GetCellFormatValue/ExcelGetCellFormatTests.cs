using Excelam.System;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Excelam.Tests.GetCellFormatValue;

/// <summary>
/// Test ExcelCellValueApi.GetCellFormat cases.
/// </summary>
[TestClass]
public class GetCellFormatTests
{
    /// <summary>
    /// Get cel format values, general and text.
    /// </summary>
    [TestMethod]
    public void GetCellFormatValuesGeneralText()
    {
        ExcelApi excelApi = new ExcelApi();

        string fileName = @"Files\GetCellFormat\GetCellFormatValuesGeneralText.xlsx";
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

        //--B3: bonjour - standard/general
        ExcelCellFormat cellFormatB3 = excelApi.ExcelCellValueApi.GetCellFormat(sheet, "B3");
        Assert.AreEqual(ExcelCellFormatValueCode.General, cellFormatB3.FormatValue.Code);
        Assert.IsInstanceOfType(cellFormatB3.FormatValue, typeof(ExcelCellFormatValueGeneral));

        //--B3: bonjour - standard/general  - col, row
        ExcelCellFormat cellFormatB3b = excelApi.ExcelCellValueApi.GetCellFormat(sheet, 2, 3);
        Assert.AreEqual(ExcelCellFormatValueCode.General, cellFormatB3b.FormatValue.Code);
        Assert.IsInstanceOfType(cellFormatB3b.FormatValue, typeof(ExcelCellFormatValueGeneral));

        //--B5: text
        ExcelCellFormat cellFormatB5 = excelApi.ExcelCellValueApi.GetCellFormat(sheet, "B5");
        Assert.AreEqual(ExcelCellFormatValueCode.Text, cellFormatB5.FormatValue.Code);
        Assert.IsInstanceOfType(cellFormatB5.FormatValue, typeof(ExcelCellFormatValueText));

        //--close the file
        res = excelApi.ExcelFileApi.CloseExcelFile(workbook, out error);
        Assert.IsTrue(res);
    }

    /// <summary>
    /// Get cell format values : decimal type.
    /// </summary>
    [TestMethod]
    public void GetCellFormatValuesDecimal()
    {
        ExcelApi excelApi = new ExcelApi();

        string fileName = @"Files\GetCellFormat\GetCellFormatValuesDecimal.xlsx";
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
        ExcelCellFormat cellFormat = excelApi.ExcelCellValueApi.GetCellFormat(sheet, "B1");
        Assert.AreEqual(1, cellFormat.FormatValue.NumberFormatId);
        Assert.AreEqual(ExcelCellFormatValueCode.Number, cellFormat.FormatValue.Code);
        ExcelCellFormatValueNumber cellFormatValueNumber = cellFormat.GetFormatValueAsNumber();
        Assert.IsInstanceOfType(cellFormat.FormatValue, typeof(ExcelCellFormatValueNumber));
        Assert.IsNotNull(cellFormatValueNumber);

        //--B3: 22,56 - decimal, a built-in format
        cellFormat = excelApi.ExcelCellValueApi.GetCellFormat(sheet, "B3");
        Assert.AreEqual(2, cellFormat.FormatValue.NumberFormatId);
        Assert.AreEqual(ExcelCellFormatValueCode.Decimal, cellFormat.FormatValue.Code);
        Assert.IsNull(cellFormat.FormatValue.ExcelNumberingFormat);
        ExcelCellFormatValueDecimal cellFormatValueDecimal = cellFormat.GetFormatValueAsDecimal();
        Assert.IsNotNull(cellFormatValueDecimal);
        Assert.AreEqual(ExcelCellDecimalCode.Decimal, cellFormatValueDecimal.SubCode);
        Assert.AreEqual(2, cellFormatValueDecimal.NumberOfDecimal);

        //--B5: 63,456 - decimal - 3dec
        cellFormat = excelApi.ExcelCellValueApi.GetCellFormat(sheet, "B5");
        Assert.IsTrue(cellFormat.FormatValue.NumberFormatId >163);
        Assert.AreEqual(ExcelCellFormatValueCode.Decimal, cellFormat.FormatValue.Code);
        Assert.AreEqual("0.000", cellFormat.FormatValue.StringFormat);
        cellFormatValueDecimal = cellFormat.GetFormatValueAsDecimal();
        Assert.IsNotNull(cellFormatValueDecimal);
        Assert.AreEqual(ExcelCellDecimalCode.Decimal, cellFormatValueDecimal.SubCode);
        Assert.AreEqual(3, cellFormatValueDecimal.NumberOfDecimal);


        //--B7: 5,6 - decimal - 1dec
        cellFormat = excelApi.ExcelCellValueApi.GetCellFormat(sheet, "B7");
        Assert.IsTrue(cellFormat.FormatValue.NumberFormatId > 163);
        Assert.AreEqual(ExcelCellFormatValueCode.Decimal, cellFormat.FormatValue.Code);
        Assert.AreEqual("0.0", cellFormat.FormatValue.StringFormat);
        cellFormatValueDecimal = cellFormat.GetFormatValueAsDecimal();
        Assert.IsNotNull(cellFormatValueDecimal);
        Assert.AreEqual(ExcelCellDecimalCode.Decimal, cellFormatValueDecimal.SubCode);
        Assert.AreEqual(1, cellFormatValueDecimal.NumberOfDecimal);

        //--B9: 123 - decimal - neg, red, no sign, format: "0.00;[Red]0.00"
        cellFormat = excelApi.ExcelCellValueApi.GetCellFormat(sheet, "B9");
        Assert.IsTrue(cellFormat.FormatValue.NumberFormatId > 163);
        Assert.AreEqual(ExcelCellFormatValueCode.Decimal, cellFormat.FormatValue.Code);
        Assert.AreEqual("0.00;[Red]0.00", cellFormat.FormatValue.StringFormat);
        cellFormatValueDecimal = cellFormat.GetFormatValueAsDecimal();
        Assert.IsNotNull(cellFormatValueDecimal);
        Assert.AreEqual(ExcelCellDecimalCode.DecimalNegRedNoSign, cellFormatValueDecimal.SubCode);
        Assert.AreEqual(2, cellFormatValueDecimal.NumberOfDecimal);

        //--B11: -123 - decimal - neg, red, sign, format: "0.00_ ;[Red]\\-0.00\\ "
        cellFormat = excelApi.ExcelCellValueApi.GetCellFormat(sheet, "B11");
        Assert.IsTrue(cellFormat.FormatValue.NumberFormatId > 163);
        Assert.AreEqual(ExcelCellFormatValueCode.Decimal, cellFormat.FormatValue.Code);
        Assert.AreEqual("0.00_ ;[Red]\\-0.00\\ ", cellFormat.FormatValue.StringFormat);
        cellFormatValueDecimal = cellFormat.GetFormatValueAsDecimal();
        Assert.IsNotNull(cellFormatValueDecimal);
        Assert.AreEqual(ExcelCellDecimalCode.DecimalNegRed, cellFormatValueDecimal.SubCode);
        Assert.AreEqual(2, cellFormatValueDecimal.NumberOfDecimal);


        //--B13: 123 000,50 -decimal, 2 dec. thousand sep, format: null
        cellFormat = excelApi.ExcelCellValueApi.GetCellFormat(sheet, "B13");
        Assert.AreEqual(4, cellFormat.FormatValue.NumberFormatId);
        Assert.AreEqual(ExcelCellFormatValueCode.Decimal, cellFormat.FormatValue.Code);
        Assert.AreEqual(string.Empty,cellFormat.FormatValue.StringFormat);
        cellFormatValueDecimal = cellFormat.GetFormatValueAsDecimal();
        Assert.IsNotNull(cellFormatValueDecimal);
        Assert.AreEqual(ExcelCellDecimalCode.DecimalBlankThousandSep, cellFormatValueDecimal.SubCode);
        Assert.AreEqual(2, cellFormatValueDecimal.NumberOfDecimal);

        //--close the file
        res = excelApi.ExcelFileApi.CloseExcelFile(workbook, out error);
        Assert.IsTrue(res);

    }

    /// <summary>
    /// Date and time cases.
    /// </summary>
    [TestMethod]
    public void GetCellFormatValuesDateTime()
    {
        ExcelApi excelApi = new ExcelApi();

        string fileName = @"Files\GetCellFormat\GetCellFormatValuesDateTime.xlsx";
        ExcelWorkbook workbook;
        ExcelError error;

        bool res = excelApi.ExcelFileApi.OpenExcelFile(fileName, out workbook, out error);
        Assert.IsTrue(res);
        Assert.IsNotNull(workbook);
        Assert.IsNull(error);

        // get the first sheet
        var sheet = excelApi.ExcelSheetApi.GetSheet(workbook, 0);
        Assert.IsNotNull(sheet);

        //--B1: 15/02/2021 - DateShort
        ExcelCellFormat cellFormat = excelApi.ExcelCellValueApi.GetCellFormat(sheet, "B1");
        Assert.AreEqual(ExcelCellFormatValueCode.DateTime, cellFormat.FormatValue.Code);
        Assert.AreEqual(ExcelCellDateTimeCode.DateShort, (cellFormat.FormatValue as ExcelCellFormatValueDateTime).DateTimeCode);

        //--todo: other cases: date large, time,...

        //--close the file
        res = excelApi.ExcelFileApi.CloseExcelFile(workbook, out error);
        Assert.IsTrue(res);
    }

    /// <summary>
    /// Date and time cases.
    /// </summary>
    [TestMethod]
    public void GetCellFormatValuesDateTimeEnglishUS()
    {
        ExcelApi excelApi = new ExcelApi();

        string fileName = @"Files\GetCellFormat\GetCellFormatValuesDateTimeEnglishUS.xlsx";
        ExcelWorkbook workbook;
        ExcelError error;

        bool res = excelApi.ExcelFileApi.OpenExcelFile(fileName, out workbook, out error);
        Assert.IsTrue(res);
        Assert.IsNotNull(workbook);
        Assert.IsNull(error);

        // get the first sheet
        var sheet = excelApi.ExcelSheetApi.GetSheet(workbook, 0);
        Assert.IsNotNull(sheet);

        //--C2: 2022-02-14 -DateShort
        ExcelCellFormat cellFormat = excelApi.ExcelCellValueApi.GetCellFormat(sheet, "C2");
        Assert.AreEqual(ExcelCellFormatValueCode.DateTime, cellFormat.FormatValue.Code);
        Assert.AreEqual(ExcelCellDateTimeCode.DateShort, (cellFormat.FormatValue as ExcelCellFormatValueDateTime).DateTimeCode);

        //--C4: May 15, 1980 -DateLarge
        cellFormat = excelApi.ExcelCellValueApi.GetCellFormat(sheet, "C4");
        Assert.AreEqual(ExcelCellFormatValueCode.DateTime, cellFormat.FormatValue.Code);
        Assert.AreEqual(ExcelCellDateTimeCode.DateShort, (cellFormat.FormatValue as ExcelCellFormatValueDateTime).DateTimeCode);

        //--todo: other cases: date large, time,...

        //--close the file
        res = excelApi.ExcelFileApi.CloseExcelFile(workbook, out error);
        Assert.IsTrue(res);
    }

    [TestMethod]
    public void GetCellFormatValuesAccounting()
    {
        ExcelApi excelApi = new ExcelApi();

        string fileName = @"Files\Cells\GetCellFormatValuesBuiltInAccounting.xlsx";
        ExcelWorkbook workbook;
        ExcelError error;

        bool res = excelApi.ExcelFileApi.OpenExcelFile(fileName, out workbook, out error);
        Assert.IsTrue(res);
        Assert.IsNotNull(workbook);
        Assert.IsNull(error);

        // get the first sheet
        var sheet = excelApi.ExcelSheetApi.GetSheet(workbook, 0);
        Assert.IsNotNull(sheet);

        //--B2: 12,34€ - accounting, 2 decimales

        //--B11: 45,21 €  - accounting
        //ExcelCellFormat cellFormatB11 = excelApi.ExcelCellValueApi.GetCellFormat(sheet, "B11");
        //Assert.AreEqual(ExcelCellFormatMainCode.Accounting, cellFormatB11.StructCode.MainCode);
        //Assert.AreEqual(ExcelCellCurrencyCode.Euro, cellFormatB11.StructCode.CurrencyCode);


        //--B4: 35,901 € - accounting, 3 decimales

        //--B6: 66,45 - accounting, 2 decimales, no currency symbo


        //--close the file
        res = excelApi.ExcelFileApi.CloseExcelFile(workbook, out error);
        Assert.IsTrue(res);


        Assert.Fail("todo:");
    }


    /// <summary>
    /// todo: rework, too many cases.
    /// </summary>
    //[TestMethod]
    //public void GetManyCellFormatValue()
    //{
    //    ExcelApi excelApi = new ExcelApi();

    //    string fileName = @"Files\Cells\GetManyCellFormatValues.xlsx";
    //    ExcelWorkbook workbook;
    //    ExcelError error;

    //    bool res = excelApi.ExcelFileApi.OpenExcelFile(fileName, out workbook, out error);
    //    Assert.IsTrue(res);
    //    Assert.IsNotNull(workbook);
    //    Assert.IsNull(error);

    //    // get the first sheet
    //    var sheet = excelApi.ExcelSheetApi.GetSheet(workbook, 0);
    //    Assert.IsNotNull(sheet);

    //    //--B1: null
    //    ExcelCellFormat cellFormatB1 = excelApi.ExcelCellValueApi.GetCellFormat(sheet, "B1");
    //    Assert.IsNull(cellFormatB1);

    //    //--B2: bonjour - standard/general
    //    ExcelCellFormat cellFormatB2 = excelApi.ExcelCellValueApi.GetCellFormat(sheet, "B2");
    //    Assert.AreEqual(ExcelCellFormatValueCode.General, cellFormatB2.FormatValue.Code);

    //    //--B2: bonjour - standard/general  - col, row
    //    ExcelCellFormat cellFormatB2b = excelApi.ExcelCellValueApi.GetCellFormat(sheet, 2, 2);

    //    //--B3: 12 - number
    //    ExcelCellFormat cellFormatB3 = excelApi.ExcelCellValueApi.GetCellFormat(sheet, "B3");
    //    Assert.AreEqual(ExcelCellFormatValueCode.Number, cellFormatB3.FormatValue.Code);

    //    //--B15: 15/02/2021 - Short date
    //    ExcelCellFormat cellFormatB15 = excelApi.ExcelCellValueApi.GetCellFormat(sheet, "B15");
    //    Assert.AreEqual(ExcelCellFormatValueCode.DateShort, cellFormatB15.FormatValue.Code);

    //    //--B16: vendredi 19 septembre 1969 - date large
    //    ExcelCellFormat cellFormatB16 = excelApi.ExcelCellValueApi.GetCellFormat(sheet, "B16");
    //    Assert.AreEqual(ExcelCellFormatValueCode.DateLarge, cellFormatB16.FormatValue.Code);

    //    //--B17: 13:45:00 - time
    //    ExcelCellFormat cellFormatB17 = excelApi.ExcelCellValueApi.GetCellFormat(sheet, "B17");
    //    Assert.AreEqual(ExcelCellFormatValueCode.Time, cellFormatB17.FormatValue.Code);

    //    //--B21: 45,21 €  - accounting
    //    ExcelCellFormat cellFormatB21 = excelApi.ExcelCellValueApi.GetCellFormat(sheet, "B21");
    //    Assert.AreEqual(ExcelCellFormatValueCode.Accounting, cellFormatB21.FormatValue.Code);
    //    Assert.AreEqual(ExcelCellCurrencyCode.Euro, cellFormatB21.StructCode.CurrencyCode);

    //    //--B22: 88,22 € - currency-euro
    //    ExcelCellFormat cellFormatB22 = excelApi.ExcelCellValueApi.GetCellFormat(sheet, "B22");
    //    Assert.AreEqual(ExcelCellFormatValueCode.Currency, cellFormatB22.FormatValue.Code);
    //    Assert.AreEqual(ExcelCellCurrencyCode.Euro, cellFormatB21.StructCode.CurrencyCode);

    //    //--B23: $91,25 - currency-dollarUS
    //    ExcelCellFormat cellFormatB23 = excelApi.ExcelCellValueApi.GetCellFormat(sheet, "B23");
    //    Assert.AreEqual(ExcelCellFormatValueCode.Currency, cellFormatB23.FormatValue.Code);
    //    Assert.AreEqual(ExcelCellCurrencyCode.UnitedStatesDollar, cellFormatB23.StructCode.CurrencyCode);


    //    //XXXXXXXXXXXX
    //    // TODO: tester les différents type euro.

    //    //--B25: € 1,2 - accounting - € euro
    //    //ExcelCellFormat cellFormatB25 = excelApi.ExcelCellValueApi.GetCellFormat(sheet, "B25");
    //    //Assert.AreEqual(ExcelCellFormatCode.Accounting, cellFormatB25.Code);
    //    //Assert.AreEqual(ExcelCellCurrencyCode.Euro, cellFormatB25.CurrencyCode);

    //    //--B26: € 2,3 - accounting - € Anglais Irlande
    //    //ExcelCellFormat cellFormatB26 = excelApi.ExcelCellValueApi.GetCellFormat(sheet, "B26");
    //    //Assert.AreEqual(ExcelCellFormatCode.Accounting, cellFormatB26.Code);
    //    //Assert.AreEqual(ExcelCellCurrencyCode.Euro, cellFormatB26.CurrencyCode);

    //    //--B27: € 4,12 - accounting $ anglais - canada
    //    //ExcelCellFormat cellFormatB27 = excelApi.ExcelCellValueApi.GetCellFormat(sheet, "B27");
    //    //Assert.AreEqual(ExcelCellFormatCode.Currency, cellFormatB27.Code);
    //    //Assert.AreEqual(ExcelCellCurrencyCode.CanadianDollar, cellFormatB27.CurrencyCode);



    //    //--E7: 45.60 - decimal - Formula
    //    ExcelCellFormat cellFormatE7 = excelApi.ExcelCellValueApi.GetCellFormat(sheet, "E7");
    //    Assert.AreEqual(ExcelCellFormatValueCode.Decimal, cellFormatE7.FormatValue.Code);
    //    Assert.IsTrue(cellFormatE7.IsFormula);

    //    //--todo: faire autres cas: built-in: fraction, percetange, scientific,...

    //    //--todo: faire autres cas speciaux: date longue, accounting, currency....


    //    //--close the file
    //    res = excelApi.ExcelFileApi.CloseExcelFile(workbook, out error);
    //    Assert.IsTrue(res);

    //}
}
