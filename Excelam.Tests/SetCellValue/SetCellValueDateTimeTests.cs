using Excelam.System;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Excelam.Tests.SetCellValue;

/// <summary>
/// Test the datetime type.
/// </summary>
[TestClass]
public class SetCellValueDateTimeTests
{
    /// <summary>
    /// Start from aa new excel file.
    /// set cell values, test all formats.
    /// </summary>
    [TestMethod]
    public void SetCellValuesDateTimeNew()
    {
        string fileName = @"Files\SetCellValues\SetCellValuesDatetimeNew.xlsx";

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

        //--B:1 set 12/06/2022 - DateShort14
        res = excelApi.ExcelCellValueApi.SetCellValueDateShort(sheet, "B1", new DateTime(2022,06,12));
        Assert.IsTrue(res);

        // check format
        ExcelCellFormat cellFormat = excelApi.ExcelCellValueApi.GetCellFormat(sheet, "B1");
        ExcelCellFormatValueDateTime formatValue = cellFormat.GetFormatValueAsDateTime();
        Assert.IsNotNull(formatValue);
        Assert.AreEqual(ExcelCellDateTimeCode.DateShort14, formatValue.DateTimeCode);
        // check value
        DateTime cellVal;
        res = excelApi.ExcelCellValueApi.GetCellValueAsDateTime(sheet, "B1", out cellVal);
        Assert.IsTrue(res);
        Assert.AreEqual(2022, cellVal.Year);
        Assert.AreEqual(6, cellVal.Month);
        Assert.AreEqual(12, cellVal.Day);

        //--B3: set 00:14:25 - Time21_hh_mm_ss
        res = excelApi.ExcelCellValueApi.SetCellValueDateTime(sheet, "B3", ExcelCellDateTimeCode.Time21_hh_mm_ss, new DateTime(2022, 06, 12,0,14,25));
        Assert.IsTrue(res);

        // check format
        cellFormat = excelApi.ExcelCellValueApi.GetCellFormat(sheet, "B3");
        formatValue = cellFormat.GetFormatValueAsDateTime();
        Assert.IsNotNull(formatValue);
        Assert.AreEqual(ExcelCellDateTimeCode.Time21_hh_mm_ss, formatValue.DateTimeCode);
        // check value
        res = excelApi.ExcelCellValueApi.GetCellValueAsDateTime(sheet, "B3", out cellVal);
        Assert.IsTrue(res);
        Assert.AreEqual(0, cellVal.Hour);
        Assert.AreEqual(14, cellVal.Minute);
        Assert.AreEqual(25, cellVal.Second);

        //--B5: set 12:24:52 - Time
        res = excelApi.ExcelCellValueApi.SetCellValueDateTime(sheet, "B5", ExcelCellDateTimeCode.Time, new DateTime(2022, 06, 12, 12, 24, 52));
        Assert.IsTrue(res);

        // check format
        cellFormat = excelApi.ExcelCellValueApi.GetCellFormat(sheet, "B5");
        formatValue = cellFormat.GetFormatValueAsDateTime();
        Assert.IsNotNull(formatValue);
        Assert.AreEqual(ExcelCellDateTimeCode.Time, formatValue.DateTimeCode);
        // check value
        res = excelApi.ExcelCellValueApi.GetCellValueAsDateTime(sheet, "B5", out cellVal);
        Assert.IsTrue(res);
        Assert.AreEqual(12, cellVal.Hour);
        Assert.AreEqual(24, cellVal.Minute);
        Assert.AreEqual(52, cellVal.Second);

        //---TODO: Autres cas!!

        //--close the file
        res = excelApi.ExcelFileApi.CloseExcelFile(workbook, out error);
        Assert.IsTrue(res);
    }

    //    public void SetCellValuesDateTimeEmpty()

}
