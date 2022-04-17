using Excelam.System;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Excelam.Tests;

/// <summary>
/// delete cell tests.
/// </summary>
[TestClass]
public class ExcelCellValueApiDelValueTests
{
    [TestMethod]
    public void CheckDeleteSomeCells()
    {
        string fileName = @"Files\Cells\DeleteCellWork.xlsx";
        ExcelWorkbook workbook;
        ExcelError error;

        // if file exists, remove it
        if (File.Exists(fileName))
            File.Delete(fileName);

        string fileNameTempl = @"Files\Cells\DeleteCell.xlsx";
        File.Copy(fileNameTempl, fileName);

        ExcelApi excelApi = new ExcelApi();
        bool res = excelApi.ExcelFileApi.OpenExcelFile(fileName, out workbook, out error);
        Assert.IsTrue(res);

        // get the first sheet
        var sheet = excelApi.ExcelSheetApi.GetSheet(workbook, 0);

        //--B1: null
        res= excelApi.ExcelCellValueApi.DeleteCell(sheet, "B1");
        Assert.IsFalse(res);

        //--B2: bonjour
        res = excelApi.ExcelCellValueApi.DeleteCell(sheet, "B2");
        Assert.IsTrue(res);

        string valB2=excelApi.ExcelCellValueApi.GetCellValueAsString(sheet, "B2");
        // no cell, removed
        Assert.IsNull(valB2);

        //--B3: 12
        res = excelApi.ExcelCellValueApi.DeleteCell(sheet, "B3");
        Assert.IsTrue(res);

        string valB3 = excelApi.ExcelCellValueApi.GetCellValueAsString(sheet, "B3");
        // no cell, removed
        Assert.IsNull(valB3);

        //remove a cell containing a formula
        // TODO:

        //--close the file
        res = excelApi.ExcelFileApi.CloseExcelFile(workbook, out error);
        Assert.IsTrue(res);

    }
}
