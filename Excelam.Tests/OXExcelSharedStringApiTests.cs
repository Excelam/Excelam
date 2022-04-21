using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using Excelam.OpenXmlLayer;
using Excelam.System;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Excelam.Tests;

[TestClass]
public class OXExcelSharedStringApiTests
{
    [TestMethod]
    public void CheckSharedStringCellValue()
    {
        string fileName = @"Files\Cells\ManyCellTypesSharedString.xlsx";
        ExcelWorkbook workbook;
        ExcelError error;

        // if file exists, remove it
        if (File.Exists(fileName))
            File.Delete(fileName);

        string fileNameTempl = @"Files\Cells\SharedString.xlsx";
        File.Copy(fileNameTempl, fileName);

        ExcelApi excelApi = new ExcelApi();
        bool res = excelApi.ExcelFileApi.OpenExcelFile(fileName, out workbook, out error);
        Assert.IsTrue(res);

        // get the first sheet
        var sheet = excelApi.ExcelSheetApi.GetSheet(workbook, 0);

        //--A1: bonjour , used 3 times, can't delete it
        Cell cellA1 = OxExcelCellValueApi.GetCell(sheet.WorkbookPart, sheet.Sheet, "A1");
        Assert.IsNotNull(cellA1);

        // The specified cell exists, is it a shared string?
        int sharedStringId= -1;
        if (cellA1.DataType != null && cellA1.DataType == CellValues.SharedString)
            // get the cell value
            cellA1.CellValue.TryGetInt(out sharedStringId);
        Assert.AreNotEqual(-1, sharedStringId);

        // remove the cell
        cellA1.Remove();
        sheet.Worksheet.Save();

        // save the count of shared string count 
        SharedStringTablePart shareStringTablePart = sheet.WorkbookPart.SharedStringTablePart;
        int sharedStringCount= shareStringTablePart.SharedStringTable.Elements<SharedStringItem>().Count();

        // remove the shared string if the cell is a text
        res = OXExcelSharedStringApi.RemoveSharedStringItem(sheet.WorkbookPart, sharedStringId);
        // not removed, used 3 times
        Assert.IsFalse(res);

        // get now the count of shared string count 
        int sharedStringCount2 = shareStringTablePart.SharedStringTable.Elements<SharedStringItem>().Count();
        Assert.AreEqual(sharedStringCount, sharedStringCount2);

        //--A6: hello, used only one time, can delete it
        Cell cellA6 = OxExcelCellValueApi.GetCell(sheet.WorkbookPart, sheet.Sheet, "A6");
        Assert.IsNotNull(cellA6);

        // remove the cell
        cellA6.Remove();
        sheet.Worksheet.Save();

        // The specified cell exists, is it a shared string?
        sharedStringId = -1;
        if (cellA6.DataType != null && cellA6.DataType == CellValues.SharedString)
            // get the cell value
            cellA6.CellValue.TryGetInt(out sharedStringId);
        Assert.AreNotEqual(-1, sharedStringId);

        // remove the shared string if the cell is a text
        res = OXExcelSharedStringApi.RemoveSharedStringItem(sheet.WorkbookPart, sharedStringId);
        // removed, used one time
        Assert.IsTrue(res);

        // get now the count of shared string count , now one is removed
        int sharedStringCount3 = shareStringTablePart.SharedStringTable.Elements<SharedStringItem>().Count();
        Assert.AreEqual(sharedStringCount, sharedStringCount3+1);

        //--B1: 12,00 - not a shared string, its a decimal
        Cell cellB1 = OxExcelCellValueApi.GetCell(sheet.WorkbookPart, sheet.Sheet, "B1");
        Assert.IsNotNull(cellB1);
        res = OxExcelCellValueApi.IsValueSharedString(sheet.WorkbookPart, cellB1);
        Assert.IsFalse(res);

        //--close the file
        res = excelApi.ExcelFileApi.CloseExcelFile(workbook, out error);
    }


}
