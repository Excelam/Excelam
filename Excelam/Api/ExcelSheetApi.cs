using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using Excelam.System;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Excelam;

public class ExcelSheetApi
{
    /// <summary>
    /// Get an excel sheet, by index, from 0.
    /// </summary>
    /// <param name="excelDoc"></param>
    /// <param name="index"></param>
    /// <returns></returns>
    public ExcelSheet? GetSheet(ExcelWorkbook excelDoc, int index)
    {
        if (excelDoc == null) return null;

        IEnumerable<Sheet> sheets = excelDoc.SpreadsheetDocument.WorkbookPart.Workbook.GetFirstChild<Sheets>().Elements<Sheet>();
        if (index +1 > sheets.Count()) return null;

        var sheet = sheets.ToList()[index];

        string relationshipId = sheet.Id.Value;
        WorksheetPart worksheetPart = (WorksheetPart)excelDoc.SpreadsheetDocument.WorkbookPart.GetPartById(relationshipId);
        Worksheet workSheet = worksheetPart.Worksheet;
        SheetData sheetData = workSheet.GetFirstChild<SheetData>();
        IEnumerable<Row> rows = sheetData.Descendants<Row>();

        ExcelSheet excelSheet = new ExcelSheet(excelDoc, index, workSheet, sheet, rows);
        return excelSheet;
    }

    /// <summary>
    /// Get an excel sheet, by index, from 0.
    /// </summary>
    /// <param name="excelDoc"></param>
    /// <param name="index"></param>
    /// <returns></returns>
    public ExcelSheet? GetSheetByName(ExcelWorkbook excelDoc, string name)
    {
        if (excelDoc == null) return null;
        if (name == null) return null;

        IEnumerable<Sheet> sheets = excelDoc.SpreadsheetDocument.WorkbookPart.Workbook.GetFirstChild<Sheets>().Elements<Sheet>();

        int i = -1;
        Sheet? sheet = null;
        foreach (Sheet s in sheets)
        {
            i++;
            if (s.Name == null)
                continue;

            if (s.Name.Value.Equals(name, StringComparison.InvariantCultureIgnoreCase))
            {
                sheet = s;
                break;
            }
        }

        if (sheet == null) return null;

        string relationshipId = sheet.Id.Value;
        WorksheetPart worksheetPart = (WorksheetPart)excelDoc.SpreadsheetDocument.WorkbookPart.GetPartById(relationshipId);
        Worksheet workSheet = worksheetPart.Worksheet;
        SheetData sheetData = workSheet.GetFirstChild<SheetData>();
        IEnumerable<Row> rows = sheetData.Descendants<Row>();

        ExcelSheet excelSheet = new ExcelSheet(excelDoc, i, workSheet, sheet, rows);
        return excelSheet;
    }

    public int GetSheetsCount(ExcelWorkbook excelDoc)
    {
        if (excelDoc == null) return 0;

        SpreadsheetDocument spreadsheetDocument = excelDoc.SpreadsheetDocument;
        IEnumerable<Sheet> sheets = spreadsheetDocument.WorkbookPart.Workbook.GetFirstChild<Sheets>().Elements<Sheet>();
        return sheets.Count();
    }

    // GetListSheet
}
