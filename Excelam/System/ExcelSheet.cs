using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Excelam.System;

public class ExcelSheet
{
    /// <summary>
    /// Constructor.
    /// </summary>
    /// <param name="excelWorkbook"></param>
    /// <param name="index"></param>
    /// <param name="workSheet"></param>
    /// <param name="sheet"></param>
    /// <param name="rows"></param>
    public ExcelSheet(ExcelWorkbook excelWorkbook, int index, Worksheet workSheet, Sheet sheet, IEnumerable<Row> rows)
    {
        WorkbookPart = excelWorkbook.SpreadsheetDocument.WorkbookPart;
        ExcelWorkbook = excelWorkbook;
        Index = index;
        Worksheet = workSheet;
        Sheet = sheet;
        Rows = rows;
    }

    /// <summary>
    /// The OpenXml workbookpart.
    /// </summary>
    public WorkbookPart WorkbookPart { get; private set; }

    /// <summary>
    /// The mworkbook parent main object.
    /// </summary>
    public ExcelWorkbook ExcelWorkbook { get; private set; }


    public int Index { get; private set; }

    /// <summary>
    /// The OpenXml worksheet.
    /// </summary>
    public Worksheet Worksheet { get; private set; }

    /// <summary>
    /// The openXml sheet.
    /// </summary>
    public Sheet Sheet { get; private set; }

    /// <summary>
    /// The OpenXml rows.
    /// </summary>
    public IEnumerable<Row> Rows { get; private set; }

    public string GetName()
    {
        if (Sheet.Name == null) return string.Empty;
        return Sheet.Name.Value;
    }
}
