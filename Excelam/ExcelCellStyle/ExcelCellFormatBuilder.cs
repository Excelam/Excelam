using DocumentFormat.OpenXml.Spreadsheet;
using Excelam.System;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Excelam;

/// <summary>
/// To build excel cell format.
/// Value, border, fill and font: later.
/// </summary>
public class ExcelCellFormatBuilder
{

    /// <summary>
    /// Create and save a new cellFormat style.
    /// </summary>
    /// <param name="stylesheet"></param>
    /// <param name="code"></param>
    /// <param name="countryCurrency"></param>
    /// <param name="borderId"></param>
    /// <param name="fillId"></param>
    /// <param name="fontId"></param>
    /// <returns></returns>
    public static int BuildCellFormat(ExcelCellStyles excelCellStyles, Stylesheet stylesheet, ExcelCellFormatValueBase formatValue, int borderId, int fillId, int fontId)
    {
        // get a numberingFormat if needed, string format set
        // TODO: pb normal?
        var excelNumberingFormat = GetOrCreateExcelNumberingFormat(excelCellStyles, stylesheet, formatValue.StringFormat);

        var cellFormat = new CellFormat();

        // create a high-level ExcelCellFormat object
        ExcelCellFormat excelCellFormat = new ExcelCellFormat();
        excelCellFormat.FormatValue = formatValue;

        if (excelNumberingFormat==null)
        {
            // it's a built-in format
            cellFormat.NumberFormatId = (uint)formatValue.NumberFormatId;
            excelCellFormat.FormatValue.ExcelNumberingFormat = null;
        }
        else
        {
            // it's a specific format, id>163
            cellFormat.NumberFormatId = excelNumberingFormat.NumberingFormat.NumberFormatId;
            excelCellFormat.FormatValue.ExcelNumberingFormat = excelNumberingFormat;
        }


        cellFormat.BorderId = (uint)borderId;
        cellFormat.FontId = (uint)fontId;
        cellFormat.FillId = (uint)fillId;

        // save the new cellFormat 
        stylesheet.CellFormats.Append(cellFormat);

        if (stylesheet.CellFormats.Count == null)
            stylesheet.CellFormats.Count = 0; 

        stylesheet.CellFormats.Count++;
        stylesheet.Save();


        excelCellFormat.BorderId = borderId;
        excelCellFormat.ExcelCellBorder = excelCellStyles.ListExcelBorder.FirstOrDefault(b => b.Id == borderId);
        excelCellFormat.FillId = fillId;
        excelCellFormat.ExcelCellFill = excelCellStyles.ListExcelFill.FirstOrDefault(b => b.Id == fillId);
        excelCellFormat.FontId = fillId;
        excelCellFormat.ExcelCellFont = excelCellStyles.ListExcelFont.FirstOrDefault(b => b.Id == fontId);
        excelCellFormat.StyleIndex = (int)(uint)excelCellStyles.DictStyleIndexExcelCellFormat.Count;

        // save in the list the new style
        excelCellStyles.DictStyleIndexExcelCellFormat.Add(excelCellFormat.StyleIndex, excelCellFormat);

        return excelCellFormat.StyleIndex;
    }

    /// <summary>
    /// Get or create a numberingFormat from/in Excel styles.
    /// </summary>
    /// <param name="excelCellStyles"></param>
    /// <param name="stylesheet"></param>
    /// <param name="stringFormat"></param>
    /// <returns></returns>
    private static ExcelNumberingFormat? GetOrCreateExcelNumberingFormat(ExcelCellStyles excelCellStyles, Stylesheet stylesheet, string stringFormat)
    {
        // no specific string format so no numbering format, bye
        if (string.IsNullOrWhiteSpace(stringFormat))
            return null;

        // try to find an existing numberingFormat
        ExcelNumberingFormat excelNumberingFormat = excelCellStyles.FindExcelNumberingFormat(stringFormat);

        // exists?
        if (excelNumberingFormat != null)
            return excelNumberingFormat;

        // have to create first a new NumberingFormat
        excelCellStyles.MaxNumberingFormatId++;

        NumberingFormat numberingFormat = new NumberingFormat();
        numberingFormat.NumberFormatId = (uint) excelCellStyles.MaxNumberingFormatId;
        numberingFormat.FormatCode =stringFormat;

        // get the manager, can be null
        if(stylesheet.NumberingFormats==null)
        {
            // create it
            stylesheet.NumberingFormats = new NumberingFormats();
        }

        NumberingFormats numberingFormats = stylesheet.NumberingFormats;

        // save it
        numberingFormats.Append(numberingFormat);
        numberingFormats.Count = (uint)numberingFormats.ChildElements.Count;

        excelNumberingFormat = new ExcelNumberingFormat();
        excelNumberingFormat.Id = (int)numberingFormat.NumberFormatId.Value;
        excelNumberingFormat.StringFormat = stringFormat;
        excelNumberingFormat.NumberingFormat= numberingFormat;
        // save it
        excelCellStyles.ListExcelNumberingFormat.Add(excelNumberingFormat);
        return excelNumberingFormat;
    }

}
