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
    public static int BuildCellFormat(ExcelCellStyles excelCellStyles, Stylesheet stylesheet, ExcelCellFormatCode code, ExcelCellCountryCurrency countryCurrency, int borderId, int fillId, int fontId)
    {
        var cellFormat = new CellFormat();
        cellFormat.NumberFormatId = ExcelCellFormatValueConverter.Convert(code);

        cellFormat.BorderId = (uint)borderId;
        cellFormat.FontId = (uint)fontId;
        cellFormat.FillId = (uint)fillId;

        // save the new cellFormat 
        stylesheet.CellFormats.Append(cellFormat);
        stylesheet.Save();

        // create a higu-level ExcelCellFormat object
        ExcelCellFormat excelCellFormat = new ExcelCellFormat();
        excelCellFormat.NumberFormatId = (int)(uint)cellFormat.NumberFormatId;
        excelCellFormat.ExcelNumberingFormat = excelCellStyles.ListExcelNumberingFormat.FirstOrDefault(i => i.Id == cellFormat.NumberFormatId);
        excelCellFormat.Code = code;
        excelCellFormat.CountryCurrency = countryCurrency;
        excelCellFormat.BorderId = borderId;
        excelCellFormat.ExcelCellBorder = excelCellStyles.ListExcelBorder.FirstOrDefault(b => b.Id == borderId);
        excelCellFormat.FillId = fillId;
        excelCellFormat.ExcelCellFill = excelCellStyles.ListExcelFill.FirstOrDefault(b => b.Id == fillId);
        excelCellFormat.FontId = fillId;
        excelCellFormat.ExcelCellFont = excelCellStyles.ListExcelFont.FirstOrDefault(b => b.Id == fontId);
        excelCellFormat.StyleIndex = (int)(uint)stylesheet.CellFormats.Count;

        // save in the list the new style
        excelCellStyles.DictStyleIndexExcelStyleIndex.Add(excelCellFormat.StyleIndex, excelCellFormat);

        return (int)(uint)stylesheet.CellFormats.Count;
    }
}
