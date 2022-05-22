using DocumentFormat.OpenXml.Spreadsheet;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Excelam.System;

/// <summary>
/// List of all excel cell styles.
/// Special case: a cell can have no style defined in these lists: styleIndex=null.
///  In this case, the cell has a value and the type is standard/general
/// </summary>
public class ExcelCellStyles
{
    public List<ExcelNumberingFormat> ListExcelNumberingFormat { get; private set; } = new();
    public List<ExcelCellFill> ListExcelFill { get; private set; } = new();
    public List<ExcelCellBorder> ListExcelBorder { get; private set; } = new();

    public List<ExcelCellFont> ListExcelFont { get; private set; } = new();

    /// <summary>
    /// dictionary of styleIndex - ExcelCellFormat.
    /// </summary>
    public Dictionary<int, ExcelCellFormat> DictStyleIndexExcelStyleIndex { get; private set; } = new();

    /// <summary>
    /// Find a style with the same value format and no other format set
    /// return the style index, or -1 if not exists.
    /// </summary>
    /// <param name="code"></param>
    /// <returns></returns>
    public int FindStyle(ExcelCellFormatMainCode code, out ExcelCellFormat cellFormat)
    {
        KeyValuePair<int,ExcelCellFormat> res= DictStyleIndexExcelStyleIndex.FirstOrDefault(cf => cf.Value.StructCode.MainCode == code && !cf.Value.HasOtherFormatThanValue());

        cellFormat = null;
        if (res.Value == null)
            return - 1;

        cellFormat = res.Value;
        return res.Key;
    }

    /// <summary>
    /// Find a style with the same all different format: value, border, fill, font.
    /// </summary>
    /// <param name="cell"></param>
    /// <returns></returns>
    public int FindStyle(ExcelCellFormatMainCode code, ExcelCellCurrencyCode countryCurrency, out ExcelCellFormat cellFormat)
    {
        return FindStyle(code, countryCurrency, 0, 0, 0, out cellFormat);
    }

    public int FindStyle(ExcelCellFormatMainCode code, ExcelCellCurrencyCode countryCurrency, int borderId, int fillId, int fontId, out ExcelCellFormat cellFormat)
    {
        // todo: ajouter alignment et protection? +compliqué!
        KeyValuePair<int, ExcelCellFormat> res = DictStyleIndexExcelStyleIndex.FirstOrDefault(cf => cf.Value.StructCode.MainCode == code && cf.Value.StructCode.CurrencyCode == countryCurrency && cf.Value.BorderId == borderId && cf.Value.FillId == fillId && cf.Value.FontId == fontId);

        cellFormat = null;
        if (res.Value == null)
            return -1;

        cellFormat = res.Value;
        return res.Key;
    }
}
