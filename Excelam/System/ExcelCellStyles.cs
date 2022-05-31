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
    public Dictionary<int, ExcelCellFormat> DictStyleIndexExcelCellFormat { get; private set; } = new();

    public ExcelCellFormat? GetStyleByIndex(int styleIndex)
    {
        if (DictStyleIndexExcelCellFormat.ContainsKey(styleIndex))
            return DictStyleIndexExcelCellFormat[styleIndex];

        return null;
    }


    /// <summary>
    /// Find a style with the same all different format: value, border, fill, font.
    /// todo: ajouter alignment et protection? +compliqué!
    /// </summary>
    /// <param name="cell"></param>
    /// <returns></returns>
    public int FindStyleIndex(ExcelCellFormatValueBase formatValue, int borderId, int fillId, int fontId)
    {
        KeyValuePair<int, ExcelCellFormat> style;

        // is the format value general?
        ExcelCellFormatValueGeneral formatValueGeneral = formatValue as ExcelCellFormatValueGeneral;
        if (formatValueGeneral != null)
            return FindStyleIndexGeneral(formatValueGeneral, borderId, fillId, fontId);


        // is the format value text?
        ExcelCellFormatValueText formatValueText = formatValue as ExcelCellFormatValueText;
        if (formatValueText != null)
            return FindStyleIndexText(formatValueText, borderId, fillId, fontId);

        // is the format value number?
        ExcelCellFormatValueNumber formatValueNumber = formatValue as ExcelCellFormatValueNumber;
        if (formatValueNumber != null)
            return FindStyleIndexNumber(formatValueNumber, borderId, fillId, fontId);

        // is the format value decimal?
        ExcelCellFormatValueDecimal formatValueDecimal = formatValue as ExcelCellFormatValueDecimal;
        if (formatValueDecimal != null)
            return FindStyleIndexDecimal(formatValueDecimal, borderId, fillId, fontId);


        // TODO: other cases: DateTime, Currency, Accounting,...
        return -1;
    }

    public int FindStyleIndexGeneral(ExcelCellFormatValueGeneral formatValueGeneral, int borderId, int fillId, int fontId)
    {
        KeyValuePair<int, ExcelCellFormat> style;

        if (formatValueGeneral == null) return -1;

        style = DictStyleIndexExcelCellFormat.FirstOrDefault(cf => cf.Value.FormatValue.Code == formatValueGeneral.Code && cf.Value.BorderId == borderId && cf.Value.FillId == fillId && cf.Value.FontId == fontId);
        if (style.Value == null)
            return -1;

        return style.Key;
    }

    public int FindStyleIndexText(ExcelCellFormatValueText formatValueText, int borderId, int fillId, int fontId)
    {
        KeyValuePair<int, ExcelCellFormat> style;

        if (formatValueText == null) return -1;

        style = DictStyleIndexExcelCellFormat.FirstOrDefault(cf => cf.Value.FormatValue.Code == formatValueText.Code && cf.Value.BorderId == borderId && cf.Value.FillId == fillId && cf.Value.FontId == fontId);
        if (style.Value == null)
            return -1;

        return style.Key;
    }

    public int FindStyleIndexNumber(ExcelCellFormatValueNumber formatValueNumber, int borderId, int fillId, int fontId)
    {
        KeyValuePair<int, ExcelCellFormat> style;

        if (formatValueNumber == null) return -1;

        style = DictStyleIndexExcelCellFormat.FirstOrDefault(cf => cf.Value.FormatValue.Code == formatValueNumber.Code && cf.Value.BorderId == borderId && cf.Value.FillId == fillId && cf.Value.FontId == fontId);
        if (style.Value == null)
            return -1;

        return style.Key;
    }

    public int FindStyleIndexDecimal(ExcelCellFormatValueDecimal formatValueDecimal, int borderId, int fillId, int fontId)
    {
        KeyValuePair<int, ExcelCellFormat> style;

        if (formatValueDecimal == null) return -1;

        // get all format value concerning general type
        List<ExcelCellFormat> selectedValues = DictStyleIndexExcelCellFormat
           .Where(cf => cf.Value.FormatValue.Code == ExcelCellFormatValueCode.Decimal && cf.Value.FillId == fillId && cf.Value.FontId == fontId)
           .Select(cf => cf.Value).ToList();

        // scan items of the list: select first on subCode and NumberOfDecimal
        var item = selectedValues.Where(cf => (cf.FormatValue as ExcelCellFormatValueDecimal).AreEquals(formatValueDecimal)).FirstOrDefault();
        if (item != null) return item.StyleIndex;
        return -1;
    }
}
