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
    /// <summary>
    /// max id existing in the ExcelNumberingFormat list.
    /// The min value is 163 (Excel/OpenXml specification).
    /// </summary>
    public int MaxNumberingFormatId { get; set; } = 163;

    /// <summary>
    /// List of ExcelNumberingFormat (none built-in/specific cell value format).
    /// exp: "0.00;[Red]0.00" representing a Decimal, 2 decimal, negative: red, no sign.
    /// </summary>
    public List<ExcelNumberingFormat> ListExcelNumberingFormat { get; private set; } = new();
    public List<ExcelCellFill> ListExcelFill { get; private set; } = new();
    public List<ExcelCellBorder> ListExcelBorder { get; private set; } = new();

    public List<ExcelCellFont> ListExcelFont { get; private set; } = new();

    /// <summary>
    /// dictionary of styleIndex - ExcelCellFormat.
    /// </summary>
    public Dictionary<int, ExcelCellFormat> DictStyleIndexExcelCellFormat { get; private set; } = new();

    /// <summary>
    /// Save the list of numbering format.
    /// Get the highest numberingFormatId.
    /// </summary>
    /// <param name="listExcelNumberingFormat"></param>
    public void SetListExcelNumberingFormat(List<ExcelNumberingFormat> listExcelNumberingFormat)
    {
        if (listExcelNumberingFormat == null) return;
        ListExcelNumberingFormat.Clear();
        ListExcelNumberingFormat.AddRange(listExcelNumberingFormat);

        // Get the highest numberingFormatId.
        foreach(ExcelNumberingFormat excelNumberingFormat in ListExcelNumberingFormat.Where(nf => nf.NumberingFormat != null))
        {
            if (excelNumberingFormat.NumberingFormat.NumberFormatId > MaxNumberingFormatId)
                MaxNumberingFormatId = (int)(uint)excelNumberingFormat.NumberingFormat.NumberFormatId;
        }
    }

    /// <summary>
    /// Find a numbering format corresponding to the string format.
    /// </summary>
    /// <param name="stringFormat"></param>
    /// <returns></returns>
    public ExcelNumberingFormat FindExcelNumberingFormat(string stringFormat)
    {
        if (string.IsNullOrEmpty(stringFormat))
            return null;

        return ListExcelNumberingFormat.Where(nf => nf.StringFormat.Equals(stringFormat)).FirstOrDefault();
    }

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
