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
        if(formatValueGeneral!=null)
        {
            style= DictStyleIndexExcelCellFormat.FirstOrDefault(cf => cf.Value.FormatValue.Code == formatValue.Code && cf.Value.BorderId == borderId && cf.Value.FillId == fillId && cf.Value.FontId == fontId);
            if (style.Value == null)
                return -1;

            return style.Key;
        }

        // is the format value text?
        ExcelCellFormatValueText formatValueText = formatValue as ExcelCellFormatValueText;
        if (formatValueText != null)
        {
            style = DictStyleIndexExcelCellFormat.FirstOrDefault(cf => cf.Value.FormatValue.Code == formatValue.Code && cf.Value.BorderId == borderId && cf.Value.FillId == fillId && cf.Value.FontId == fontId);
            if (style.Value == null)
                return -1;

            return style.Key;
        }

        // is the format value number?
        // todo:
         
        // is the format value decimal?
        ExcelCellFormatValueDecimal formatValueDecimal = formatValue as ExcelCellFormatValueDecimal;
        if (formatValueDecimal != null)
        {
            // get all format value concerning general type
            // todo:
            List<ExcelCellFormat> selectedValues = DictStyleIndexExcelCellFormat
               .Where(cf => cf.Value.FormatValue.Code == ExcelCellFormatValueCode.Decimal && cf.Value.FillId == fillId && cf.Value.FontId == fontId)
               .Select(cf => cf.Value).ToList();

            // scan the list
            // todo:
        }


        // TODO: ne va pas marcher suite rework
        // todo: ajouter alignment et protection? +compliqué!
        KeyValuePair<int, ExcelCellFormat> res = DictStyleIndexExcelCellFormat.FirstOrDefault(cf => cf.Value.FormatValue.Code == formatValue.Code  && cf.Value.BorderId == borderId && cf.Value.FillId == fillId && cf.Value.FontId == fontId);

        if (res.Value == null)
            return -1;

        return res.Key;
    }

    /// <summary>
    /// Find a style with the same value format and no other format set
    /// return the style index, or -1 if not exists.
    /// TODO: pb avec les autres infos: borderId, FillId, checker à 0 non??
    /// </summary>
    /// <param name="code"></param>
    /// <returns></returns>
    public int FindStyle(ExcelCellFormatValueBase formatValue)
    {
        // todo: va pas marcher dans tous les cas!!!
        KeyValuePair<int, ExcelCellFormat> res = DictStyleIndexExcelCellFormat.FirstOrDefault(cf => cf.Value.FormatValue.Code == formatValue.Code && !cf.Value.HasOtherFormatThanValue());

        if (res.Value == null)
            return -1;

        return res.Key;
    }

}
