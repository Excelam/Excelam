using DocumentFormat.OpenXml.Spreadsheet;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Excelam.System;

/// <summary>
/// An Excel cell format.
/// Defined in the file: /xl/styles.xml.
/// Has a position defined by IndexStyle.
/// 
/// Contains format for: Value, border, fill, font and also alignment and protection.
/// </summary>
public class ExcelCellFormat
{
    public static ExcelCellFormat Create(ExcelCellFormatValueCode code)
    {
        ExcelCellFormat excelCellFormat = new();

        if (code == ExcelCellFormatValueCode.General)
            excelCellFormat.FormatValue = new ExcelCellFormatValueGeneral();
        return excelCellFormat;
    }

    /// <summary>
    /// The format key, found in StyleIndex cell property.
    /// if styleIndex is null, is set to -1.
    /// 0 is the first style index in the dictionary.
    /// </summary>
    public int StyleIndex { get; set; } = -1;

    /// <summary>
    /// Cell format value code.
    /// TODO: classe de base + hiérarchie
    /// 
    /// ExcelCellFormatValueBase  
    /// 
    ///     ExcelCellFormatValueGeneral
    ///     ExcelCellFormatValueText
    ///     ExcelCellFormatValueNumber
    ///     
    ///     ExcelCellFormatValueDecimal
    ///     ExcelCellFormatValueDateTime
    ///     ExcelCellFormatValueAccounting
    ///     ExcelCellFormatValueCurrency
    ///     ExcelCellFormatValuePercentage
    ///     ExcelCellFormatValueFraction
    ///     ExcelCellFormatValueScientific ?? faire ?
    /// enum  ExcelCellFormatValueCode
    /// </summary>
    public ExcelCellFormatValueBase FormatValue { get; set; } = null;


    /// <summary>
    /// TODO: a supprimer.
    /// </summary>
    public ExcelCellFormatStructCode StructCode { get; set; }

    /// <summary>
    /// id, from OpenXml.
    /// </summary>
    public int NumberFormatId { get; set; } = 0;

    /// <summary>
    /// Set only if a format string is present.
    /// For none built-in code, except for the accounting code/44.
    /// </summary>
    public ExcelNumberingFormat? ExcelNumberingFormat { get; set; } = null;

    public int BorderId { get; set; } = 0;

    public ExcelCellBorder? ExcelCellBorder { get; set; } = null;

    public int FillId { get; set; } = 0;

    public ExcelCellFill? ExcelCellFill { get; set; } = null;
    public int FontId { get; set; } = 0;

    public ExcelCellFont? ExcelCellFont { get; set; }

    /// <summary>
    /// Raw OpenXml tag.
    /// </summary>
    public Alignment? Alignment { get; set; } = null;

    /// <summary>
    /// Raw OpenXml tag.
    /// </summary>
    public Protection? Protection { get; set; } = null;


    /// <summary>
    /// Return true if the cell contains a formula.
    /// </summary>
    public bool IsFormula { get; set; } = false;

    /// <summary>
    /// Has no other format than value? like fill, font, border, alignement or protection.
    /// </summary>
    /// <returns></returns>
    public bool HasOtherFormatThanValue()
    {
        // at least, one other format exists
        if (FillId > 0 || FontId > 0 || BorderId > 0 || Alignment != null || Protection != null)
            return true;
        return false;
    }


    public override string ToString()
    {
        string s=string.Empty;
        string code = ExcelNumberingFormat.ValueBase.Code.ToString();

        if (ExcelNumberingFormat != null)
            // todo: revoir
            s = "| " + code + " " + ExcelNumberingFormat.FormatCode;

        // add other styles
        if (FillId > 0)
            s += ", FillId=" + FillId.ToString();
        if (BorderId > 0)
            s += ", BorderId=" + BorderId.ToString();
        if (FontId > 0)
            s += ", FontId=" + FontId.ToString();

        return "FmtId=" + NumberFormatId +"/" + code + s;
    }

}
