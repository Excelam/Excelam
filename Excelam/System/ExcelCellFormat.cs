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
    /// <summary>
    /// The format key, found in StyleIndex cell property.
    /// if styleIndex is null, is set to -1.
    /// 0 is the first style index in the dictionary.
    /// </summary>
    public int StyleIndex { get; set; } = -1;

    /// <summary>
    /// More precise value format code. 
    /// Corresponds to a numberFormatId.
    /// exp: general, Number, Decimal, Fraction, DateShort, Time, Accounting, CurrencyEuro,...
    /// </summary>
    public ExcelCellFormatCode Code { get; set; } = ExcelCellFormatCode.Undefined;

    /// <summary>
    /// Set when code is a currency, in some case, when the country is identified.
    /// </summary>
    public ExcelCellCurrencyCode CurrencyCode { get; set; } = ExcelCellCurrencyCode.Unknown;

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

    //public string Formula { get; set; } = String.Empty;


    /// <summary>
    /// Return true if the cell contains a formula.
    /// </summary>
    public bool IsFormula { get; set; } = false;
    //{
    //    get {  if(string.IsNullOrEmpty(Formula)) return false; return true; } 
    //}

    /// <summary>
    /// Has no other format than value?
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
        if (ExcelNumberingFormat != null)
            s = "| " + ExcelNumberingFormat.Code + " " + ExcelNumberingFormat.FormatCode;

        // add other styles
        if (FillId > 0)
            s += ", FillId=" + FillId.ToString();
        if (BorderId > 0)
            s += ", BorderId=" + BorderId.ToString();
        if (FontId > 0)
            s += ", FontId=" + FontId.ToString();

        return "FmtId=" + NumberFormatId +"/" +Code+ s;
    }

}
