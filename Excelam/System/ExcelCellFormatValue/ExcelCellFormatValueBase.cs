using DocumentFormat.OpenXml.Spreadsheet;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Excelam.System;

/// <summary>
/// base class of cell format value:
/// General, Text, Number Decimal, DateTime, Currency, Accounting, Percentage and Fraction.
/// </summary>
public abstract class ExcelCellFormatValueBase
{
    public ExcelCellFormatValueCode Code { get; set; } = ExcelCellFormatValueCode.Undefined;

    /// <summary>
    /// excel value.
    /// </summary>
    public uint NumberFormatId { get; set; }

    /// <summary>
    /// string format code, defined for none built-in format,
    /// except for 44/Accounting.
    /// exp: decimal with 3 decimals: 0.000 
    /// accounting/44: _-* #,##0.00\ "€"_-;\-* #,##0.00\ "€"_-;_-* "-"??\ "€"_-;_-@_-
    /// </summary>
    //public string FormatCode { get; set; } = string.Empty;
    public string StringFormat { get; set; } = string.Empty;

    public ExcelNumberingFormat ExcelNumberingFormat { get; set; }

    /// <summary>
    /// The original excel openXml object.
    /// TODO: a deplacer 
    /// </summary>
    //public NumberingFormat NumberingFormat { get; set; }

    public override string ToString()
    {
        return Code + " - " + StringFormat;
    }

}
