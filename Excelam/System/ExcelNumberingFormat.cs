using DocumentFormat.OpenXml.Spreadsheet;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Excelam.System;

/// <summary>
/// Wrapper on Excel openXml NumberingFormat class.
/// </summary>
public class ExcelNumberingFormat
{
    public int Id { get; set; }

    /// <summary>
    /// string format code, defined for none built-in format,
    /// except for 44/Accounting.
    /// </summary>
    public string FormatCode { get; set; } = string.Empty;

    /// <summary>
    /// More precise code.
    /// </summary>
    public ExcelCellFormatCode Code { get; set; }

    /// <summary>
    /// Set when the code is a currency in some case.
    /// </summary>
    public ExcelCellCurrencyCode CurrencyCode { get; set; } = ExcelCellCurrencyCode.Undefined;

    /// <summary>
    /// The original excel openXml object.
    /// </summary>
    public NumberingFormat NumberingFormat { get; set; }

    public override string ToString()
    {
        return Id + " - " + FormatCode;
    }
}
