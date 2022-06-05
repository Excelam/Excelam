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
    /// exp: decimal with 3 decimals: 0.000 
    /// accounting/44: _-* #,##0.00\ "€"_-;\-* #,##0.00\ "€"_-;_-* "-"??\ "€"_-;_-@_-
    /// </summary>
    public string StringFormat { get; set; } = string.Empty;


    /// <summary>
    /// The original excel openXml object.
    /// </summary>
    public NumberingFormat? NumberingFormat { get; set; }

    public override string ToString()
    {
        return Id + " - " + StringFormat;
    }
}
