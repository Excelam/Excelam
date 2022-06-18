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
/// 
/// TODO: rename
/// ExcelCellFormatValueCategoryBase
/// </summary>
public abstract class ExcelCellFormatValueBase
{
    public ExcelCellFormatValueCategoryCode Code { get; set; } = ExcelCellFormatValueCategoryCode.Undefined;

    /// <summary>
    /// excel/openXml number format id.
    /// </summary>
    public int NumberFormatId { get; set; }

    /// <summary>
    /// string format code, defined for none built-in format,
    /// except for 44/Accounting.
    /// exp: decimal with 3 decimals: 0.000 
    /// accounting/44: _-* #,##0.00\ "€"_-;\-* #,##0.00\ "€"_-;_-* "-"??\ "€"_-;_-@_-
    /// 
    /// TODO: supprimer!! est dans numberingFormat
    /// 
    /// </summary>
    public string StringFormat { get; set; } = string.Empty;

    public ExcelNumberingFormat ExcelNumberingFormat { get; set; }

    public override string ToString()
    {
        return Code + " - " + StringFormat;
    }

}
