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
    /// Defined in some cases.
    /// </summary>
    public string StringFormat { get; set; }= string.Empty;
}
