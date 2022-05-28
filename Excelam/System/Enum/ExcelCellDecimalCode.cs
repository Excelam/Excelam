using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Excelam.System;

/// <summary>
/// cell format code
/// SubCode.
/// </summary>
public enum ExcelCellDecimalCode
{
    Undefined,

    /// <summary>
    /// default case.
    /// See the decimal count (after the dot).
    /// </summary>
    Decimal,

    /// <summary>
    /// decimal with blank thousand separator.
    /// It's a built-in format=4.
    /// </summary>
    DecimalBlankThousandSep,

    /// <summary>
    /// negative value is diplayed in red.
    /// format=0.00_ ;[Red]\\-0.00\\ 
    /// </summary>
    DecimalNegRed,

    /// <summary>
    /// negative value is diplayed in red.
    /// minus sign is not displayed.
    /// format=0.00;[Red]0.00
    /// </summary>
    DecimalNegRedNoSign



}
