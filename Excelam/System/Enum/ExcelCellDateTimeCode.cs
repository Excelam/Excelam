using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Excelam.System;

/// <summary>
/// Specific Excel cell format value datetime code.
/// </summary>
public enum ExcelCellDateTimeCode
{
    Undefined,

    /// <summary>
    /// built-in code=14.
    /// </summary>
    DateShort,

    DateLarge,

    Time
}
