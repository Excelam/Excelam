using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Excelam.System;

/// <summary>
/// Excel cell value, number negative option.
/// </summary>
public enum ExcelCellValueNegativeOption
{
    /// <summary>
    /// defaut color, with sign.
    /// </summary>
    Default,

    /// <summary>
    /// negative number is displayed in red, with the minus sign
    /// </summary>
    RedWithSign,

    /// <summary>
    /// negative number is displayed in red, without the minus sign
    /// </summary>
    RedWithoutSign,

}
