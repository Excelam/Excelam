using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Excelam.System;

public enum ExcelCellPercentageCode
{
    Undefined,

    /// <summary>
    /// Built-in case.
    /// 9 = '0%'
    /// </summary>
    Percentage9Int,

    // 10 = '0.00%'
    Percentage10Decimal2,

    PercentageN,
    PercentageOtherCases

}
