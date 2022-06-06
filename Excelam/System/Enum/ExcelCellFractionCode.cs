using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Excelam.System;

public enum ExcelCellFractionCode
{
    Undefined,

    /// <summary>
    /// Built-in.
    /// 12= '# ?/?'
    /// </summary>
    Fraction12,

    /// <summary>
    /// Built-in
    // 13 = '# ??/??'
    /// </summary>
    Fraction13,


    FractionByTwo,

}
