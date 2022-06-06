using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Excelam.System;

public enum ExcelCellScientificCode
{
    Undefined,

    /// <summary>
    /// Built-in, 11.
    /// format is: '0.00E+00'
    /// </summary>
    Scientific11,

    ScientificOtherCases

}
