using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Excelam.System;

public class ExcelCellFormatValuePercentage: ExcelCellFormatValueBase
{
    public ExcelCellFormatValuePercentage()
    { 
        Code = ExcelCellFormatValueCategoryCode.Percentage;
        // todo: not in all cases!
        NumberFormatId = (int)ExcelCellBuiltInFormatCode.Percentage9Int;
    }

    public ExcelCellPercentageCode PercentageCode { get; set; } = ExcelCellPercentageCode.Undefined;

    /// <summary>
    /// Number of decimal after the dot.
    /// </summary>
    public int NumberOfDecimal { get; set; } = 0;


}
