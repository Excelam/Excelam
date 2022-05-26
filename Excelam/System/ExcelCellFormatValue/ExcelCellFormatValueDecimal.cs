using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Excelam.System;

/// <summary>
/// represents a decimal excel cell format value.
/// </summary>
public class ExcelCellFormatValueDecimal :ExcelCellFormatValueBase
{
    public ExcelCellFormatValueDecimal()
    {
        Code = ExcelCellFormatValueCode.Decimal;

        // todo: not in all cases!
        NumberFormatId = (uint)ExcelCellBuiltInFormatCode.Decimal;

    }

    public int NumberOfDecimal { get; set; } = 0;
}
