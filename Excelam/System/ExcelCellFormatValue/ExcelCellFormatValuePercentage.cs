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
        Code = ExcelCellFormatValueCode.Percentage;
        // todo: not in all cases!
        NumberFormatId = (int)ExcelCellBuiltInFormatCode.Percentage2Dec;

    }
}
