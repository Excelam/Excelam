using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Excelam.System;

public class ExcelCellFormatValueFraction: ExcelCellFormatValueBase
{
    public ExcelCellFormatValueFraction()
    {
        Code = ExcelCellFormatValueCode.Fraction;
        // todo: not in all cases!
        NumberFormatId = (int)ExcelCellBuiltInFormatCode.Fraction12;

    }
}
