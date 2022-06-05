using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Excelam.System;

public class ExcelCellFormatValueText :ExcelCellFormatValueBase
{
    public ExcelCellFormatValueText()
    {
        Code = ExcelCellFormatValueCode.Text;
        NumberFormatId = (int)ExcelCellBuiltInFormatCode.Text49;

    }
}
