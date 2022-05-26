using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Excelam.System;

public class ExcelCellFormatValueGeneral: ExcelCellFormatValueBase
{
    public ExcelCellFormatValueGeneral()
    {
        Code = ExcelCellFormatValueCode.General;
        NumberFormatId = (uint)ExcelCellBuiltInFormatCode.General;
    }
}
