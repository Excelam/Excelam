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
        Code = ExcelCellFormatValueCategoryCode.General;
        NumberFormatId = (int)ExcelCellBuiltInFormatCode.General0;
    }
}
