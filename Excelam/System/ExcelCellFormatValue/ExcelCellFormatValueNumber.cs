using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Excelam.System;

/// <summary>
/// Built-in excel cell format value.
/// id=1.
/// </summary>
public class ExcelCellFormatValueNumber: ExcelCellFormatValueBase
{
    public ExcelCellFormatValueNumber()
    {
        Code = ExcelCellFormatValueCategoryCode.Number;
        NumberFormatId = (int)ExcelCellBuiltInFormatCode.Number1;

    }
}
