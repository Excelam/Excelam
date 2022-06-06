using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Excelam.System;

public class ExcelCellFormatValueScientific : ExcelCellFormatValueBase
{
    public ExcelCellFormatValueScientific()
    {
        Code = ExcelCellFormatValueCode.Scientific;
        NumberFormatId = (int)ExcelCellBuiltInFormatCode.Scientific11;
    }

    public ExcelCellScientificCode ScientificCode { get; set; } = ExcelCellScientificCode.Undefined;

}
