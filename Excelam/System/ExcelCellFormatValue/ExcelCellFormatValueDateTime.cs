using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Excelam.System;

public class ExcelCellFormatValueDateTime :ExcelCellFormatValueBase
{
    public ExcelCellFormatValueDateTime()
    {
        Code = ExcelCellFormatValueCode.DateTime;
    }

    public ExcelCellDateTimeCode DateTimeCode { get; set; } = ExcelCellDateTimeCode.Undefined;
}
