using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Excelam.System;

public class ExcelCellFormatValueCurrency:ExcelCellFormatValueBase
{
    public ExcelCellFormatValueCurrency()
    {
        Code = ExcelCellFormatValueCode.Currency;
    }
}
