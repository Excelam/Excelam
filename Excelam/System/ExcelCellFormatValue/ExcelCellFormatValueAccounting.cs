﻿using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Excelam.System;

public class ExcelCellFormatValueAccounting:ExcelCellFormatValueBase
{
    public ExcelCellFormatValueAccounting()
    {
        Code = ExcelCellFormatValueCode.Accounting;
        NumberFormatId = (int)ExcelCellBuiltInFormatCode.Accounting44;
    }

    public ExcelCellCurrencyCode CurrencyCode { get; set; } = ExcelCellCurrencyCode.Undefined;
}
