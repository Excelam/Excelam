using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Excelam.System;

/// <summary>
/// represents a decimal excel cell format value.
/// </summary>
public class ExcelCellFormatValueDecimal :ExcelCellFormatValueBase
{
    ExcelCellDecimalCode _subCode = ExcelCellDecimalCode.Undefined;

    /// <summary>
    /// Constructor.
    /// </summary>
    public ExcelCellFormatValueDecimal()
    {
        Code = ExcelCellFormatValueCode.Decimal;

        // todo: not in all cases!
        NumberFormatId = (uint)ExcelCellBuiltInFormatCode.Decimal;

    }

    /// <summary>
    /// Get the decimal sub code.
    /// </summary>
    public ExcelCellDecimalCode SubCode 
    { 
        get { return _subCode;  }
        //set { SetSubCode(value); }
    }

    public int NumberOfDecimal { get; private set; } = 0;

    public void SetSubCode(ExcelCellDecimalCode subCode, int numberOfDecimal)
    {
        if (numberOfDecimal < 0) return;

        if(subCode == ExcelCellDecimalCode.Decimal && numberOfDecimal== 2)
            NumberFormatId = 2;

        if (subCode == ExcelCellDecimalCode.DecimalBlankThousandSep && numberOfDecimal == 2)
            NumberFormatId = 2;

        _subCode = subCode;
        NumberOfDecimal = numberOfDecimal;
    }
}
