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

        NumberFormatId = (int)ExcelCellBuiltInFormatCode.Decimal;
    }

    /// <summary>
    /// Get the decimal sub code.
    /// </summary>
    public ExcelCellDecimalCode SubCode 
    { 
        get { return _subCode;  }
    }

    /// <summary>
    /// Number of decimal after the dot.
    /// </summary>
    public int NumberOfDecimal { get; private set; } = 0;

    /// <summary>
    /// Set a subCode withe the number of decimal (after the dot).
    /// </summary>
    /// <param name="subCode"></param>
    /// <param name="numberOfDecimal"></param>
    public void SetSubCode(ExcelCellDecimalCode subCode, int numberOfDecimal)
    {
        if (numberOfDecimal < 0) return;

        _subCode = subCode;
        NumberOfDecimal = numberOfDecimal;

        // std case decimal=2
        if (subCode == ExcelCellDecimalCode.Decimal && numberOfDecimal== 2)
        {
            NumberFormatId = 2;
            return;
        }

        // std case decimal=4
        if (subCode == ExcelCellDecimalCode.DecimalBlankThousandSep && numberOfDecimal == 2)
        {
            NumberFormatId = 4;
            return;
        }

        // Decimal,	1, "0.0"
        if (subCode == ExcelCellDecimalCode.Decimal && numberOfDecimal == 1)
        {
            NumberFormatId = 0;
            StringFormat = "0.0";
            return;
        }

        // Decimal, 3, "0.000"
        if (subCode == ExcelCellDecimalCode.Decimal && numberOfDecimal == 3)
        {
            NumberFormatId = 0;
            StringFormat = "0.000";
            return;
        }


        // DecimalNegRed, 2, "0.00_ ;[Red]\\-0.00\\ "
        // todo:

        // DecimalNegRedNoSign, 2, "0.00;[Red]0.00"
        // todo:

    }

    public bool AreEquals(ExcelCellFormatValueDecimal other)
    {
        if (_subCode == other._subCode && NumberOfDecimal==other.NumberOfDecimal) return true;
        return false;
    }
}
