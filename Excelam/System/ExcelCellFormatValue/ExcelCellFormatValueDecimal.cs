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
    ExcelCellDecimalCode _decimalCode = ExcelCellDecimalCode.Undefined;

    /// <summary>
    /// Constructor.
    /// </summary>
    public ExcelCellFormatValueDecimal()
    {
        Code = ExcelCellFormatValueCode.Decimal;

        NumberFormatId = (int)ExcelCellBuiltInFormatCode.Decimal2;
    }

    /// <summary>
    /// Get the decimal sub code.
    /// </summary>
    public ExcelCellDecimalCode DecimalCode 
    { 
        get { return _decimalCode;  }
    }

    /// <summary>
    /// Number of decimal after the dot.
    /// </summary>
    public int NumberOfDecimal { get; private set; } = 0;

    public bool HasThousandSeparator { get; private set; } = false;

    public ExcelCellValueNegativeOption NegativeOption { get; private set; } = ExcelCellValueNegativeOption.Default;

    /// <summary>
    /// TODO: refactorer! passer numberOfDecimal, NegativeOption et HasThousandSeparator.
    /// -> va set le bon subCode.
    /// 
    /// Set a subCode withe the number of decimal (after the dot).
    /// </summary>
    /// <param name="subCode"></param>
    /// <param name="numberOfDecimal"></param>
    public void Define(int numberOfDecimal, bool hasThousandSeparator, ExcelCellValueNegativeOption negativeOption)
        //ExcelCellDecimalCode subCode, int numberOfDecimal)
    {
        // not a number, should have at least one decimal after the dot
        if (numberOfDecimal < 1) numberOfDecimal = 2;

        NumberOfDecimal = numberOfDecimal;
        HasThousandSeparator =hasThousandSeparator;
        NegativeOption = negativeOption;

        // std case decimal=2
        if (numberOfDecimal==2 && !hasThousandSeparator && negativeOption== ExcelCellValueNegativeOption.Default)
        {
            _decimalCode = ExcelCellDecimalCode.Decimal2;
            NumberFormatId = 2;
            return;
        }

        // std case decimal=4
        if (numberOfDecimal == 2 && hasThousandSeparator && negativeOption == ExcelCellValueNegativeOption.Default)
        {
            _decimalCode = ExcelCellDecimalCode.Decimal4BlankThousandSep;
            NumberFormatId = 4;
            return;
        }

        // default value
        NumberFormatId = 0;
        _decimalCode = ExcelCellDecimalCode.DecimalN;

        // Decimal, 1:  "0.0"
        if (numberOfDecimal == 1 && !hasThousandSeparator && negativeOption == ExcelCellValueNegativeOption.Default)
        {
            StringFormat = "0.0";
            return;
        }
        // Decimal, 3 "0.000"
        if (numberOfDecimal == 3 && !hasThousandSeparator && negativeOption == ExcelCellValueNegativeOption.Default)
        {
            StringFormat = "0.000";
            return;
        }

        // Decimal, 2 decimal, negative: red
        if (numberOfDecimal == 2 && !hasThousandSeparator && negativeOption == ExcelCellValueNegativeOption.RedWithSign)
        {
            StringFormat = "0.00_ ;[Red]\\-0.00\\ ";
            return;
        }

        // Decimal, 2 decimal, negative: red, no sign
        if (numberOfDecimal == 2 && !hasThousandSeparator && negativeOption == ExcelCellValueNegativeOption.RedWithoutSign)
        {
            StringFormat = "0.00;[Red]0.00";
        }
    }
        
    /// <summary>
    /// same subcode, same number of decimals
    /// same flag hasThousandSeperator,
    /// same negativeOtion.
    /// </summary>
    /// <param name="other"></param>
    /// <returns></returns>
    public bool AreEquals(ExcelCellFormatValueDecimal other)
    {
        if (_decimalCode == other._decimalCode && 
            NumberOfDecimal==other.NumberOfDecimal &&
            HasThousandSeparator== other.HasThousandSeparator &&
            NegativeOption== other.NegativeOption) 
                return true;
        return false;
    }
}
