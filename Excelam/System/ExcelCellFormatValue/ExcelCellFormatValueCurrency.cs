using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Excelam.System;

/// <summary>
/// Represents a currency format cell value.
/// </summary>
public class ExcelCellFormatValueCurrency:ExcelCellFormatValueBase
{
    ExcelCellCurrencyCode _currencyCode = ExcelCellCurrencyCode.Undefined;

    public ExcelCellFormatValueCurrency()
    {
        Code = ExcelCellFormatValueCode.Currency;
    }

    public ExcelCellCurrencyCode CurrencyCode
    {
        get { return _currencyCode; }
    }

    /// <summary>
    /// Number of decimal after the dot.
    /// </summary>
    public int NumberOfDecimal { get; private set; } = 0;

    public ExcelCellValueNegativeOption NegativeOption { get; private set; } = ExcelCellValueNegativeOption.Default;

    /// <summary>
    /// Define a cell value as a currency.
    /// </summary>
    /// <param name="subCode"></param>
    /// <param name="numberOfDecimal"></param>
    public void Define(ExcelCellCurrencyCode currencyCode, int numberOfDecimal, ExcelCellValueNegativeOption negativeOption)
    {
        // not a number, should have at least one decimal after the dot
        if (numberOfDecimal < 1) numberOfDecimal = 2;

        _currencyCode = currencyCode;
        NumberOfDecimal = numberOfDecimal;
        NegativeOption = negativeOption;
        // no built-in case exists for currency format
        NumberFormatId = 0;

        // euro, 2 dec, has thousand sep, neg: std
        if (currencyCode== ExcelCellCurrencyCode.Euro && numberOfDecimal == 2 && negativeOption == ExcelCellValueNegativeOption.Default)
        {
            StringFormat = "#,##0.00\\ \"€\"";
            return;
        }

        if (currencyCode == ExcelCellCurrencyCode.UnitedStatesDollar && numberOfDecimal == 2 && negativeOption == ExcelCellValueNegativeOption.Default)
        {
            StringFormat = "[$$-409]#,##0.00";
            return;
        }

        // TODO: add others code
    }

    /// <summary>
    /// same subcode, same number of decimals
    /// same flag hasThousandSeperator,
    /// same negativeOtion.
    /// </summary>
    /// <param name="other"></param>
    /// <returns></returns>
    public bool AreEquals(ExcelCellFormatValueCurrency other)
    {
        if (_currencyCode == other._currencyCode &&
            NumberOfDecimal == other.NumberOfDecimal &&
            NegativeOption == other.NegativeOption)
            return true;
        return false;
    }
}
