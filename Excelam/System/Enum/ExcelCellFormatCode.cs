using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Excelam.System;

/// <summary>
/// General code, Excel independant (but linked to it).
/// </summary>
public enum ExcelCellFormatCode
{
    /// <summary>
    /// Undefined, the cell value is null.
    /// </summary>
    Undefined,

    /// <summary>
    /// null or empty,
    /// concerns cell having just a border.
    /// and/or colored.
    /// </summary>
    //Nothing,

    /// <summary>
    /// default cell value format: Standard/General.
    /// Built-in code.
    /// NumFormatId=0
    /// </summary>
    General,

    /// <summary>
    /// number in excel: it's an integer.
    /// Built-in code.
    /// </summary>
    Number,

    /// <summary>
    /// decimal in excel: it's a double.
    /// Built-in code.
    /// </summary>
    Decimal,

    /// <summary>
    /// 9 = '0%'
    /// only int, no decimal part.
    /// Built-in code.
    /// </summary>
    Percentage1,

    /// <summary>
    /// 10 = '0.00%'
    /// With 2 decimal part.
    /// Built-in code.
    /// </summary>
    Percentage2,

    Scientific,


    /// <summary>
    /// 12 = '# ?/?';
    /// </summary>
    Fraction,

    /// <summary>
    /// 13 = '# ??/??'
    /// </summary>
    Fraction2Digit,

    /// <summary>
    /// 0.0%
    /// None built-in.
    /// </summary>
    PercentageOneDotOne,


    /// <summary>
    /// 0.000%
    /// None built-in.
    /// </summary>
    PercentageOneDotThree,

    /// <summary>
    /// #" "?/2
    /// None built-in.
    /// </summary>
    FractionByTwo,

    DateShort,
    DateLarge,
    Time,

    /// <summary>
    /// special case, num format=44.
    /// need to save the string format and the currency.
    /// Built-in code.
    /// </summary>
    Accounting,

    /// <summary>
    /// Num format=49.
    /// Built-in code.
    /// </summary>
    Text,

    CurrencyEuro,
    CurrencyDollar,
    CurrencyPound,

    /// <summary>
    /// ¥
    /// japanese, yen
    /// 
    /// todo: pb meme symbole que china!
    /// </summary>
    CurrencyYen,

    /// <summary>
    /// yuan (china)
    /// renminbi , chinese
    /// 
    /// JP¥50 and CN¥50 when disambiguation is needed. 
    /// </summary>
    CurrencyChinese,

    /// <summary>
    /// South Korean
    /// </summary>
    CurrencyWon,

    CurrencyUkranian,
    CurrencyRussian,


    CurrencyBitcoin,
}
