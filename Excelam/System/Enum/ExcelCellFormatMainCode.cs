using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Excelam.System;

/// <summary>
/// Excel cell format: Main code, Excel independant (but linked to it).
/// TODO: remove it!!
/// </summary>
public enum ExcelCellFormatMainCode_REMOVE
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

    /// <summary>
    /// For date, Time and datetime.
    /// </summary>
    DateTime,

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

    /// <summary>
    /// see the other enum defined the currency.
    /// </summary>
    Currency,

}
