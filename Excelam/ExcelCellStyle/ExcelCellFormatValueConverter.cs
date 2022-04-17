using Excelam.System;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Excelam;

/// <summary>
/// Convert from ExcelCellFormatValue high level objects to OpenXml number format.
/// </summary>
public class ExcelCellFormatValueConverter
{
    /// <summary>
    /// Convert from an excel cell format code to an OpenXml numbering format.
    /// only built-in code are decoded!
    /// </summary>
    /// <param name="code"></param>
    /// <returns></returns>
    public static uint Convert(ExcelCellFormatCode code)
    {
        if (code == ExcelCellFormatCode.General)
            return (uint) ExcelCellBuiltInFormatCode.General;

        if (code == ExcelCellFormatCode.Text)
            return (uint)ExcelCellBuiltInFormatCode.Text;

        if (code == ExcelCellFormatCode.Number)
            return (uint)ExcelCellBuiltInFormatCode.Number;

        if (code == ExcelCellFormatCode.Decimal)
            return (uint)ExcelCellBuiltInFormatCode.Decimal;

        if (code == ExcelCellFormatCode.Percentage1)
            return (uint)ExcelCellBuiltInFormatCode.PercentageInt;

        if (code == ExcelCellFormatCode.Percentage2)
            return (uint)ExcelCellBuiltInFormatCode.Percentage2Dec;

        if (code == ExcelCellFormatCode.Scientific)
            return (uint)ExcelCellBuiltInFormatCode.Scientific;

        if (code == ExcelCellFormatCode.Fraction)
            return (uint)ExcelCellBuiltInFormatCode.Fraction;

        if (code == ExcelCellFormatCode.Fraction2Digit)
            return (uint)ExcelCellBuiltInFormatCode.Fraction2Digit;

        if (code == ExcelCellFormatCode.DateShort)
            return (uint)ExcelCellBuiltInFormatCode.DateShort;

        // error
        return 0;
    }
}
