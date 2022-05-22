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
    public static uint Convert(ExcelCellFormatMainCode code)
    {
        if (code == ExcelCellFormatMainCode.General)
            return (uint) ExcelCellBuiltInFormatCode.General;

        if (code == ExcelCellFormatMainCode.Text)
            return (uint)ExcelCellBuiltInFormatCode.Text;

        if (code == ExcelCellFormatMainCode.Number)
            return (uint)ExcelCellBuiltInFormatCode.Number;

        if (code == ExcelCellFormatMainCode.Decimal)
            return (uint)ExcelCellBuiltInFormatCode.Decimal;

        if (code == ExcelCellFormatMainCode.Percentage1)
            return (uint)ExcelCellBuiltInFormatCode.PercentageInt;

        if (code == ExcelCellFormatMainCode.Percentage2)
            return (uint)ExcelCellBuiltInFormatCode.Percentage2Dec;

        if (code == ExcelCellFormatMainCode.Scientific)
            return (uint)ExcelCellBuiltInFormatCode.Scientific;

        if (code == ExcelCellFormatMainCode.Fraction)
            return (uint)ExcelCellBuiltInFormatCode.Fraction;

        if (code == ExcelCellFormatMainCode.Fraction2Digit)
            return (uint)ExcelCellBuiltInFormatCode.Fraction2Digit;

        if (code == ExcelCellFormatMainCode.DateShort)
            return (uint)ExcelCellBuiltInFormatCode.DateShort;

        // error
        return 0;
    }
}
