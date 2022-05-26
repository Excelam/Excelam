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
    /// TODO: remove;
    /// </summary>
    /// <param name="code"></param>
    /// <returns></returns>
    //public static uint Convert(ExcelCellFormatValueBase formatValue)
    //{
    //    if (formatValue.Code == ExcelCellFormatValueCode.General)
    //        return (uint) ExcelCellBuiltInFormatCode.General;


    //    if (formatValue.Code == ExcelCellFormatValueCode.Text)
    //        return (uint)ExcelCellBuiltInFormatCode.Text;

    //    if (formatValue.Code == ExcelCellFormatValueCode.Number)
    //        return (uint)ExcelCellBuiltInFormatCode.Number;

    //    if (formatValue.Code == ExcelCellFormatValueCode.Decimal)
    //        return (uint)ExcelCellBuiltInFormatCode.Decimal;

    //    if (formatValue.Code == ExcelCellFormatValueCode.DateTime)
    //        return (uint)ExcelCellBuiltInFormatCode.DateShort;

    //    //if (formatValue.Code == ExcelCellFormatValueCode.Percentage1)
    //    //    return (uint)ExcelCellBuiltInFormatCode.PercentageInt;

    //    //if (formatValue.Code == ExcelCellFormatValueCode.Percentage2)
    //    //    return (uint)ExcelCellBuiltInFormatCode.Percentage2Dec;

    //    //if (formatValue.Code == ExcelCellFormatValueCode.Scientific)
    //    //    return (uint)ExcelCellBuiltInFormatCode.Scientific;

    //    if (formatValue.Code == ExcelCellFormatValueCode.Fraction)
    //        return (uint)ExcelCellBuiltInFormatCode.Fraction;

    //    //if (formatValue.Code == ExcelCellFormatValueCode.Fraction2Digit)
    //    //    return (uint)ExcelCellBuiltInFormatCode.Fraction2Digit;


    //    // error
    //    return 0;
    //}
}
