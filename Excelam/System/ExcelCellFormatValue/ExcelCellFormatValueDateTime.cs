using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Excelam.System;

public class ExcelCellFormatValueDateTime :ExcelCellFormatValueBase
{
    ExcelCellDateTimeCode _dateTimeCode = ExcelCellDateTimeCode.Undefined;

    public ExcelCellFormatValueDateTime()
    {
        Code = ExcelCellFormatValueCategoryCode.DateTime;
        StringFormat = string.Empty; 
    }

    public ExcelCellDateTimeCode DateTimeCode 
    { 
        get { return _dateTimeCode; }
    }

    public void DefineSpecialCase(ExcelCellDateTimeCode dateTimeCode, string stringFormat)
    {
        _dateTimeCode = dateTimeCode;
        NumberFormatId = 0;
        StringFormat = stringFormat;
    }

    /// <summary>
    /// define standard cases.
    /// </summary>
    /// <param name="dateTimeCode"></param>
    public void Define(ExcelCellDateTimeCode dateTimeCode)
    {
        _dateTimeCode = dateTimeCode;

        if (dateTimeCode == ExcelCellDateTimeCode.DateShort14)
        {
            NumberFormatId = (int)ExcelCellBuiltInFormatCode.DateShort14;
            StringFormat = string.Empty;
            return;
        }

        if (dateTimeCode == ExcelCellDateTimeCode.Time21_hh_mm_ss)
        {
            NumberFormatId = (int)ExcelCellBuiltInFormatCode.Time21_hh_mm_ss;
            StringFormat = string.Empty;
            return;
        }

        if (dateTimeCode == ExcelCellDateTimeCode.Time)
        {
            NumberFormatId = 0;
            StringFormat = "[$-F400]h:mm:ss\\ AM/PM";
            return;
        }

        if (dateTimeCode == ExcelCellDateTimeCode.DateLarge)
        {
            NumberFormatId = 0;
            StringFormat = "[$-F800]dddd\\,\\ mmmm\\ dd\\,\\ yyyy";
            return;
        }

        if (dateTimeCode == ExcelCellDateTimeCode.Date_yyyy_mm_dd)
        {
            NumberFormatId = 0;
            StringFormat = "yyyy\\-mm\\-dd;@";
            return;
        }

        if (dateTimeCode == ExcelCellDateTimeCode.DateLargeEnglishUS)
        {
            NumberFormatId = 0;
            StringFormat = "[$-409]mmmm\\ d\\,\\ yyyy;@";
            return;
        }

        if (dateTimeCode == ExcelCellDateTimeCode.DateLargeGermanGermany)
        {
            NumberFormatId = 0;
            StringFormat = "[$-407]d\\.\\ mmmm\\ yyyy;@";
            return;
        }

        if (dateTimeCode == ExcelCellDateTimeCode.DateLargeGermanSwitzerland)
        {
            NumberFormatId = 0;
            StringFormat = "[$-807]d\\.\\ mmmm\\ yyyy;@";
        }
    }

    /// <summary>
    /// same subcode, same number of decimals
    /// same flag hasThousandSeperator,
    /// same negativeOtion.
    /// </summary>
    /// <param name="other"></param>
    /// <returns></returns>
    public bool AreEquals(ExcelCellFormatValueDateTime other)
    {
        if (_dateTimeCode == other.DateTimeCode)
            return true;
        return false;
    }

}

