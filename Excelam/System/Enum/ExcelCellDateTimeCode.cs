using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Excelam.System;

/// <summary>
/// Specific Excel cell format value datetime code.
/// list of codes: https://blog.csdn.net/2066/article/details/45555
/// </summary>
public enum ExcelCellDateTimeCode
{
    Undefined,

    /// <summary>
    /// built-in code.
    ///  14 = 'm/d/yyyy'  
    /// </summary>
    DateShort14,

    /// <summary>
    /// built-in code.
    /// 21 = 'hh:mm:ss'
    /// </summary>
    Time21_hh_mm_ss,

    /// <summary>
    /// format: "yyyy\\-mm\\-dd;@"	
    /// exp: 2022-02-14
    /// </summary>
    Date_yyyy_mm_dd,

    DateLarge,

    Time,

    /// <summary>
    /// "[$-409]mmmm\\ d\\,\\ yyyy;@"
    /// English, US
    /// </summary>
    DateLargeEnglishUS,

    /// <summary>
    /// "[$-407]d\\.\\ mmmm\\ yyyy;@"
    /// German, Germany
    /// </summary>
    DateLargeGermanGermany,

    /// <summary>
    /// "[$-807]d\\.\\ mmmm\\ yyyy;@"
    /// German, Switzerland
    /// </summary>
    DateLargeGermanSwitzerland,

    /// <summary>
    /// DateTime other cases.
    /// </summary>
    DateTimeOtherCases,

    /// <summary>
    /// Date other cases.
    /// </summary>
    DateOtherCases,

    /// <summary>
    /// Time other cases.
    /// </summary>
    TimeOtherCases


}
