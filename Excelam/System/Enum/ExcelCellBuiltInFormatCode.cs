using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Excelam.System;

/// <summary>
/// 
/// Excel built-in fixed cell type.
/// This list define thz most common formatIds.
/// 
/// https://stackoverflow.com/questions/36670768/openxml-cell-datetype-is-null
/// 
/// https://stackoverflow.com/questions/4655565/reading-dates-from-openxml-excel-files
/// 
/// I found the list in ECMA-376, Second Edition, Part 1 - Fundamentals And Markup Language Reference section 18.8.30 page 1786. 
/// https://www.ecma-international.org/publications-and-standards/standards/ecma-376/
/// 
/// The missing values are mainly related to east asian variant formats.
///
/// 0 = 'General'
/// 1 = '0'
/// 2 = '0.00'
/// 3 = '#,##0'
/// 4 = '#,##0.00'  Decimal, 2dec, space as thousand separator.
/// 5 = '$#,##0;\-$#,##0'
/// 6 = '$#,##0;[Red]\-$#,##0'
/// 7 = '$#,##0.00;\-$#,##0.00'
/// 8 = '$#,##0.00;[Red]\-$#,##0.00'
/// 9 = '0%'
/// 10 = '0.00%'
/// 11 = '0.00E+00'
/// 12 = '# ?/?';
/// 13 = '# ??/??'
/// 14 = 'm/d/yyyy'  
/// 15 = 'd-mmm-yy'
/// 16 = 'd-mmm'
/// 17 = 'mmm-yy'
/// 18 = 'h:mm AM/PM'
/// 19 = 'h:mm:ss AM/PM'
/// 20 = 'h:mm'
/// 21 = 'h:mm:ss'
/// 22 = 'm/d/yy h:mm'  ou  "m/d/yyyy h:mm"
///---
/// 37 = '#,##0 ;(#,##0)'               ou "#,##0_);(#,##0)"
/// 38 = '#,##0 ;[Red](#,##0)'          ou "#,##0_);[Red]"
/// 39 = '#,##0.00;(#,##0.00)'          ou "#,##0.00_);(#,##0.00)"
/// 40 = '#,##0.00;[Red](#,##0.00)'     ou  "#,##0.00_);[Red]"
///
/// 44 = '_("$"* #,##0.00_);_("$"* \(#,##0.00\);_("$"* "-"??_);_(@_)'
/// 45 = 'mm:ss'
/// 46 = '[h]:mm:ss'
/// 47 = 'mmss.0'       ou "mm:ss.0"
/// 48 = '##0.0E+0'
/// 49 = '@'
///---
/// 27 = '[$-404]e/m/d'
/// 28 = [$-404]e"年"m"月"d"日" m"月"d"日"
/// 30 = 'm/d/yy'
/// 36 = '[$-404]e/m/d'
/// 50 = '[$-404]e/m/d'
/// 55 = 'yyyy/mm/dd'
/// 57 = '[$-404]e/m/d'
/// 59 = 't0'
/// 60 = 't0.00'
/// 61 = 't#,##0'
/// 62 = 't#,##0.00'
/// 67 = 't0%'
/// 68 = 't0.00%'
/// 69 = 't# ?/?'
/// 70 = 't# ??/??'
/// </summary>
public enum ExcelCellBuiltInFormatCode
{
    /// <summary>
    /// string
    /// </summary>
    General0 = 0,

    /// <summary>
    /// Integer
    /// </summary>
    Number1 = 1,

    /// <summary>
    /// Decimal, 2 dec after the dot.
    /// </summary>
    Decimal2 = 2,

    /// <summary>
    /// Decimal, 2 dec after the dot.
    /// blank thousand separator.
    /// </summary>
    Decimal4BlankThousandSep = 4,

    /// <summary>
    /// 9 = '0%'
    /// only int, no decimal part.
    /// </summary>
    PercentageInt = 9,

    /// <summary>
    /// 10 = '0.00%'
    /// With 2 decimal part.
    /// </summary>
    Percentage2Dec = 10,

    /// <summary>
    /// Double
    /// </summary>
    Scientific11 = 11,

    /// <summary>
    /// 12 = '# ?/?';
    /// </summary>
    Fraction12 = 12,

    /// <summary>
    /// 13 = '# ??/??'
    /// </summary>
    Fraction2Digit = 13,

    /// <summary>
    /// 'm/d/yyyy'  
    /// No format.
    /// </summary>
    DateShort14 = 14,

    /// <summary>
    ///  'hh:mm:ss'
    ///  no format.
    /// </summary>
    Time21_hh_mm_ss = 21,

    Accounting44 = 44,
    Text49 = 49


    // Special
    // Displays a number as a postal code (ZIP Code), phone number, or Social Security number.
}
