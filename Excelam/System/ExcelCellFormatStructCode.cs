using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Excelam.System;

/// <summary>
/// Excel cell format code structure.
/// </summary>
public class ExcelCellFormatStructCode
{
    public ExcelCellFormatMainCode MainCode { get; set; } = ExcelCellFormatMainCode.Undefined;

    public ExcelCellDateTimeCode DateTimeCode { get; set; } = ExcelCellDateTimeCode.Undefined;

    public ExcelCellCurrencyCode CurrencyCode { get; set; } = ExcelCellCurrencyCode.Undefined;

    //		Accounting(compta),
    //		pourcentage,
    //		fraction

}
