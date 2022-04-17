using DocumentFormat.OpenXml.Spreadsheet;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Excelam.System;

public class ExcelCellFill
{
    public int Id { get; set; } 

    /// <summary>
    /// The original excel object.
    /// </summary>
    public Fill Fill { get; set; }
}
