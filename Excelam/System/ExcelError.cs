using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Excelam.System;

public class ExcelError
{
    public ExcelError()
    {
        Msg = "";
    }
    public ExcelErrorCode Code { get; set; }

    /// <summary>
    /// exp: "C2"
    /// </summary>
    public string CellAddress { get; set; }

    /// <summary>
    /// the cell raw text value.
    /// </summary>
    public string CellValue { get; set; }
    public string Msg { get; set; }
    public Exception Exception { get; set; }


    public static ExcelError Create(ExcelErrorCode code, string msg)
    {
        ExcelError error = new ExcelError();
        error.Code = code;
        error.Msg = msg;
        return error;
    }
}
