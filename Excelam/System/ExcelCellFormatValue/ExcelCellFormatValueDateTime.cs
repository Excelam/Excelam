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
        Code = ExcelCellFormatValueCode.DateTime;
    }

    public ExcelCellDateTimeCode DateTimeCode 
    { 
        get { return _dateTimeCode; }
        set 
        { 
            if(value== ExcelCellDateTimeCode.DateShort14)
                NumberFormatId= (int)ExcelCellBuiltInFormatCode.DateShort14;
            _dateTimeCode = value;
        } 
    } 
}
