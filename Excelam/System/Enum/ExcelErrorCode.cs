using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Excelam.System;

public enum ExcelErrorCode
{
    ExcelFileNameIsNull,
    ExcelFileAlreadyExists,
    UnableToCreateExcelFile,
    ExcelFileNotFound,
    UnableToOpenExcelFile,
    UnableToCloseExcelFile,

    

    ExcelSheetNameIsNull,

    CellIsNull,

    CellValIntExpected,
    CellValDecimalExpected,

    CellValPercentageExpected,
    CellValScientificExpected,
    CellValFractionExpected,

    CellValDateShortExpected,
    CellValDateLongExpected,

    CellValCurrencyEuroExpected,
    CellValCurrencyDollarExpected,
    CellValCurrencyLivreSterlingExpected,

    CellValTypeNotManaged,

    UnableDecodeStyleFormat,

}
