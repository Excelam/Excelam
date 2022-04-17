using DocumentFormat.OpenXml.Packaging;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Excelam.System;

/// <summary>
/// Represent an Excel Workbook.
/// </summary>
public class ExcelWorkbook
{
    /// <summary>
    /// Constructor.
    /// </summary>
    /// <param name="fileName"></param>
    /// <param name="spreadsheetDocument"></param>
    /// <param name="dictUriOpenXmlPart"></param>
    /// <param name="excelCellStyles"></param>
    public ExcelWorkbook(string fileName, SpreadsheetDocument spreadsheetDocument, Dictionary<string, OpenXmlPart> dictUriOpenXmlPart, ExcelCellStyles excelCellStyles)
    {
        FileName = fileName;
        SpreadsheetDocument = spreadsheetDocument;
        DictUriOpenXmlPart = dictUriOpenXmlPart;
        ExcelCellStyles = excelCellStyles;
    }
    public string FileName { get; private set; }

    /// <summary>
    /// Excel object
    /// </summary>
    public SpreadsheetDocument SpreadsheetDocument { get; private set; }

    /// <summary>
    ///  dictionnary of Uri and corresponding OpenXmlPart.
    ///  exp: 
    ///     /xl/workbook.xml, WorkbookPart
    ///     /xl/styles.xml, WorkbookStylesPart
    ///     ...
    /// </summary>
    public Dictionary<string, OpenXmlPart> DictUriOpenXmlPart { get; private set; }

    /// <summary>
    /// get all dynamic style format: date and currency by styleIndex.
    /// It's specific for each file!
    /// </summary>
    public ExcelCellStyles ExcelCellStyles { get; private set; }

    public WorkbookStylesPart GetWorkbookStylesPart()
    {
        if (!DictUriOpenXmlPart.ContainsKey("/xl/styles.xml"))
            return null;

        return (WorkbookStylesPart)DictUriOpenXmlPart["/xl/styles.xml"];
    }

}
