using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using Excelam.OpenXmlLayer;
using Excelam.System;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Excelam;

/// <summary>
/// Read/write cell value content.
/// (to create later :ExcelCellBorderApi, ExcelCellFillApi, ExcelCellFontApi).
/// </summary>
public class ExcelCellValueApi
{
    #region GetCell methods.


    public ExcelCellFormat? GetCellFormat(ExcelSheet excelSheet, int col, int row)
    {
        return GetCellFormat(excelSheet, ExcelCellAddressApi.ConvertAddress(col, row));
    }

    /// <summary>
    /// Return the format of the cell.
    /// If the cell is not set: no value, no border, no fill, nothing, return null.
    /// </summary>
    /// <param name="excelSheet"></param>
    /// <param name="cellAddress"></param>
    /// <returns></returns>
    public ExcelCellFormat? GetCellFormat(ExcelSheet excelSheet, string cellAddress)
    {
        Cell cell = OxExcelCellValueApi.GetCell(excelSheet.WorkbookPart, excelSheet.Sheet, cellAddress);


        // the cell contains nothing? no value, no border, no fill,...
        if (cell == null)
            return null;

        // the styleIndex in the some cases can be null/-1: generic cell value
        int styleIndex = OxExcelCellValueApi.GetCellStyleIndex(cell);

        // get the cell format
        if (excelSheet.ExcelWorkbook.ExcelCellStyles.DictStyleIndexExcelCellFormat.ContainsKey(styleIndex))
        {
            ExcelCellFormat excelCellFormat= excelSheet.ExcelWorkbook.ExcelCellStyles.DictStyleIndexExcelCellFormat[styleIndex];
            string formula;
            if(OxExcelCellValueApi.IsCellFormula(excelSheet.WorkbookPart,cell, out formula))
            {
                excelCellFormat.IsFormula = true;
            }
            return excelCellFormat;
        }

        // is the cell value a shared string?
        if (OxExcelCellValueApi.IsValueSharedString(excelSheet.WorkbookPart, cell))
        {
            // special case
            ExcelCellFormat excelCellFormat = ExcelCellFormat.Create(ExcelCellFormatValueCode.General);

            string formula;
            if (OxExcelCellValueApi.IsCellFormula(excelSheet.WorkbookPart, cell, out formula))
            {
                excelCellFormat.IsFormula = true;
            }

            return excelCellFormat;
        }

        return null;
    }

    /// <summary>
    /// Return the formula of the cell if exists.
    /// If not return an empty string.
    /// </summary>
    /// <param name="excelSheet"></param>
    /// <param name="col"></param>
    /// <param name="row"></param>
    /// <returns></returns>
    public string? GetCellFormula(ExcelSheet excelSheet, int col, int row)
    {
        return GetCellValueAsString(excelSheet, ExcelCellAddressApi.ConvertAddress(col, row));
    }

    /// <summary>
    /// Return the formula of the cell if exists.
    /// If not return an empty string.
    /// </summary>
    /// <param name="excelSheet"></param>
    /// <param name="cellAddress"></param>
    /// <returns></returns>
    public string? GetCellFormula(ExcelSheet excelSheet, string cellAddress)
    {
        Cell cell = OxExcelCellValueApi.GetCell(excelSheet.WorkbookPart, excelSheet.Sheet, cellAddress);


        // the cell contains nothing? no value, no border, no fill,...
        if (cell == null)
            return null;

        string formula;
        if (OxExcelCellValueApi.IsCellFormula(excelSheet.WorkbookPart, cell, out formula))
        {
            return formula;
        }
        return string.Empty;
    }

    public string? GetCellValueAsString(ExcelSheet excelSheet, int col, int row)
    {
        return GetCellValueAsString(excelSheet, ExcelCellAddressApi.ConvertAddress(col, row));
    }

    /// <summary>
    /// Return the cell value as a string.
    /// If the cell doesn't exists, return null
    /// </summary>
    /// <param name="sheet"></param>
    /// <param name="cellAddress"></param>
    /// <returns></returns>
    public string? GetCellValueAsString(ExcelSheet excelSheet, string cellAddress)
    {
        Cell cell = OxExcelCellValueApi.GetCell(excelSheet.WorkbookPart, excelSheet.Sheet, cellAddress);

        // the cell doesn't exists
        if (cell == null)return null;

        return GetCellValueAsString(excelSheet, cell);
    }

    /// <summary>
    /// Return the cell value as a string.
    /// If the cell doesn't exists, return null
    /// </summary>
    /// <param name="sheet"></param>
    /// <param name="cellAddress"></param>
    /// <returns></returns>
    public string? GetCellValueAsString(ExcelSheet excelSheet, Cell cell)
    {
        // the cell doesn't exists
        if (cell == null) return null;

        // is the cell value a shared string?
        string value;
        if (OXExcelSharedStringApi.GetCellSharedStringValue(excelSheet.ExcelWorkbook.SpreadsheetDocument.WorkbookPart, cell, out value))
            return value;

        if (cell.CellValue == null) return null;
        return cell.CellValue.InnerXml;
    }

    public bool GetCellValueAsNumber(ExcelSheet excelSheet, int col, int row, out int value)
    {
        return GetCellValueAsNumber(excelSheet, ExcelCellAddressApi.ConvertAddress(col, row), out value);
    }

    /// <summary>
    /// Return the cell value as an integer/excel number
    /// If the cell doesn't exists, return 0.
    /// </summary>
    /// <param name="sheet"></param>
    /// <param name="cellAddress"></param>
    /// <returns></returns>
    public bool GetCellValueAsNumber(ExcelSheet excelSheet, string cellAddress, out int value)
    {
        value = 0;
        string? valueStr = GetCellValueAsString(excelSheet, cellAddress);
        if (string.IsNullOrWhiteSpace(valueStr)) return false;

        int valRes;

        if (int.TryParse(valueStr, out value))
            return true;

        // error
        return false;
    }

    public bool GetCellValueAsDecimal(ExcelSheet excelSheet, string cellAddress, out double value)
    {
        value = 0;

        Cell cell = OxExcelCellValueApi.GetCell(excelSheet.WorkbookPart, excelSheet.Sheet, cellAddress);

        // the cell doesn't exists
        if (cell == null) return false;

        return GetCellValueAsDecimal(excelSheet, cell, out value); 
    }

    public bool GetCellValueAsDecimal(ExcelSheet excelSheet, Cell cell, out double value)
    {
        value = 0;
        string? valueStr = GetCellValueAsString(excelSheet, cell);
        if (string.IsNullOrWhiteSpace(valueStr)) return false;

        // exp: 10.5 -> 10,5
        valueStr = valueStr.Replace('.', ',');


        double valRes;

        if (double.TryParse(valueStr, out value))
            return true;

        // error
        return false;
    }

    /// <summary>
    /// Return the cell value as a dateTime, can be a short, a large date, or a time.
    /// </summary>
    /// <param name="excelSheet"></param>
    /// <param name="cellAddress"></param>
    /// <param name="value"></param>
    /// <returns></returns>
    public bool GetCellValueAsDateTime(ExcelSheet excelSheet, string cellAddress, out DateTime value)
    {
        value = DateTime.Now;

        Cell cell = OxExcelCellValueApi.GetCell(excelSheet.WorkbookPart, excelSheet.Sheet, cellAddress);
        if (cell == null) return false;

        // case 1, dataType is null
        if (cell.DataType == null || cell.DataType == CellValues.Number)
        {
            double oaDate;
            GetCellValueAsDecimal(excelSheet, cell, out oaDate);

            if (oaDate == 0)
            {
                value = DateTime.Now;
                return false;
            }

            value = DateTime.FromOADate(oaDate);
            return true;
        }

        // case 2, dataType is set: date or string
        if (cell.DataType == CellValues.Date || cell.DataType == CellValues.SharedString)
        {
            try
            {
                value = Convert.ToDateTime(GetCellValueAsString(excelSheet, cell));
                return true;
            }
            catch
            {
                value = DateTime.Now;
                // date format is wrong
                return false;
            }
        }

        // not the expected type
        value = DateTime.Now;
        return false;

    }


    #endregion

    #region SetCell methods.

    public bool SetCellValueGeneral(ExcelSheet excelSheet, int col, int row, string value)
    {
        return SetCellValueGeneral(excelSheet, ExcelCellAddressApi.ConvertAddress(col, row), value);
    }

    /// <summary>
    /// Set a value to a cell, type is general/standard.
    /// (rework)
    /// cases:
    ///     -1/ no cell
    ///     -2/ cell exists, same value format, can be: 
    ///         2.1/ SharedString, 2.2/ inlineString.
    ///             Can have a style or not, it's not important.
    ///             
    ///     -3/ cell exists, not the same value format
    ///         3.1/ a same style exists, use it
    ///         3.2/ no style exists, create a new one
    /// 
    /// </summary>
    /// <param name="excelSheet"></param>
    /// <param name="cellAddress"></param>
    /// <param name="value"></param>
    /// <returns></returns>
    public bool SetCellValueGeneral(ExcelSheet excelSheet, string cellAddress, string value)
    {
        // check or create the shared string table
        OXExcelSharedStringApi.CreateSharedStringTablePart(excelSheet.WorkbookPart);

        // get the cell if it exists?
        Cell cell = OxExcelCellValueApi.GetCell(excelSheet.WorkbookPart, excelSheet.Sheet, cellAddress);

        string colName;
        int rowIndex;
        ExcelCellAddressApi.SplitCellAddress(cellAddress, out colName, out _, out rowIndex);

        //--1/ the cell doesn't exists
        if (cell == null)
        {
            if (string.IsNullOrEmpty(value))
                // empty value, no cell exists, so nothing to do, bye
                return true;

            // insert a new cell with value, the type is shared string
            if (OxExcelCellValueApi.InsertCellSharedString(excelSheet.WorkbookPart, excelSheet.Worksheet, colName, rowIndex, value) == null)
                // error occurs
                return false;
            // ok job done
            return true;
        }

        //--2.1/ cell exists, same value format, is it a SharedString?
        if (OxExcelCellValueApi.IsValueSharedString(excelSheet.WorkbookPart, cell))
        {
            // set a empty value, so remove the cell
            if (string.IsNullOrEmpty(value))
            {
                // set an empty value, so delete the cell
                OxExcelCellValueApi.DeleteCell(excelSheet.WorkbookPart, excelSheet.Worksheet, cell);
                return true;
            }

            // replace the actual shared string of the cell by a new one
            OXExcelSharedStringApi.ReplaceSharedStringItem(excelSheet.WorkbookPart, cell, value);
            return true;
        }

        //--2.2/ cell exists, same value format, is it a InlineString? its a rich text
        if (OxExcelCellValueApi.IsValueInlineString(excelSheet.WorkbookPart, cell))
        {
            // set a empty value, so remove the cell
            if (string.IsNullOrEmpty(value))
            {
                // set an empty value, so delete the cell
                OxExcelCellValueApi.DeleteCell(excelSheet.WorkbookPart, excelSheet.Worksheet, cell);
                return true;
            }

            // replace the actual string of the cell by a new one
            cell.CellValue.InnerXml= value;
            return true;
        }

        //--3/ cell exists, not the same value format

        // get the style,  in the some cases can be null/-1: generic cell value
        int styleIndex = OxExcelCellValueApi.GetCellStyleIndex(cell);

        //--2.3/ cell exists, is it a a formula ?
        if (OxExcelCellValueApi.IsCellFormula(excelSheet.WorkbookPart, cell))
        {
            // set a empty value, so remove the cell
            if (string.IsNullOrEmpty(value))
            {
                // set an empty value, so delete the cell
                OxExcelCellValueApi.DeleteCell(excelSheet.WorkbookPart, excelSheet.Worksheet, cell);
                return true;
            }

            // no style found?
            if(styleIndex <0)
            {
                // set the string value
                OxExcelCellValueApi.SetCellSharedString(excelSheet.WorkbookPart, cell, value);

                // ok, general value type is set ot the cell
                return true;
            }

            // has a style, so continue with code below
        }

        // get the cell format style
        ExcelCellFormat? cellFormat = excelSheet.ExcelWorkbook.ExcelCellStyles.GetStyleByIndex(styleIndex);


        // find a style with the same value format: general, and other format set
        ExcelCellFormatValueGeneral formatValueGeneralToFind = new ExcelCellFormatValueGeneral();
        int styleIndexOther2 = excelSheet.ExcelWorkbook.ExcelCellStyles.FindStyleIndex(formatValueGeneralToFind, cellFormat.BorderId, cellFormat.FillId, cellFormat.FontId);

        // no style found?
        if (styleIndexOther2 < 0)
        {
            // build first a cell format value
            ExcelCellFormatValueGeneral formatValueGeneral = new ExcelCellFormatValueGeneral();

            // 3.2/ no style found, so have to create a new one, with existing formats
            styleIndexOther2 = ExcelCellFormatBuilder.BuildCellFormat(excelSheet.ExcelWorkbook.ExcelCellStyles, excelSheet.ExcelWorkbook.GetWorkbookStylesPart().Stylesheet, formatValueGeneral, cellFormat.BorderId, cellFormat.FillId, cellFormat.FontId);
        }

        // 3.1/ a style exists, so use it  and 3.2 case
        cell.StyleIndex = (uint)styleIndexOther2;

        // set the string value
        OxExcelCellValueApi.SetCellSharedString(excelSheet.WorkbookPart, cell, value);

        // ok, general value type is set ot the cell
        return true;
    }

    /// <summary>
    /// Set a number in a cell.
    /// If the cell exists, remove it before.
    /// 
    ///     -1/ no cell
    ///     -2/ cell exists, same value format: Number.
    ///             Can have a style or not, it's not important.
    ///             
    ///     -3/ cell exists, not the same value format
    ///         3.1/ no style index, no cellformat, type is a general/standard
    //          3.2/ a same style exists, use it
    ///         3.3/ no style exists, create a new one
    ///         
    /// </summary>
    /// <param name="excelSheet"></param>
    /// <param name="cellAddress"></param>
    /// <param name="value"></param>
    /// <returns></returns>
    /// <exception cref="Exception"></exception>
    public bool SetCellValueNumber(ExcelSheet excelSheet, string cellAddress, int value)
    {
        ExcelCellFormatValueNumber cellFormatValue;

        // get the cell if it exists?
        Cell cell = OxExcelCellValueApi.GetCell(excelSheet.WorkbookPart, excelSheet.Sheet, cellAddress);

        string colName;
        int rowIndex;
        ExcelCellAddressApi.SplitCellAddress(cellAddress, out colName, out _, out rowIndex);

        //--1/ the cell doesn't exists
        if (cell == null)
        {
            cellFormatValue = new ExcelCellFormatValueNumber();
            CreateCell(excelSheet, colName, rowIndex, cellFormatValue, new CellValue(value));
            return true;
        }

        // get the style,  in the some cases can be null/-1: generic cell value
        int styleIndex = OxExcelCellValueApi.GetCellStyleIndex(cell);

        // get the cell format style
        ExcelCellFormat? cellFormat = excelSheet.ExcelWorkbook.ExcelCellStyles.GetStyleByIndex(styleIndex);

        //--2/ cell exists, same value format: Number
        if (cellFormat!=null && cellFormat.FormatValue.Code == ExcelCellFormatValueCode.Number)
        {
            // change the cell value
            cell.CellValue = new CellValue(value);
            return true;
        }

        //--3/ cell exists, not the same value format
        //--3.1/ no style index, no cellformat, type is a general/standard
        if(styleIndex<0)
        {
            cellFormatValue = new ExcelCellFormatValueNumber();
            ReplaceCellContentGeneral(excelSheet, cell, cellFormatValue, new CellValue(value));
            return true;
        }

        // cell exists, has a style
        cellFormatValue = new ExcelCellFormatValueNumber();
        return ReplaceCellValueAndStyle(excelSheet, cell, cellFormat, cellFormatValue, new CellValue(value));
    }

    /// <summary>
    /// Set a decimal value in the cell.
    /// </summary>
    /// <param name="excelSheet"></param>
    /// <param name="cellAddress"></param>
    /// <param name="value"></param>
    /// <returns></returns>
    public bool SetCellValueDecimal(ExcelSheet excelSheet, string cellAddress, ExcelCellDecimalCode subCode, int numberOfDecimal, double value)
    {
        ExcelCellFormatValueDecimal cellFormatValue = new ExcelCellFormatValueDecimal();
        cellFormatValue.SetSubCode(subCode, numberOfDecimal);

        // get the cell if it exists?
        Cell cell = OxExcelCellValueApi.GetCell(excelSheet.WorkbookPart, excelSheet.Sheet, cellAddress);

        string colName;
        int rowIndex;
        ExcelCellAddressApi.SplitCellAddress(cellAddress, out colName, out _, out rowIndex);

        //--1/ the cell doesn't exists
        if (cell == null)
        {
            CreateCell(excelSheet, colName, rowIndex, cellFormatValue, new CellValue(value));
            return true;
        }

        // get the style,  in the some cases can be null/-1: generic cell value
        int styleIndex = OxExcelCellValueApi.GetCellStyleIndex(cell);

        // get the cell format style
        ExcelCellFormat? cellFormat = excelSheet.ExcelWorkbook.ExcelCellStyles.GetStyleByIndex(styleIndex);

        //--2/ cell exists, same value format: Number
        if (cellFormat != null && cellFormat.FormatValue.Code == ExcelCellFormatValueCode.Decimal)
        {
            // change the cell value
            cell.CellValue = new CellValue(value);
            return true;
        }

        //--3/ cell exists, not the same value format
        //--3.1/ no style index, no cellformat, type is a general/standard
        if (styleIndex < 0)
        {
            ReplaceCellContentGeneral(excelSheet, cell, cellFormatValue, new CellValue(value));
            return true;
        }

        // cell exists, has a style
        return ReplaceCellValueAndStyle(excelSheet, cell, cellFormat, cellFormatValue, new CellValue(value));
    }

    public bool SetCellValueDateShort(ExcelSheet excelSheet, int col, int row, DateTime value)
    {
        return SetCellValueDateShort(excelSheet, ExcelCellAddressApi.ConvertAddress(col, row), value);
    }

    /// <summary>
    /// set a dateShort value into a cell.
    /// </summary>
    /// <param name="excelSheet"></param>
    /// <param name="cellAddress"></param>
    /// <param name="value"></param>
    /// <returns></returns>
    public bool SetCellValueDateShort(ExcelSheet excelSheet, string cellAddress, DateTime value)
    {
        ExcelCellFormatValueDateTime cellFormatValue;

        // get the cell if it exists?
        Cell cell = OxExcelCellValueApi.GetCell(excelSheet.WorkbookPart, excelSheet.Sheet, cellAddress);

        string colName;
        int rowIndex;
        ExcelCellAddressApi.SplitCellAddress(cellAddress, out colName, out _, out rowIndex);

        //--1/ the cell doesn't exists
        if (cell == null)
        {
            cellFormatValue = new ExcelCellFormatValueDateTime();
            cellFormatValue.DateTimeCode = ExcelCellDateTimeCode.DateShort;
            CreateCell(excelSheet, colName, rowIndex, cellFormatValue, new CellValue(value.ToOADate()));
            return true;
        }

        // get the style,  in the some cases can be null/-1: generic cell value
        int styleIndex = OxExcelCellValueApi.GetCellStyleIndex(cell);

        // get the cell format style
        ExcelCellFormat? cellFormat = excelSheet.ExcelWorkbook.ExcelCellStyles.GetStyleByIndex(styleIndex);

        //--2/ cell exists, same value format: DateTime/DateShort
        // TODO: test pas complet!!
        if (cellFormat != null && cellFormat.FormatValue.Code == ExcelCellFormatValueCode.DateTime)
        {
            // change the cell value
            // todo: revoir, marche pas!!
            //ici();
            cell.CellValue = new CellValue(value.ToOADate());
            return true;
        }

        cellFormatValue = new ExcelCellFormatValueDateTime();
        cellFormatValue.DateTimeCode = ExcelCellDateTimeCode.DateShort;

        //--3/ cell exists, not the same value format
        //--3.1/ no style index, no cellformat, type is a general/standard
        if (styleIndex < 0)
        {
            ReplaceCellContentGeneral(excelSheet, cell, cellFormatValue, new CellValue(value.ToOADate()));
            return true;
        }

        // cell exists, has a style
        return ReplaceCellValueAndStyle(excelSheet, cell, cellFormat, cellFormatValue, new CellValue(value.ToOADate()));
    }

    #endregion

    #region DelCell methods.

    public bool DeleteCell(ExcelSheet excelSheet, int col, int row)
    {
        return DeleteCell(excelSheet, ExcelCellAddressApi.ConvertAddress(col, row));
    }

    /// <summary>
    /// Delete a cell.
    /// </summary>
    /// <param name="excelSheet"></param>
    /// <param name="cellAddress"></param>
    /// <returns></returns>
    public bool DeleteCell(ExcelSheet excelSheet, string cellAddress)
    {
        Cell cell = OxExcelCellValueApi.GetCell(excelSheet.WorkbookPart, excelSheet.Sheet, cellAddress);

        // the cell contains nothing? no value, no border, no fill,...
        if (cell == null) return false;

        OxExcelCellValueApi.DeleteCell(excelSheet.WorkbookPart, excelSheet.Worksheet, cell);
        return true;
    }
    #endregion

    #region Private  methods.

    /// <summary>
    /// Create a new cell, set a value.
    /// Not a shared string!
    /// </summary>
    /// <param name="excelWorkbook"></param>
    /// <param name="worksheet"></param>
    /// <param name="colName"></param>
    /// <param name="rowIndex"></param>
    /// <param name="cellValue"></param>
    /// <returns></returns>
    private Cell CreateCell(ExcelSheet excelSheet, string colName, int rowIndex, ExcelCellFormatValueBase cellFormatValue, CellValue cellValue)
    {

        // insert an empty cell
        Cell? newCell = OxExcelCellValueApi.InsertCell(excelSheet.Worksheet, colName, rowIndex);

        // Set the value of cell
        newCell.CellValue = cellValue;

        // find a style with the same value format: general, and other format set
        int styleIndexOther = excelSheet.ExcelWorkbook.ExcelCellStyles.FindStyleIndex(cellFormatValue, 0, 0, 0);

        // no style found
        if (styleIndexOther< 0)
            styleIndexOther = ExcelCellFormatBuilder.BuildCellFormat(excelSheet.ExcelWorkbook.ExcelCellStyles, excelSheet.ExcelWorkbook.GetWorkbookStylesPart().Stylesheet, cellFormatValue, 0, 0, 0);

        // a style exists, so use it
        newCell.StyleIndex = (uint) styleIndexOther;
        return newCell;
    }

    /// <summary>
    /// replace the cell content which type is general by another type.
    /// The cell type must be general, shared string!
    /// </summary>
    /// <param name="excelSheet"></param>
    /// <param name="cell"></param>
    /// <param name="cellFormatCode"></param>
    /// <param name="cellValue"></param>
    /// <returns></returns>
    private bool ReplaceCellContentGeneral(ExcelSheet excelSheet, Cell cell, ExcelCellFormatValueBase cellFormatValue, CellValue cellValue)
    {
        int sharedStringItemId = OXExcelSharedStringApi.GetSharedStringId(cell);

        // clear the shared string in the cell
        cell.DataType = null;
        cell.CellValue.Remove();

        // remove the shared string
        OXExcelSharedStringApi.RemoveSharedStringItem(excelSheet.WorkbookPart, sharedStringItemId);

        // change the cell value
        cell.CellValue = cellValue;

        // Find a style with the same cell format: Number, other formar not set
        int styleIndexSameAs = excelSheet.ExcelWorkbook.ExcelCellStyles.FindStyleIndex(cellFormatValue, 0, 0, 0);
        if (styleIndexSameAs< 0)
            styleIndexSameAs = ExcelCellFormatBuilder.BuildCellFormat(excelSheet.ExcelWorkbook.ExcelCellStyles, excelSheet.ExcelWorkbook.GetWorkbookStylesPart().Stylesheet, cellFormatValue, 0, 0, 0);
            
        // a style exists, so use it
        cell.StyleIndex = (uint) styleIndexSameAs;
        return true;
    }

    /// <summary>
    /// Replace the cell value and the style.
    /// </summary>
    /// <param name="excelSheet"></param>
    /// <param name="cell"></param>
    /// <param name="cellFormatCode"></param>
    /// <param name="cellValue"></param>
    /// <returns></returns>
    private bool ReplaceCellValueAndStyle(ExcelSheet excelSheet, Cell cell, ExcelCellFormat cellFormat, ExcelCellFormatValueBase cellFormatValue, CellValue cellValue)
    {
        if (cell == null) return false;
        if (cellFormat == null) return false;

        // if the current content of the cell is a shared string, remove it
        if(OxExcelCellValueApi.IsValueSharedString(excelSheet.WorkbookPart, cell))
        {
            int sharedStringItemId = OXExcelSharedStringApi.GetSharedStringId(cell);

            // clear the shared string in the cell
            cell.DataType = null;
            cell.CellValue.Remove();

            // remove the shared string
            OXExcelSharedStringApi.RemoveSharedStringItem(excelSheet.WorkbookPart, sharedStringItemId);
        }

        // set the new cell value
        cell.CellValue = cellValue;

        // find a similar style 
        int styleIndexSameAs2 = excelSheet.ExcelWorkbook.ExcelCellStyles.FindStyleIndex(cellFormatValue, cellFormat.BorderId, cellFormat.FillId, cellFormat.FontId);

        //--3.3/ no style exists, create a new one
        if (styleIndexSameAs2 < 0)
            styleIndexSameAs2 = ExcelCellFormatBuilder.BuildCellFormat(excelSheet.ExcelWorkbook.ExcelCellStyles, excelSheet.ExcelWorkbook.GetWorkbookStylesPart().Stylesheet, cellFormatValue, cellFormat.BorderId, cellFormat.FillId, cellFormat.FontId);

        // a style exists, so use it
        cell.StyleIndex = (uint)styleIndexSameAs2;

        return true;

    }

    #endregion
}
