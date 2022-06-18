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
/// </summary>
public class ExcelCellValueApi: ExcelCellValueApiBase
{

    #region SetCellValueGeneral methods.

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
    #endregion    

    #region SetCellValueNumber methods.

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
        if (cellFormat!=null && cellFormat.FormatValue.Code == ExcelCellFormatValueCategoryCode.Number)
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

    #endregion

    #region SetCellValueDecimal methods.

    public bool SetCellValueDecimal(ExcelSheet excelSheet, int col, int row, int numberOfDecimal, bool hasThousandSeparator, ExcelCellValueNegativeOption negativeOption, double value)
    {
        return SetCellValueDecimal(excelSheet, ExcelCellAddressApi.ConvertAddress(col, row), numberOfDecimal, hasThousandSeparator, negativeOption, value);
    }

    /// <summary>
    /// Set a decimal value in the cell.
    /// </summary>
    /// <param name="excelSheet"></param>
    /// <param name="cellAddress"></param>
    /// <param name="value"></param>
    /// <returns></returns>
    public bool SetCellValueDecimal(ExcelSheet excelSheet, string cellAddress, int numberOfDecimal, bool hasThousandSeparator, ExcelCellValueNegativeOption negativeOption, double value)
    {        
        ExcelCellFormatValueDecimal cellFormatValue = new ExcelCellFormatValueDecimal();
        cellFormatValue.Define(numberOfDecimal, hasThousandSeparator, negativeOption);

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
        if (cellFormat != null && cellFormat.FormatValue.Code == ExcelCellFormatValueCategoryCode.Decimal)
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

    #endregion

    #region SetCellValue DateTime methods.

    /// <summary>
    /// Set cell value as a dateShort.
    /// It's a built-in format, code=14.
    /// </summary>
    /// <param name="excelSheet"></param>
    /// <param name="cellAddress"></param>
    /// <param name="value"></param>
    /// <returns></returns>
    public bool SetCellValueDateShort(ExcelSheet excelSheet, string cellAddress, DateTime value)
    {
        return SetCellValueDateTime(excelSheet, cellAddress, ExcelCellDateTimeCode.DateShort14, value);
    }

    /// <summary>
    /// set a dateTime value into a cell.
    /// </summary>
    /// <param name="excelSheet"></param>
    /// <param name="cellAddress"></param>
    /// <param name="value"></param>
    /// <returns></returns>
    public bool SetCellValueDateTime(ExcelSheet excelSheet, string cellAddress, ExcelCellDateTimeCode dateTimeCode, DateTime value)
    {
        ExcelCellFormatValueDateTime cellFormatValue;

        // check the code
        // TODO: dateTimeCode pas les specialCode!! OtherCode

        // get the cell if it exists?
        Cell cell = OxExcelCellValueApi.GetCell(excelSheet.WorkbookPart, excelSheet.Sheet, cellAddress);

        string colName;
        int rowIndex;
        ExcelCellAddressApi.SplitCellAddress(cellAddress, out colName, out _, out rowIndex);

        //--1/ the cell doesn't exists
        if (cell == null)
        {
            cellFormatValue = new ExcelCellFormatValueDateTime();
            cellFormatValue.Define(dateTimeCode);
            CreateCell(excelSheet, colName, rowIndex, cellFormatValue, new CellValue(value.ToOADate()));
            return true;
        }

        // get the style,  in the some cases can be null/-1: generic cell value
        int styleIndex = OxExcelCellValueApi.GetCellStyleIndex(cell);

        // get the cell format style
        ExcelCellFormat? cellFormat = excelSheet.ExcelWorkbook.ExcelCellStyles.GetStyleByIndex(styleIndex);

        //--2/ cell exists, same value format: DateTime/DateShort
        // TODO: test pas complet!!
        if (cellFormat != null && cellFormat.FormatValue.Code == ExcelCellFormatValueCategoryCode.DateTime)
        {
            // change the cell value
            cell.CellValue = new CellValue(value.ToOADate());
            return true;
        }

        cellFormatValue = new ExcelCellFormatValueDateTime();
        cellFormatValue.Define(dateTimeCode);

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

    #region SetCellValue Currency methods.
    #endregion

    #region SetCellValue Accounting methods.
    #endregion
}
