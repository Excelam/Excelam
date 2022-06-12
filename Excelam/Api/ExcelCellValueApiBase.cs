using DocumentFormat.OpenXml.Spreadsheet;
using Excelam.OpenXmlLayer;
using Excelam.System;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Excelam;

public abstract class ExcelCellValueApiBase
{
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
            ExcelCellFormat excelCellFormat = excelSheet.ExcelWorkbook.ExcelCellStyles.DictStyleIndexExcelCellFormat[styleIndex];
            string formula;
            if (OxExcelCellValueApi.IsCellFormula(excelSheet.WorkbookPart, cell, out formula))
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
        if (cell == null) return null;

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

    /// <summary>
    /// Concerns excel cell number value.
    /// </summary>
    /// <param name="excelSheet"></param>
    /// <param name="col"></param>
    /// <param name="row"></param>
    /// <param name="value"></param>
    /// <returns></returns>
    public bool GetCellValueAsInt(ExcelSheet excelSheet, int col, int row, out int value)
    {
        return GetCellValueAsInt(excelSheet, ExcelCellAddressApi.ConvertAddress(col, row), out value);
    }

    /// <summary>
    /// Return the cell value as an integer/excel number
    /// If the cell doesn't exists, return 0.
    /// </summary>
    /// <param name="sheet"></param>
    /// <param name="cellAddress"></param>
    /// <returns></returns>
    public bool GetCellValueAsInt(ExcelSheet excelSheet, string cellAddress, out int value)
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

    public bool GetCellValueAsDouble(ExcelSheet excelSheet, string cellAddress, out double value)
    {
        value = 0;

        Cell cell = OxExcelCellValueApi.GetCell(excelSheet.WorkbookPart, excelSheet.Sheet, cellAddress);

        // the cell doesn't exists
        if (cell == null) return false;

        return GetCellValueAsDouble(excelSheet, cell, out value);
    }

    public bool GetCellValueAsDouble(ExcelSheet excelSheet, Cell cell, out double value)
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
            GetCellValueAsDouble(excelSheet, cell, out oaDate);

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
    protected Cell CreateCell(ExcelSheet excelSheet, string colName, int rowIndex, ExcelCellFormatValueBase cellFormatValue, CellValue cellValue)
    {

        // insert an empty cell
        Cell? newCell = OxExcelCellValueApi.InsertCell(excelSheet.Worksheet, colName, rowIndex);

        // Set the value of cell
        newCell.CellValue = cellValue;

        // find a style with the same value format: general, and other format set
        int styleIndexOther = excelSheet.ExcelWorkbook.ExcelCellStyles.FindStyleIndex(cellFormatValue, 0, 0, 0);

        // no style found
        if (styleIndexOther < 0)
            styleIndexOther = ExcelCellFormatBuilder.BuildCellFormat(excelSheet.ExcelWorkbook.ExcelCellStyles, excelSheet.ExcelWorkbook.GetWorkbookStylesPart().Stylesheet, cellFormatValue, 0, 0, 0);

        // a style exists, so use it
        newCell.StyleIndex = (uint)styleIndexOther;
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
    protected bool ReplaceCellContentGeneral(ExcelSheet excelSheet, Cell cell, ExcelCellFormatValueBase cellFormatValue, CellValue cellValue)
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
        if (styleIndexSameAs < 0)
            styleIndexSameAs = ExcelCellFormatBuilder.BuildCellFormat(excelSheet.ExcelWorkbook.ExcelCellStyles, excelSheet.ExcelWorkbook.GetWorkbookStylesPart().Stylesheet, cellFormatValue, 0, 0, 0);

        // a style exists, so use it
        cell.StyleIndex = (uint)styleIndexSameAs;
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
    protected bool ReplaceCellValueAndStyle(ExcelSheet excelSheet, Cell cell, ExcelCellFormat cellFormat, ExcelCellFormatValueBase cellFormatValue, CellValue cellValue)
    {
        if (cell == null) return false;
        if (cellFormat == null) return false;

        // if the current content of the cell is a shared string, remove it
        if (OxExcelCellValueApi.IsValueSharedString(excelSheet.WorkbookPart, cell))
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
