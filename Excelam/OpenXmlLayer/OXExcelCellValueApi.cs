using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Excelam.OpenXmlLayer;

/// <summary>
/// OpenXml layer.
/// </summary>
public class OxExcelCellValueApi
{
    /// <summary>
    /// is the cell value a shared string?
    /// </summary>
    /// <param name="workbookPart"></param>
    /// <param name="cell"></param>
    /// <returns></returns>
    public static bool IsValueSharedString(WorkbookPart workbookPart, Cell cell)
    {
        if (cell == null) return false;
        if (cell.DataType == null) return false;

        if (cell.DataType.Value != CellValues.SharedString) return false;

        return true;
    }

    /// <summary>
    /// is the cell value an inline string?
    /// It's a rich text.
    /// </summary>
    /// <param name="workbookPart"></param>
    /// <param name="cell"></param>
    /// <returns></returns>
    public static bool IsValueInlineString(WorkbookPart workbookPart, Cell cell)
    {
        if (cell == null) return false;
        if (cell.DataType == null) return false;

        if (cell.DataType.Value != CellValues.InlineString) return false;

        return true;
    }

    /// <summary>
    /// is the cell value a string?
    /// It's a formula.
    /// </summary>
    /// <param name="workbookPart"></param>
    /// <param name="cell"></param>
    /// <returns></returns>
    public static bool IsCellFormula(WorkbookPart workbookPart, Cell cell)
    {
        if (cell == null) return false;

        if (cell.CellFormula != null) 
            // current way, its a formula
            return true;

        // oldest way to define a formula
        if (cell.DataType == null) return false;

        if (cell.DataType.Value != CellValues.String) return false;

        return true;
    }

    /// <summary>
    /// is the cell value a formula?
    /// It's a formula.
    /// </summary>
    /// <param name="workbookPart"></param>
    /// <param name="cell"></param>
    /// <returns></returns>
    public static bool IsCellFormula(WorkbookPart workbookPart, Cell cell, out string formula)
    {
        formula = string.Empty;

        if (cell == null) return false;

        // current way
        if (cell.CellFormula != null)
        {
            formula = cell.CellFormula.InnerText;
            return true;
        }

        // older way
        if (cell.DataType == null) return false;

        if (cell.DataType.Value != CellValues.String) return false;

        if(cell.CellFormula != null)
            formula = cell.CellFormula.InnerText;
        return true;
    }

    /// <summary>
    /// Return the styleIndex value of the cell.
    /// If doesn't exists, return -1.
    /// </summary>
    /// <param name="cell"></param>
    /// <returns></returns>
    public static int GetCellStyleIndex(Cell cell)
    {
        int styleIndex = -1;
        if (cell.StyleIndex == null) return styleIndex;

        if(cell.StyleIndex.HasValue)
            return (int)cell.StyleIndex.Value;
        return styleIndex;
    }


    /// <summary>
    /// Create a new cell in the sheet.
    /// Remove the previous one.
    /// remove the sharedString if its the case.
    /// 
    /// TODO: refactorer
    /// </summary>
    /// <param name="excelDoc"></param>
    /// <param name="excelSheet"></param>
    /// <param name="cellAddress"></param>
    /// <param name="newCell"></param>
    /// <returns></returns>
    public static bool CreateCellRemovePrevious(WorkbookPart workbookPart, Worksheet worksheet, string colName, int rowIndex, out Cell newCell)
    {
        newCell = null;
        if (workbookPart == null) return false;
        if (worksheet == null) return false;

        // Get the cell at the specified column and row, to remove it
        Cell cell = GetCell(worksheet, colName, rowIndex);
        if (cell != null)
        {
            // delete the cell
            DeleteCell(workbookPart, worksheet, cell);
        }

        // insert a new cell
        newCell =InsertCell(worksheet, colName, rowIndex);

        return true;
    }

    /// <summary>
    /// Insert a new cell with the value, the type is shared string.
    /// Get or create a shared string.
    /// </summary>
    /// <param name="workbookPart"></param>
    /// <param name="worksheet"></param>
    /// <param name="colName"></param>
    /// <param name="rowIndex"></param>
    /// <param name="value"></param>
    /// <returns></returns>
    public static Cell? InsertCellSharedString(WorkbookPart workbookPart, Worksheet worksheet, string colName, int rowIndex, string value)
    {
        // get or create the shared string value
        int textIndex = OXExcelSharedStringApi.InsertSharedStringItem(workbookPart.SharedStringTablePart, value);

        // insert an empty cell
        Cell? newCell = OxExcelCellValueApi.InsertCell(worksheet, colName, rowIndex);

        // Set the value of cell
        newCell.DataType = new EnumValue<CellValues>(CellValues.SharedString);
        newCell.CellValue = new CellValue(textIndex.ToString());
        return newCell;
    }

    /// <summary>
    /// Set a shared string item to a cell value.
    /// </summary>
    /// <param name="workbookPart"></param>
    /// <param name="cell"></param>
    /// <param name="text"></param>
    /// <returns></returns>
    public static int SetCellSharedString(WorkbookPart workbookPart, Cell cell, string text)
    {
        int textIndex = OXExcelSharedStringApi.InsertSharedStringItem(workbookPart.SharedStringTablePart, text);

        if (OXExcelSharedStringApi.InsertSharedStringItem(workbookPart.SharedStringTablePart, text) < 0)
            return -1;

        cell.DataType = new EnumValue<CellValues>(CellValues.SharedString);
        // replace the shared string in the cell
        cell.CellValue = new CellValue(textIndex.ToString());

        // clear the formula string
        RemoveCellFormula(workbookPart, cell);

        return textIndex;
    }

    /// <summary>
    /// Remove the formula of the cell.
    /// Update the workbookpart.
    /// </summary>
    /// <param name="workbookPart"></param>
    /// <param name="cell"></param>
    public static void RemoveCellFormula(WorkbookPart workbookPart, Cell cell)
    {
        // clear the formula string
        if (cell.CellFormula == null)
            return;

        CalculationChainPart calculationChainPart = workbookPart.CalculationChainPart;
        CalculationChain calculationChain = calculationChainPart.CalculationChain;
        var calculationCells = calculationChain.Elements<CalculationCell>().ToList();

        string cellRef = cell.CellReference;
        CalculationCell calculationCell = calculationCells.Where(c => c.CellReference == cellRef).FirstOrDefault();

        cell.CellFormula.Remove();
        if (calculationCell != null)
        {
            calculationCell.Remove();
            calculationCells.Remove(calculationCell);
        }
        else
        {
            //Something is went wrong - log it
        }
    
        if (calculationCells.Count == 0)
                workbookPart.DeletePart(calculationChainPart);
    }

    /// <summary>
    /// Given a column name, a row index, and a WorksheetPart, inserts a cell into the worksheet. 
    /// If the cell already exists, returns it. 
    /// return true if its a new cell.
    /// </summary>
    /// <param name="columnName"></param>
    /// <param name="rowIndex"></param>
    /// <param name="worksheetPart"></param>
    /// <returns></returns>
    public static Cell? InsertCell(Worksheet worksheet, string columnName, int rowIndex)
    {
        Cell cellInserted;
        SheetData sheetData = worksheet.GetFirstChild<SheetData>();
        string cellReference = columnName + rowIndex;

        // If the worksheet does not contain a row with the specified row index, insert one.
        Row row;
        if (sheetData.Elements<Row>().Where(r => r.RowIndex == rowIndex).Count() != 0)
        {
            row = sheetData.Elements<Row>().Where(r => r.RowIndex == rowIndex).First();
        }
        else
        {
            row = new Row() { RowIndex = (uint)rowIndex };
            sheetData.Append(row);
        }

        // If there is not a cell with the specified column name, insert one.  
        if (row.Elements<Cell>().Where(c => c.CellReference.Value == columnName + rowIndex).Count() > 0)
        {
            cellInserted = row.Elements<Cell>().Where(c => c.CellReference.Value == cellReference).First();
            // the cell exists
            return null;
        }
        else
        {
            // Cells must be in sequential order according to CellReference. Determine where to insert the new cell.
            Cell refCell = null;
            foreach (Cell cell in row.Elements<Cell>())
            {
                if (string.Compare(cell.CellReference.Value, cellReference, true) > 0)
                {
                    refCell = cell;
                    break;
                }
            }

            cellInserted = new Cell() { CellReference = cellReference };
            row.InsertBefore(cellInserted, refCell);

            worksheet.Save();

            // new cell inserted
            return cellInserted;
        }
    }

    /// <summary>
    /// get the excel cell object, by the address.
    /// exp: A1
    /// </summary>
    /// <param name="workSheetPart"></param>
    /// <param name="cellAddress"></param>
    /// <returns></returns>
    public static Cell? GetCell(WorkbookPart workbookPart, Sheet sheet, string cellAddress)
    {
        if (workbookPart == null) return null;
        string relationshipId = sheet.Id.Value;
        WorksheetPart worksheetPart = (WorksheetPart)workbookPart.GetPartById(relationshipId);

        if (worksheetPart == null) return null;

        Cell cell = worksheetPart.Worksheet.Descendants<Cell>()
                                    .SingleOrDefault(c => cellAddress.Equals(c.CellReference));

        return cell;
    }

    /// <summary>
    /// Given a worksheet, a column name, and a row index, gets the cell at the specified column and row
    /// TODO: autre méthode GetCell, diff avec celle du dessus??
    /// </summary>
    /// <param name="worksheet"></param>
    /// <param name="columnName"></param>
    /// <param name="rowIndex"></param>
    /// <returns></returns>
    public static Cell? GetCell(Worksheet worksheet, string columnName, int rowIndex)
    {
        IEnumerable<Row> rows = worksheet.GetFirstChild<SheetData>().Elements<Row>().Where(r => r.RowIndex == rowIndex);
        if (rows.Count() == 0)
        {
            // A cell does not exist at the specified row.
            return null;
        }

        IEnumerable<Cell> cells = rows.First().Elements<Cell>().Where(c => string.Compare(c.CellReference.Value, columnName + rowIndex, true) == 0);
        if (cells.Count() == 0)
        {
            // A cell does not exist at the specified column, in the specified row.
            return null;
        }

        return cells.First();
    }

    /// <summary>
    /// Delete a cell.
    /// if the cell contains a text, remove it from the shared string table (if its the case).
    /// </summary>
    /// <param name="spreadsheetDocument"></param>
    /// <param name="worksheet"></param>
    /// <param name="cell"></param>
    public static void DeleteCell(WorkbookPart workbookPart, Worksheet worksheet, Cell cell)
    {
        int shareStringId = -1;

        // The specified cell exists, is it a shared string?
        if (cell.DataType != null && cell.DataType == CellValues.SharedString)
        {
            // get the cell value
            cell.CellValue.TryGetInt(out shareStringId);
        }

        // The specified cell exists
        cell.Remove();
        worksheet.Save();

        // remove the shared string if the cell is a text
        if (shareStringId > -1)
            OXExcelSharedStringApi.RemoveSharedStringItem(workbookPart, shareStringId);
    }

    #region Private methods.


    #endregion
}
