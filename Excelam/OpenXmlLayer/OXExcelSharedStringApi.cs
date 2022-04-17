using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Excelam.OpenXmlLayer;

public class OXExcelSharedStringApi
{
    /// <summary>
    /// is the cell value a shared string? return it.
    /// </summary>
    /// <param name="workbookPart"></param>
    /// <param name="cell"></param>
    /// <param name="sharedString"></param>
    /// <returns></returns>
    public static bool IsValueSharedString(WorkbookPart workbookPart, Cell cell, out string sharedString)
    {
        sharedString = string.Empty;

        if (cell == null) return false;
        if (cell.DataType == null) return false;

        if (cell.DataType.Value != CellValues.SharedString) return false;

        SharedStringTablePart stringTablePart = workbookPart.SharedStringTablePart;
        string value = cell.CellValue.InnerXml;
        sharedString = stringTablePart.SharedStringTable.ChildElements[Int32.Parse(value)].InnerText;
        return true;
    }


    /// <summary>
    /// Replace a shared string value in a cell.
    /// </summary>
    /// <param name="workbookPart"></param>
    /// <param name="cell"></param>
    /// <param name="newText"></param>
    /// <returns></returns>
    public static int ReplaceSharedStringItem(WorkbookPart workbookPart, Cell cell, string newText)
    {
        if (cell == null) return -1;

        // check that the cell value format is a shared string
        if (!OxExcelCellValueApi.IsValueSharedString(workbookPart, cell))
            return -1;

        // get the current sharedstring id
        string textIndexStr = cell.CellValue.InnerXml;
        int textIndex = Int32.Parse(textIndexStr);

        // create a new shared string and replace it in the cell
        int newTextIndex = OXExcelSharedStringApi.InsertSharedStringItem(workbookPart.SharedStringTablePart, newText);

        // replace the shared string in the cell
        cell.CellValue = new CellValue(newTextIndex.ToString());

        // detach the shared string for the cell, remove it if not used
        OXExcelSharedStringApi.RemoveSharedStringItem(workbookPart, textIndex);

        // return the new shared string id
        return newTextIndex;
    }

    /// <summary>
    /// Given text and a SharedStringTablePart, creates a SharedStringItem with the specified text 
    /// and inserts it into the SharedStringTablePart. 
    /// If the item already exists, returns its index.
    /// </summary>
    /// <param name="shareStringPart"></param>
    /// <param name="text"></param>
    /// <returns></returns>
    public static int InsertSharedStringItem(SharedStringTablePart shareStringPart, string text)
    {
        // If the part does not contain a SharedStringTable, create one.
        if (shareStringPart.SharedStringTable == null)
        {
            shareStringPart.SharedStringTable = new SharedStringTable();
        }

        int i = 0;

        // Iterate through all the items in the SharedStringTable. If the text already exists, return its index.
        foreach (SharedStringItem item in shareStringPart.SharedStringTable.Elements<SharedStringItem>())
        {
            if (item.InnerText == text)
            {
                return i;
            }

            i++;
        }

        // The text does not exist in the part. Create the SharedStringItem and return its index.
        shareStringPart.SharedStringTable.AppendChild(new SharedStringItem(new Text(text)));
        shareStringPart.SharedStringTable.Save();

        return i;
    }

    /// <summary>
    /// Given a shared string ID and a SpreadsheetDocument, verifies that other cells in the document no longer 
    /// reference the specified SharedStringItem and removes the item.
    /// return true if the item is removed from the table.
    /// </summary>
    /// <param name="document"></param>
    /// <param name="shareStringId"></param>
    public static bool RemoveSharedStringItem(WorkbookPart workbookPart, int shareStringId)
    {
        SharedStringTablePart shareStringTablePart = workbookPart.SharedStringTablePart;
        if (shareStringTablePart == null)
            // no shared string table exists, bye
            return false;

        // find if the shared string is used 
        foreach (var part in workbookPart.GetPartsOfType<WorksheetPart>())
        {
            Worksheet worksheet = part.Worksheet;
            foreach (var cell in worksheet.GetFirstChild<SheetData>().Descendants<Cell>())
            {
                // Verify if other cells in the document reference the item.
                if (cell.DataType != null &&
                    cell.DataType.Value == CellValues.SharedString &&
                    cell.CellValue.Text == shareStringId.ToString())
                {
                    // Other cells in the document still reference the item. Do not remove the item.
                    return false;
                }
            }
        }

        // Other cells in the document do not reference the item. 
        SharedStringItem item = shareStringTablePart.SharedStringTable.Elements<SharedStringItem>().ElementAt(shareStringId);
        if (item == null) return false;

        // Remove the item.
        item.Remove();

        // Refresh all the shared string references.
        foreach (var part in workbookPart.GetPartsOfType<WorksheetPart>())
        {
            Worksheet worksheet = part.Worksheet;
            foreach (var cell in worksheet.GetFirstChild<SheetData>().Descendants<Cell>())
            {
                if (cell.DataType != null &&
                    cell.DataType.Value == CellValues.SharedString)
                {
                    int itemIndex = int.Parse(cell.CellValue.Text);
                    if (itemIndex > shareStringId)
                    {
                        cell.CellValue.Text = (itemIndex - 1).ToString();
                    }
                }
            }
            worksheet.Save();
        }

        workbookPart.SharedStringTablePart.SharedStringTable.Save();

        // ok, item removed
        return true;
    }
    
}
