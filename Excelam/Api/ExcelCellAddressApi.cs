using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Excelam;

/// <summary>
/// Excel cell address Api.
/// </summary>
public static class ExcelCellAddressApi
{
    /// <summary>
    /// return the column name (letter) and the column index.
    /// 
    /// exp: 
    /// B6  -> return B,2,6
    /// AB3 -> return AB,28,3
    /// </summary>
    /// <param name="colRowName"></param>
    /// <param name="colName"></param>
    /// <param name="colIndex"></param>
    /// <returns></returns>
    public static bool SplitCellAddress(string colRowName, out string colName, out int colIndex, out int rowIndex)
    {
        colName = "";
        colIndex = 0;
        rowIndex = 0;

        if (string.IsNullOrEmpty(colRowName)) return false;
        colRowName = colRowName.ToUpperInvariant();

        // remove the row index
        for (int i = 0; i < colRowName.Length; i++)
        {
            // not a digit, continue
            if (!char.IsLetter(colRowName[i])) break;
            colName = String.Concat(colName, colRowName[i]);
        }

        if (colName == string.Empty) return false;

        int sum = 0;

        for (int i = 0; i < colName.Length; i++)
        {
            sum *= 26;
            sum += (int)(colName[i] - 'A' + 1);
        }
        colIndex = sum;


        // get the row index
        string rowStr = colRowName.Remove(0, colName.Length);

        if (rowStr.Length == 0)
            return false;
        if (!int.TryParse(rowStr, out rowIndex))
            return false;

        return true;
    }

    /// <summary>
    /// Convert to a standard excel address.
    /// exp: 1,1 -> A1
    /// </summary>
    /// <param name="col"></param>
    /// <param name="rowIndex"></param>
    /// <returns></returns>
    public static string ConvertAddress(int colIndex, int rowIndex)
    {
        if (colIndex < 1) return string.Empty;
        if (rowIndex < 1) return string.Empty;

        return GetColumnName(colIndex) + rowIndex.ToString();
    }

    /// <summary>
    /// return the column name of the col index.
    /// exp: 1 -> A
    /// </summary>
    /// <param name="index"></param>
    /// <returns></returns>
    public static string GetColumnName(int index)
    {
        if (index < 1) return String.Empty;

        index--;
        const string letters = "ABCDEFGHIJKLMNOPQRSTUVWXYZ";

        var value = "";

        if (index >= letters.Length)
            value += letters[index / letters.Length - 1];

        value += letters[index % letters.Length];

        return value;
    }

}
