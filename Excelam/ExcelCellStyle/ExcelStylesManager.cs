using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using Excelam.System;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Excelam;

/// <summary>
/// Excel styles manager.
/// </summary>
public static class ExcelStylesManager
{
    /// <summary>
    /// get all dynamic style format: date and currency.
    /// It's specific for each file!!
    /// 
    /// https://stackoverflow.com/questions/4655565/reading-dates-from-openxml-excel-files
    /// 
    /// numberingFormats:
    /// https://stackoverflow.com/questions/4730152/what-indicates-an-office-open-xml-cell-contains-a-date-time-value
    /// 
    /// NumFormatId>= 164:
    /// [$-F800] dddd\,\ mmmm\ dd\,\ yyyy   date longue
    /// [$-F400] h:mm:ss\ AM/PM             heure
    ///
    /// #,##0.00\ "€"				Currency, Euro
    /// [$$-409]#,##0.00			Currency, Dollar, US
    /// [$$-C09]#,##0.00			Currency, Dollar, Australien
    /// 
    /// 00000					    Code postal France
    /// 
    /// 
    /// ExcelCellFormatType -> que pour les types builtin, <=164.
    /// 
    /// ExcelCellFormatManager
    /// GetStylesCellFormat()
    /// ExcelCellFormatDecoder
    /// </summary>
    /// <param name="document"></param>
    public static bool GetStylesCellFormat(Dictionary<string, OpenXmlPart> dictUriOpenXmlPart, out ExcelCellStyles excelCellStyles, out ExcelError excelError)
    {
        excelCellStyles = new ExcelCellStyles();
        excelError = null;

        try
        {
            if (!dictUriOpenXmlPart.ContainsKey("/xl/styles.xml"))
                return true;

            // The only way to tell dates from numbers is by looking at the style index. 
            // This indexes cellXfs, which contains NumberFormatIds, which index NumberingFormats.
            // This method creates a list of the style indexes that pertain to dates.
            WorkbookStylesPart workbookStylesPart = (WorkbookStylesPart)dictUriOpenXmlPart["/xl/styles.xml"];
            Stylesheet styleSheet = workbookStylesPart.Stylesheet;

            // get the cell formats list, if exists
            CellFormats cellFormats = styleSheet.CellFormats;
            if (cellFormats == null)
                // nothing so bye 
                return true;

            // load and decode all numberingFormat item
            NumberingFormats numberingFormats = styleSheet.NumberingFormats;
            excelCellStyles.ListExcelNumberingFormat.AddRange(LoadListExcelNumberingFormat(styleSheet.NumberingFormats));
            excelCellStyles.ListExcelBorder.AddRange(LoadListExcelBorder(styleSheet.Borders));
            excelCellStyles.ListExcelFill.AddRange(LoadListExcelFill(styleSheet.Fills));
            excelCellStyles.ListExcelFont.AddRange(LoadListExcelFont(styleSheet.Fonts));
            //Alignment alignment;
            //Protection protection;

            // convert all ExcelNumberingFormat into ExcelCellFormatValue
            OxExcelCellFormatValueDecoder.Decode(excelCellStyles.ListExcelNumberingFormat);

            int styleIndex = 0;

            //--scan each existing cellFormat and decode it: value, border, fill, font, Alignment and protection
            foreach (CellFormat cellFormat in cellFormats)
            {
                // decode the number format                
                uint numberFormatId = 0;
                if (cellFormat.NumberFormatId != null) numberFormatId = cellFormat.NumberFormatId;

                var excelNumberingFormat = excelCellStyles.ListExcelNumberingFormat.FirstOrDefault(i => i.Id == (int)numberFormatId);
                //ExcelCellCurrencyCode currencyCode;
                //ExcelCellDateTimeCode dateTimeCode;
                ExcelCellFormatStructCode code = DecodeNumberFormatId((int)numberFormatId, excelNumberingFormat);


                uint borderId = 0;
                if (cellFormat.BorderId != null) borderId = cellFormat.BorderId;

                uint fillId = 0;
                if(cellFormat.FillId !=null) fillId= cellFormat.FillId;

                uint fontId=0;
                if (cellFormat.FontId != null) fontId = cellFormat.FontId;

                // build an excelformat object
                ExcelCellFormat excelCellFormat = new()
                {
                    StyleIndex = styleIndex,
                    NumberFormatId = (int)numberFormatId,

                    StructCode = code,
                    //DateTimeCode= dateTimeCode,
                    //CurrencyCode = currencyCode,
                    ExcelNumberingFormat = excelNumberingFormat,

                    // todo: besoin de ce champ?
                    BorderId = (int)borderId,
                    ExcelCellBorder= excelCellStyles.ListExcelBorder.FirstOrDefault(b=>b.Id== (int)borderId),
                    FillId = (int)fillId,
                    ExcelCellFill = excelCellStyles.ListExcelFill.FirstOrDefault(fi => fi.Id == (int)fillId),
                    FontId = (int)fontId,
                    ExcelCellFont = excelCellStyles.ListExcelFont.FirstOrDefault(fo => fo.Id == (int)fontId),

                    Alignment= cellFormat.Alignment,
                    Protection= cellFormat.Protection
                };

                // save the decoded cell style
                excelCellStyles.DictStyleIndexExcelStyleIndex.Add(styleIndex, excelCellFormat);

                // next cell style
                styleIndex++;
            }

            return true;
        }
        catch (Exception e)
        {
            excelError = new ExcelError();
            excelError.Code = ExcelErrorCode.UnableDecodeStyleFormat;
            excelError.Exception = e;
            return false;
        }
    }

    /// <summary>
    /// Create a simple default style sheet, used for a new excel file.
    /// </summary>
    /// <returns></returns>
    public static Stylesheet CreateEmptyStylesheet()
    {
        Fonts fonts = new Fonts(
                     // Index 0 - default
                     new Font(
                         new FontSize() { Val = 11 }
                     ));

        Fills fills = new Fills(
            // Index 0 - default
            new Fill(new PatternFill() { PatternType = PatternValues.None })
        );

        Borders borders = new Borders(
            // index 0 default
            new Border());

        // create a first empty cellFormat, strange but mandatory!
        var cellFormat = new CellFormat();
        cellFormat.NumberFormatId = 0;
        CellFormats cellFormats = new CellFormats();
        cellFormats.Append(cellFormat);

        Stylesheet styleSheet = new Stylesheet(fonts, fills, borders, cellFormats);

        return styleSheet;
    }

    /// <summary>
    /// Get or create a style sheet.
    /// Should exists in normal case.
    /// </summary>
    /// <param name="workbookStylesPart"></param>
    /// <returns></returns>
    //public static Stylesheet GetOrCreateStylesheet(WorkbookStylesPart workbookStylesPart)
    //{
    //    Stylesheet stylesheet = workbookStylesPart.Stylesheet;
    //    if (stylesheet == null)
    //    {
    //        // build a simple styles sheet
    //        stylesheet = CreateEmptyStylesheet();
    //        stylesheet.Save();
    //        return stylesheet;
    //    }

    //    // the cellFormat should not exists, create it
    //    if (stylesheet.CellFormats == null)
    //    {
    //        // create a first empty cellFormat, strange but mandatory!!
    //        var cellFormat = new CellFormat();
    //        cellFormat.NumberFormatId = 0;
    //        CellFormats cellFormats = new CellFormats();
    //        cellFormats.Append(cellFormat);

    //        stylesheet.Append(cellFormats);
    //        stylesheet.Save();
    //    }

    //    return stylesheet;
    //}

    #region Private methods.

    private static ExcelCellFormatStructCode DecodeNumberFormatId(int numberFormatId, ExcelNumberingFormat excelNumberingFormat)
    {
        //countryCurrency = ExcelCellCurrencyCode.Undefined;
        //dateTimeCode = ExcelCellDateTimeCode.Undefined;

        if (excelNumberingFormat != null)
        {
            // already decoded
            //countryCurrency = excelNumberingFormat.CurrencyCode;
            return excelNumberingFormat.Code;
        }

        // if null, its a builtin case, decode it to obtain the code
        ExcelCellFormatStructCode code;
        OxExcelCellFormatValueDecoder.DecodeNumberingFormat(numberFormatId,string.Empty, out code);
        return code;
    }
    

    private static List<ExcelNumberingFormat> LoadListExcelNumberingFormat(NumberingFormats numberingFormats)
    {
        List<ExcelNumberingFormat> listExcelNumberingFormat = new();
        if (numberingFormats == null) return listExcelNumberingFormat;

        numberingFormats.Cast<NumberingFormat>().ToList().ForEach(numberingFormat =>
        {
            ExcelNumberingFormat excelNumberingFormat = new();
            excelNumberingFormat.Id = (int)numberingFormat.NumberFormatId.Value;
            if(numberingFormat.FormatCode!=null)
                excelNumberingFormat.FormatCode = numberingFormat.FormatCode;

            excelNumberingFormat.NumberingFormat = numberingFormat;
            listExcelNumberingFormat.Add(excelNumberingFormat);
        });

        return listExcelNumberingFormat;
    }

    private static List<ExcelCellFill> LoadListExcelFill(Fills fills)
    {
        List<ExcelCellFill> listExcelFill = new();
        if (fills == null) return listExcelFill;
        int i = 1;
        fills.Cast<Fill>().ToList().ForEach(fill =>
        {
            ExcelCellFill excelFill = new();
            excelFill.Id = i;
            //excelFill.Id = fill.PatternFill;
            excelFill.Fill = fill;
            i++;

            listExcelFill.Add(excelFill);
        });

        return listExcelFill;
    }

    private static List<ExcelCellBorder> LoadListExcelBorder(Borders borders)
    {
        List<ExcelCellBorder> listExcelBorder = new();
        if (borders == null) return listExcelBorder;
        int i = 1;
        borders.Cast<Border>().ToList().ForEach(border =>
        {
            ExcelCellBorder excelBorder = new();
            excelBorder.Id = i;
            excelBorder.Border = border;
            i++;

            listExcelBorder.Add(excelBorder);
        });

        return listExcelBorder;
    }

    private static List<ExcelCellFont> LoadListExcelFont(Fonts fonts)
    {
        List<ExcelCellFont> listExcelFont = new();
        if (fonts == null) return listExcelFont;
        int i = 1;
        fonts.Cast<Font>().ToList().ForEach(font =>
        {
            ExcelCellFont excelFont = new();
            excelFont.Id = i;
            excelFont.Font = font;
            i++;

            listExcelFont.Add(excelFont);
        });

        return listExcelFont;
    }

    #endregion
}
