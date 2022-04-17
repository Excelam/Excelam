using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using Excelam.System;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Excelam;

public class ExcelFileApi
{
    public ExcelFileApi()
    { }

    public string DefaultFirstSheetName
    {
        get { return "MySheet"; }
    }

    public bool CreateExcelFile(string fileName, string sheetName, out ExcelWorkbook? excelDoc, out ExcelError? error)
    {
        excelDoc = null;

        if (!CheckFileName(fileName, true, out error))
            return false;

        if (string.IsNullOrWhiteSpace(sheetName))
        {
            error = new ExcelError();
            error.Code = ExcelErrorCode.ExcelSheetNameIsNull;
            error.Msg = fileName;
            return false;
        }

        try
        {
            // Create a spreadsheet document by supplying the filepath.
            // By default, AutoSave = true, Editable = true, and Type = xlsx.
            SpreadsheetDocument spreadsheetDocument = SpreadsheetDocument.Create(fileName, SpreadsheetDocumentType.Workbook);

            // add file properties
            spreadsheetDocument.AddExtendedFilePropertiesPart();
            spreadsheetDocument.ExtendedFilePropertiesPart.Properties = new DocumentFormat.OpenXml.ExtendedProperties.Properties();

            // todo:
            spreadsheetDocument.ExtendedFilePropertiesPart.Properties.Company = new DocumentFormat.OpenXml.ExtendedProperties.Company("My Company");
            spreadsheetDocument.PackageProperties.Creator = "Me";


            // Add a WorkbookPart to the document.
            WorkbookPart workbookpart = spreadsheetDocument.AddWorkbookPart();
            workbookpart.Workbook = new Workbook();
            workbookpart.Workbook.Save();

            // Add a WorksheetPart to the WorkbookPart.
            WorksheetPart worksheetPart = workbookpart.AddNewPart<WorksheetPart>();
            worksheetPart.Worksheet = new Worksheet(new SheetData());
            worksheetPart.Worksheet.Save();

            // Shared string table
            SharedStringTablePart sharedStringTablePart = spreadsheetDocument.WorkbookPart.AddNewPart<SharedStringTablePart>();
            sharedStringTablePart.SharedStringTable = new SharedStringTable();
            sharedStringTablePart.SharedStringTable.Save();

            // Add Sheets to the Workbook.
            Sheets sheets = spreadsheetDocument.WorkbookPart.Workbook.AppendChild<Sheets>(new Sheets());

            // Append the first worksheet and associate it with the workbook.
            Sheet sheet = new Sheet()
            {
                Id = spreadsheetDocument.WorkbookPart.
                GetIdOfPart(worksheetPart),
                SheetId = 1,
                Name = sheetName
            };
            sheets.Append(sheet);

            // Stylesheet
            WorkbookStylesPart workbookStylesPart = spreadsheetDocument.WorkbookPart.AddNewPart<WorkbookStylesPart>();
            workbookStylesPart.Stylesheet = ExcelStylesManager.CreateEmptyStylesheet();
            workbookStylesPart.Stylesheet.Save();

            workbookpart.Workbook.Save();
            spreadsheetDocument.Save();

            // get all uri-part
            Dictionary<string, OpenXmlPart> dictUriOpenXmlPart = BuildUriPartDictionary(spreadsheetDocument);

            // load cellFormat of the excel file, some cell format can be undefined (not yet implemented)
            ExcelCellStyles excelCellStyles;
            if (!ExcelStylesManager.GetStylesCellFormat(dictUriOpenXmlPart, out excelCellStyles, out error))
                return false;

            // create an excelDoc with all important data
            excelDoc = new ExcelWorkbook(fileName, spreadsheetDocument, dictUriOpenXmlPart, excelCellStyles);


            return true;
        }
        catch (Exception ex)
        {
            error = new ExcelError();
            error.Code = ExcelErrorCode.UnableToCreateExcelFile;
            error.Exception = ex;
            return false;
        }
    }

    public bool OpenExcelFile(string fileName, out ExcelWorkbook? excelDoc, out ExcelError? error)
    {
        return OpenExcelFile(fileName, out excelDoc, out error, true);
    }

    public bool OpenExcelFileReadOnly(string fileName, out ExcelWorkbook? excelDoc, out ExcelError? error)
    {
        return OpenExcelFile(fileName, out excelDoc, out error, false);
    }

    /// <summary>
    /// Open the excel file, get the spread sheet document.
    /// By default, open the file in read/write mode.
    /// </summary>
    /// <param name="importCharge"></param>
    /// <param name="pathName"></param>
    /// <param name="fileName"></param>
    /// <param name="spreadsheetDocument"></param>
    /// <returns></returns>
    public bool OpenExcelFile(string fileName, out ExcelWorkbook? excelDoc, out ExcelError? error, bool readWriteMode)
    {
        excelDoc = null;
        try
        {
            var spreadsheetDocument = SpreadsheetDocument.Open(fileName, readWriteMode);

            // configure 
            if (spreadsheetDocument.WorkbookPart.Workbook.CalculationProperties != null)
            {
                spreadsheetDocument.WorkbookPart.Workbook.CalculationProperties.ForceFullCalculation = true;
                spreadsheetDocument.WorkbookPart.Workbook.CalculationProperties.FullCalculationOnLoad = true;
            }

            // get all uri-part
            Dictionary<string, OpenXmlPart> dictUriOpenXmlPart = BuildUriPartDictionary(spreadsheetDocument);

            // load cellFormat of the excel file, some cell format can be undefined (not yet implemented)
            ExcelCellStyles excelCellStyles;
            if (!ExcelStylesManager.GetStylesCellFormat(dictUriOpenXmlPart, out excelCellStyles, out error))
            {
                error = new ExcelError();
                error.Code = ExcelErrorCode.UnableToOpenExcelFile;
                return false;
            }

            // create an excelDoc with all important data
            excelDoc = new ExcelWorkbook(fileName, spreadsheetDocument, dictUriOpenXmlPart, excelCellStyles);
            return true;
        }
        catch (Exception e)
        {
            error = new ExcelError();
            error.Code = ExcelErrorCode.UnableToOpenExcelFile;
            error.Exception = e;
            return false;
        }
    }

    public bool SaveExcelFile(ExcelWorkbook excelDoc, out ExcelError? error)
    {
        if (excelDoc == null)
        {
            error = new ExcelError();
            error.Code = ExcelErrorCode.UnableToCloseExcelFile;
            return false;
        }

        //if (excelDoc.Workbook == null)
        //{
        //    error = new ExcelError();
        //    error.Code = ExcelErrorCode.UnableToCloseExcelFile;
        //    return false;
        //}

        error = null;
        try
        {
            //excelDoc.Workbook.Save();
            error = null;
            return true;
        }
        catch (Exception ex)
        {
            error = new ExcelError();
            error.Code = ExcelErrorCode.UnableToCloseExcelFile;
            error.Exception = ex;
            return false;
        }

    }

    /// <summary>
    /// Close the opened excel file.
    /// </summary>
    /// <param name="spreadsheetDocument"></param>
    public bool CloseExcelFile(ExcelWorkbook excelDoc, out ExcelError? excelError)
    {
        excelError = null;
        if (excelDoc == null) return false;
        if (excelDoc.SpreadsheetDocument == null) return false;

        try
        {
            // fermer le fichier excel
            excelDoc.SpreadsheetDocument.Close();
            return true;
        }
        catch (Exception e)
        {
            excelError = new ExcelError();
            excelError.Code = ExcelErrorCode.UnableToCloseExcelFile;
            excelError.Exception = e;
            return false;
        }
    }

    #region Private methods.

    private bool CheckFileName(string fileName, bool creationAction, out ExcelError? error)
    {
        if (string.IsNullOrWhiteSpace(fileName))
        {
            error = ExcelError.Create(ExcelErrorCode.ExcelFileNameIsNull, fileName);
            return false;
        }

        // extract the path of the file name and check it
        string fileNameOnly = Path.GetFileName(fileName);

        if (string.IsNullOrWhiteSpace(fileNameOnly))
        {
            error = ExcelError.Create(ExcelErrorCode.ExcelFileNameIsNull, fileName);
            return false;
        }

        // the path should exists
        string pathOnly = Path.GetDirectoryName(fileName);

        if (!Directory.Exists(pathOnly))
        {
            error = ExcelError.Create(ExcelErrorCode.UnableToCreateExcelFile, fileName);
            return false;
        }

        // openFile, the file should exists
        if (!creationAction && !File.Exists(fileName))
        {
            error = ExcelError.Create(ExcelErrorCode.ExcelFileNotFound, fileName);
            return false;
        }

        // createFile, the file shouldn't exists
        if (creationAction && File.Exists(fileName))
        {
            error = ExcelError.Create(ExcelErrorCode.ExcelFileAlreadyExists, fileName);
            return false;
        }
        error = null;
        return true;
    }

    /// <summary>
    /// Build uri OpenXmlPart dictionnary.
    /// </summary>
    /// <param name="document"></param>
    /// <returns></returns>
    private static Dictionary<string, OpenXmlPart> BuildUriPartDictionary(SpreadsheetDocument document)
    {
        var uriPartDictionary = new Dictionary<string, OpenXmlPart>();
        var queue = new Queue<OpenXmlPartContainer>();
        queue.Enqueue(document);
        while (queue.Count > 0)
        {
            foreach (var part in queue.Dequeue().Parts.Where(part => !uriPartDictionary.Keys.Contains(part.OpenXmlPart.Uri.ToString())))
            {
                uriPartDictionary.Add(part.OpenXmlPart.Uri.ToString(), part.OpenXmlPart);
                queue.Enqueue(part.OpenXmlPart);
            }
        }
        return uriPartDictionary;
    }

    #endregion



}
