namespace Excelam
{
    public class ExcelApi
    {
        /// <summary>
        /// Constructor.
        /// </summary>
        public ExcelApi()
        {
            ExcelFileApi = new ExcelFileApi();
            ExcelSheetApi = new ExcelSheetApi();
            ExcelCellValueApi = new ExcelCellValueApi();
        }
        public ExcelFileApi ExcelFileApi { get; private set; }
        public ExcelSheetApi ExcelSheetApi { get; private set; }

        /// <summary>
        /// To access cell value.
        /// </summary>
        public ExcelCellValueApi ExcelCellValueApi { get; private set; }
    }
}