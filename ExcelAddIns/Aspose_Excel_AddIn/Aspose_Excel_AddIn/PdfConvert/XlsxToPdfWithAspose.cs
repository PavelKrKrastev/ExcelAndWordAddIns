using Excel = Microsoft.Office.Interop.Excel;
using WordAndExcelToPdf;

namespace Aspose_Excel_AddIn.PdfConvert
{
    public class XlsxToPdfWithAspose
    {
        private Excel.Workbook _currentWorkBook = Globals.ThisAddIn.Application.ActiveWorkbook;

        public void Convert()
        {
            ConvertToPdf XlsxToPdfAspose = new ConvertToPdf();
            XlsxToPdfAspose.ConvertXlsxToPdfWithAspose(_currentWorkBook.Path, _currentWorkBook.Name);
        }
    }
}
