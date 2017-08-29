using Microsoft.Office.Tools.Ribbon;
using Excel = Microsoft.Office.Interop.Excel;
using WordAndExcelToPdf;


namespace Aspose_Excel_AddIn
{
    public partial class AsposeExcelRibbon
    {
        private void AsposeExcelRibbon_Load(object sender, RibbonUIEventArgs e)
        {

        }

        private void button1_Click(object sender, RibbonControlEventArgs e)
        {
            Excel.Workbook _currentWorkBook = Globals.ThisAddIn.Application.ActiveWorkbook;
            PdfConverter XlsxToPdfAspose = new PdfConverter();
            XlsxToPdfAspose.ConvertXlsxToPdfWithAspose(_currentWorkBook.Path, _currentWorkBook.Name);
        }
    }
}
