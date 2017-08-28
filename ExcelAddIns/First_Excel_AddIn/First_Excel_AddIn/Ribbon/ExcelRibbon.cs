using Microsoft.Office.Tools.Ribbon;
using Excel = Microsoft.Office.Interop.Excel;
using WordAndExcelToPdf;

namespace First_Excel_AddIn
{
    public partial class ConverRibbon
    {
        private void ConverRibbon_Load(object sender, RibbonUIEventArgs e)
        {

        }

        private void btnConvert_Click(object sender, RibbonControlEventArgs e)
        {
            Excel.Workbook _currentWorkBook = Globals.ThisAddIn.Application.ActiveWorkbook;
            PdfCoverter XlsxToPdf = new PdfCoverter();
            XlsxToPdf.ConvertXlsxToPdf(_currentWorkBook);
        }
    }
}
