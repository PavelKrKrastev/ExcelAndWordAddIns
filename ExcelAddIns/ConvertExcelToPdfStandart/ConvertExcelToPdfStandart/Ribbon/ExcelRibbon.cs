using Microsoft.Office.Tools.Ribbon;
using Excel = Microsoft.Office.Interop.Excel;
using WordAndExcelConverter;

namespace ConvertExcelToPdfStandart
{
    public partial class ConvertRibbon
    {
        private void ConverRibbon_Load(object sender, RibbonUIEventArgs e)
        {

        }

        private void btnConvert_Click(object sender, RibbonControlEventArgs e)
        {
            Excel.Workbook _currentWorkBook = Globals.ThisAddIn.Application.ActiveWorkbook;
            PdfConverter XlsxToPdf = new PdfConverter();
            XlsxToPdf.ConvertXlsxToPdf(_currentWorkBook);
        }
    }
}
