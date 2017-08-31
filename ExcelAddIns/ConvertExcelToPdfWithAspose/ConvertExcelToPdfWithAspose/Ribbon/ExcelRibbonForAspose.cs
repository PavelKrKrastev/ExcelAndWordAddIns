using Microsoft.Office.Tools.Ribbon;
using Excel = Microsoft.Office.Interop.Excel;
using WordAndExcelConverter;


namespace ConvertExcelToPdfWithAspose
{
    public partial class AsposeConvertRibbon
    {
        private void AsposeExcelRibbon_Load(object sender, RibbonUIEventArgs e)
        {

        }

        private void button1_Click(object sender, RibbonControlEventArgs e)
        {
            Excel.Workbook currentWorkBook = Globals.ThisAddIn.Application.ActiveWorkbook;
            PdfConverter XlsxToPdfAspose = new PdfConverter();
            XlsxToPdfAspose.ConvertXlsxToPdfWithAspose(currentWorkBook.Path, currentWorkBook.Name);
        }
    }
}
