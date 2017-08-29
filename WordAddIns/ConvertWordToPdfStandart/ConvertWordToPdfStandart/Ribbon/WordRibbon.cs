using Word = Microsoft.Office.Interop.Word;
using WordAndExcelConverter;
using Microsoft.Office.Tools.Ribbon;

namespace ConvertWordToPdfStandart
{
    public partial class ConvertRibbon
    {
        private void ExportRibbon_Load(object sender, RibbonUIEventArgs e)
        {

        }

        private void btn_Save_Pdf(object sender, RibbonControlEventArgs e)
        {
            Word.Document _currentWordDocument = Globals.ThisAddIn.Application.ActiveDocument;
            PdfConverter DocToPdf = new PdfConverter();
            DocToPdf.ConvertDocToPdf(_currentWordDocument);
        }
    }
}
