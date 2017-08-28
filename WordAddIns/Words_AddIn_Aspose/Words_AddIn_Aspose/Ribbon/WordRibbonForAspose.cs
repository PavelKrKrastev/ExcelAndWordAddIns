using Microsoft.Office.Tools.Ribbon;
using Word = Microsoft.Office.Interop.Word;
using WordAndExcelToPdf;

namespace Words_AddIn_Aspose
{
    public partial class AsposeWordRibbon
    {
        private void AsposeWordRibbon_Load(object sender, RibbonUIEventArgs e)
        {

        }

        private void btnConvert_Click(object sender, RibbonControlEventArgs e)
        {
            Word.Document _currentWordDocument = Globals.ThisAddIn.Application.ActiveDocument;
            PdfCoverter DocToPdfAspose = new PdfCoverter();
            DocToPdfAspose.ConvertDocToPdfWithAspose(_currentWordDocument.Path, _currentWordDocument.Name);
        }
    }
}
