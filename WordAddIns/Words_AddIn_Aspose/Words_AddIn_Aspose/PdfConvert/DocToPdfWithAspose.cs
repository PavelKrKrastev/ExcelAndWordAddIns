using Word = Microsoft.Office.Interop.Word;
using WordAndExcelToPdf;

namespace Words_AddIn_Aspose.PdfConvert
{
    class DocToPdfWithAspose
    {

        private Word.Document _currentWordDocument = Globals.ThisAddIn.Application.ActiveDocument;

        public void convertToPdf()
        {
            ConvertToPdf DocToPdfAspose = new ConvertToPdf();

            DocToPdfAspose.OpenedDocumentName = _currentWordDocument.Name;
            DocToPdfAspose.OpenedDocumentPath = _currentWordDocument.Path;

            DocToPdfAspose.ConvertDocToPdfWithAspose();
            }
        }
    }

