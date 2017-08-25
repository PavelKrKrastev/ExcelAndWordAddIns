using System.Windows.Forms;
using System.IO;
using Microsoft.Office.Interop.Excel;
using Excel = Microsoft.Office.Interop.Excel;
using Word = Microsoft.Office.Interop.Word;
using Aspose.Cells;
using Cells = Aspose.Cells;
using Words = Aspose.Words;

namespace WordAndExcelToPdf
{
    public class ConvertToPdf
    {
        #region Private properties
        private string _saveAsPdfPath { get; set; }
        private SaveFileDialog _saveDialog = new SaveFileDialog() { Filter = "PDF|*.pdf" };
        #endregion

        #region Public properties
        public string OpenedDocumentName { get; set; }
        public string OpenedDocumentPath { get; set; }
        #endregion

        #region Convert Excel Files To Pdf
        public void ConvertXlsxToPdf(Excel.Workbook CurrentWorkBook)
        {
            SetDefaultPdfName();

            if (_saveDialog.ShowDialog() == DialogResult.OK)
            {
                _saveAsPdfPath = _saveDialog.FileName;
                CurrentWorkBook.ExportAsFixedFormat(XlFixedFormatType.xlTypePDF, _saveAsPdfPath);
            }
        }

        public void ConvertXlsxToPdfWithAspose()
        {
            SetDefaultPdfName();

            if (_saveDialog.ShowDialog() == DialogResult.OK)
            {
                using (FileStream docStream = new FileStream(OpenedDocumentPath + "\\" + OpenedDocumentName, FileMode.Open, FileAccess.Read, FileShare.ReadWrite))
                {
                    Cells.Workbook AsposeWorkbook = new Cells.Workbook(docStream);
                    PdfSaveOptions SaveOptions = new PdfSaveOptions(Aspose.Cells.SaveFormat.Pdf);
                    SaveOptions.AllColumnsInOnePagePerSheet = true;
                    _saveAsPdfPath = _saveDialog.FileName;
                    AsposeWorkbook.Save(_saveAsPdfPath, SaveOptions);
                }
            }
        }
        #endregion

        #region Convert Word Files To Pdf
        public void ConvertDocToPdf(Word.Document CurrentWordDoc)
        {
            SetDefaultPdfName();

            if (_saveDialog.ShowDialog() == DialogResult.OK)
            {
                _saveAsPdfPath = _saveDialog.FileName;
                CurrentWordDoc.ExportAsFixedFormat(_saveAsPdfPath, Word.WdExportFormat.wdExportFormatPDF);
            }
        }

        public void ConvertDocToPdfWithAspose()
        {
            SetDefaultPdfName();

            if (_saveDialog.ShowDialog() == DialogResult.OK)
            {
                using (FileStream docStream = new FileStream(OpenedDocumentPath + "\\" + OpenedDocumentName, FileMode.Open, FileAccess.Read, FileShare.ReadWrite))
                {
                    Words.Document WordDocument = new Words.Document(docStream);
                    _saveAsPdfPath = _saveDialog.FileName;
                    WordDocument.Save(_saveAsPdfPath, Aspose.Words.SaveFormat.Pdf);
                }
            }
        }
        #endregion

        private void SetDefaultPdfName()
        {
            _saveDialog.FileName = Path.GetFileNameWithoutExtension(OpenedDocumentPath + "\\" + OpenedDocumentName) + "_PDF";
        }
    }
}
