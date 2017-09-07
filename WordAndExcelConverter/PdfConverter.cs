using System.IO;
using System.Windows.Forms;
using Cells = Aspose.Cells;
using Excel = Microsoft.Office.Interop.Excel;
using Word = Microsoft.Office.Interop.Word;
using Words = Aspose.Words;

namespace WordAndExcelConverter
{
    public class PdfConverter
    {
        #region Constants
        private const string _saveDialogFilter = "PDF|*.pdf";
        private const string _defaultPdfNameFormat = " PDF";
        #endregion Constants

        #region Private Properties
        private string _saveAsPdfPath { get; set; }
        private SaveFileDialog _saveDialog = new SaveFileDialog() { Filter = _saveDialogFilter };
        #endregion Private Properties

        #region Convert Excel Files To Pdf
        public void ConvertXlsxToPdf(Excel.Workbook CurrentExcelWorkBook)
        {
            GetDefaultPdfName(CurrentExcelWorkBook.Path, CurrentExcelWorkBook.Name);

            if (_saveDialog.ShowDialog() == DialogResult.OK)
            {
                _saveAsPdfPath = _saveDialog.FileName;
                CurrentExcelWorkBook.ExportAsFixedFormat(Excel.XlFixedFormatType.xlTypePDF, _saveAsPdfPath);
            }
        }

        public void ConvertXlsxToPdfWithAspose(string OpenedDocumentPath, string OpenedDocumentName)
        {
            GetDefaultPdfName(OpenedDocumentPath, OpenedDocumentName);

            if (_saveDialog.ShowDialog() == DialogResult.OK)
            {
                using (FileStream OpenedDocumentStream = new FileStream(OpenedDocumentPath + Path.DirectorySeparatorChar + OpenedDocumentName, FileMode.Open, FileAccess.Read, FileShare.ReadWrite))
                {
                    Cells.Workbook AsposeWorkbook = new Cells.Workbook(OpenedDocumentStream);
                    Cells.PdfSaveOptions SaveOptions = new Cells.PdfSaveOptions();
                    SaveOptions.AllColumnsInOnePagePerSheet = true;
                    _saveAsPdfPath = _saveDialog.FileName;
                    AsposeWorkbook.Save(_saveAsPdfPath, SaveOptions);
                }
            }
        }
        #endregion Convert Excel Files To Pdf

        #region Convert Word Files To Pdf
        public void ConvertDocxToPdf(Word.Document CurrentWordDocument)
        {
            GetDefaultPdfName(CurrentWordDocument.Path, CurrentWordDocument.Path);

            if (_saveDialog.ShowDialog() == DialogResult.OK)
            {
                _saveAsPdfPath = _saveDialog.FileName;
                CurrentWordDocument.ExportAsFixedFormat(_saveAsPdfPath, Word.WdExportFormat.wdExportFormatPDF);
            }
        }

        public void ConvertDocxToPdfWithAspose(string OpenedDocumentPath, string OpenedDocumentName)
        {
            GetDefaultPdfName(OpenedDocumentPath, OpenedDocumentName);

            if (_saveDialog.ShowDialog() == DialogResult.OK)
            {
                using (FileStream OpenedDocumentStream = new FileStream(OpenedDocumentPath + Path.DirectorySeparatorChar + OpenedDocumentName, FileMode.Open, FileAccess.Read, FileShare.ReadWrite))
                {
                    Words.Document WordDocument = new Words.Document(OpenedDocumentStream);
                    _saveAsPdfPath = _saveDialog.FileName;
                    WordDocument.Save(_saveAsPdfPath, Aspose.Words.SaveFormat.Pdf);
                }
            }
        }
        #endregion Convert Word Files To Pdf

        #region Pdf default name
        private void GetDefaultPdfName(string DocumentPath, string DocumentName)
        {
            _saveDialog.FileName = Path.GetFileNameWithoutExtension(DocumentPath + Path.DirectorySeparatorChar + DocumentName) + _defaultPdfNameFormat;
        }
        #endregion Pdf default name
    }
}
