﻿using Microsoft.Office.Interop.Excel;
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
        #region Private Properties
        private string _saveAsPdfPath { get; set; }
        private SaveFileDialog _saveDialog = new SaveFileDialog() { Filter = "PDF|*.pdf" };
        #endregion Private Properties

        #region Convert Excel Files To Pdf
        public void ConvertXlsxToPdf(Excel.Workbook CurrentWorkBook)
        {
            if (_saveDialog.ShowDialog() == DialogResult.OK)
            {
                _saveAsPdfPath = _saveDialog.FileName;
                CurrentWorkBook.ExportAsFixedFormat(XlFixedFormatType.xlTypePDF, _saveAsPdfPath);
            }
        }

        public void ConvertXlsxToPdfWithAspose(string OpenedDocumentPath, string OpenedDocumentName)
        {
            if (_saveDialog.ShowDialog() == DialogResult.OK)
            {
                using (FileStream docStream = new FileStream(OpenedDocumentPath + "\\" + OpenedDocumentName, FileMode.Open, FileAccess.Read, FileShare.ReadWrite))
                {
                    Cells.Workbook AsposeWorkbook = new Cells.Workbook(docStream);
                    Cells.PdfSaveOptions SaveOptions = new Cells.PdfSaveOptions(Aspose.Cells.SaveFormat.Pdf);
                    SaveOptions.AllColumnsInOnePagePerSheet = true;
                    _saveAsPdfPath = _saveDialog.FileName;
                    AsposeWorkbook.Save(_saveAsPdfPath, SaveOptions);
                }
            }
        }
        #endregion Convert Excel Files To Pdf

        #region Convert Word Files To Pdf
        public void ConvertDocToPdf(Word.Document CurrentWordDoc)
        {
            if (_saveDialog.ShowDialog() == DialogResult.OK)
            {
                _saveAsPdfPath = _saveDialog.FileName;
                CurrentWordDoc.ExportAsFixedFormat(_saveAsPdfPath, Word.WdExportFormat.wdExportFormatPDF);
            }
        }

        public void ConvertDocToPdfWithAspose(string OpenedDocumentPath, string OpenedDocumentName)
        {
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
        #endregion Convert Word Files To Pdf
    }
}