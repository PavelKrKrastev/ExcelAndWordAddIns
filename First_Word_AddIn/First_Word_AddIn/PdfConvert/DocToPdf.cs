﻿using Word = Microsoft.Office.Interop.Word;
using WordAndExcelToPdf;

namespace First_Word_AddIn.PdfConvert
{
    class DocToPdf
    {
        private Word.Document _currentWordDocument = Globals.ThisAddIn.Application.ActiveDocument;

        public void convertToPdf()
        {
            ConvertToPdf DocToPdf = new ConvertToPdf();

            DocToPdf.OpenedDocumentName = _currentWordDocument.Name;
            DocToPdf.OpenedDocumentPath = _currentWordDocument.Path;

            DocToPdf.ConvertDocToPdf(_currentWordDocument);
        }
    }
}
