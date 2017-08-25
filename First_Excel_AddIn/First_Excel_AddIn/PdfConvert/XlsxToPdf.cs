﻿using Excel = Microsoft.Office.Interop.Excel;
using WordAndExcelToPdf;

namespace First_Excel_AddIn.PdfConvert
{
    public class XlsxToPdf
    {
        private Excel.Workbook _currentWorkBook = Globals.ThisAddIn.Application.ActiveWorkbook;

        public void Convert()
        {
            ConvertToPdf XlsxToPdf = new ConvertToPdf();

            XlsxToPdf.OpenedDocumentName = _currentWorkBook.Name;
            XlsxToPdf.OpenedDocumentPath = _currentWorkBook.Path;

            XlsxToPdf.ConvertXlsxToPdf(_currentWorkBook);
        }
    }
}