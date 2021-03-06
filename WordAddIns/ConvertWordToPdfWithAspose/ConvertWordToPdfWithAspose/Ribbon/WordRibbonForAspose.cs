﻿using Microsoft.Office.Tools.Ribbon;
using Word = Microsoft.Office.Interop.Word;
using WordAndExcelConverter;

namespace ConvertWordToPdfWithAspose
{
    public partial class AsposeConvertRibbon
    {
        private void AsposeWordRibbon_Load(object sender, RibbonUIEventArgs e)
        {

        }

        private void btnConvert_Click(object sender, RibbonControlEventArgs e)
        {
            Word.Document currentWordDocument = Globals.ThisAddIn.Application.ActiveDocument;
            PdfConverter DocToPdfAspose = new PdfConverter();
            DocToPdfAspose.ConvertDocxToPdfWithAspose(currentWordDocument.Path, currentWordDocument.Name);
        }
    }
}
