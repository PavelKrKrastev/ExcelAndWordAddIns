using First_Word_AddIn.PdfConvert;
using Microsoft.Office.Tools.Ribbon;

namespace First_Word_AddIn
{
    public partial class ExportRibbon
    {
        private void ExportRibbon_Load(object sender, RibbonUIEventArgs e)
        {

        }

        private void btn_Save_Pdf(object sender, RibbonControlEventArgs e)
        {
            DocToPdf pdfConvert = new DocToPdf();
            pdfConvert.convertToPdf();
        }
    }
}
