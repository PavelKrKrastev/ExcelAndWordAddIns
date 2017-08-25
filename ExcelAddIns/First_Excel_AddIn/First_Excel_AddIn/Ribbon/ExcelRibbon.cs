using Microsoft.Office.Tools.Ribbon;
using First_Excel_AddIn.PdfConvert; 

namespace First_Excel_AddIn
{
    public partial class ConverRibbon
    {
        private void ConverRibbon_Load(object sender, RibbonUIEventArgs e)
        {

        }

        private void btnConvert_Click(object sender, RibbonControlEventArgs e)
        {
            XlsxToPdf pdfConvert = new XlsxToPdf();
            pdfConvert.Convert();
        }
    }
}
