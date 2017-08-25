using Microsoft.Office.Tools.Ribbon;
using Words_AddIn_Aspose.PdfConvert;

namespace Words_AddIn_Aspose
{
    public partial class AsposeWordRibbon
    {
        private void AsposeWordRibbon_Load(object sender, RibbonUIEventArgs e)
        {

        }

        private void btnConvert_Click(object sender, RibbonControlEventArgs e)
        {
            DocToPdfWithAspose pdfConvert = new DocToPdfWithAspose();
            pdfConvert.convertToPdf();
        }
    }
}
