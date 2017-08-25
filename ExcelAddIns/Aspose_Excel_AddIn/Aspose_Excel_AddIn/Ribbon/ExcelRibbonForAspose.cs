using Microsoft.Office.Tools.Ribbon;
using Aspose_Excel_AddIn.PdfConvert;

namespace Aspose_Excel_AddIn
{
    public partial class AsposeExcelRibbon
    {
        private void AsposeExcelRibbon_Load(object sender, RibbonUIEventArgs e)
        {

        }

        private void button1_Click(object sender, RibbonControlEventArgs e)
        {
            XlsxToPdfWithAspose convertToPdf = new XlsxToPdfWithAspose();
            convertToPdf.Convert();
        }
    }
}
