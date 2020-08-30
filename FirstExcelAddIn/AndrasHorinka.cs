
namespace FirstExcelAddIn
{
    using Microsoft.Office.Tools.Ribbon;

    public partial class AndrasHorinka
    {
        private void AndrasHorinka_Load(object sender, RibbonUIEventArgs e)
        {

        }

        private void button1_Click(object sender, RibbonControlEventArgs e)
        {

            Globals.ThisAddIn.RetrieveDataFromMNB();
        }

        private void button2_Click(object sender, RibbonControlEventArgs e)
        {
            Globals.ThisAddIn.LogReasonForQuery();
        }
    }
}
