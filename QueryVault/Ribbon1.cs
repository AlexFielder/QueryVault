using Microsoft.Office.Tools.Ribbon;
using System.Windows.Forms;

//using Microsoft.Office.Tools.Excel;


namespace QueryVault
{
    public partial class Ribbon1
    {

        private void Ribbon1_Load(object sender, RibbonUIEventArgs e)
        {

        }

        private void QueryVault_Click(object sender, RibbonControlEventArgs e)
        {
            if (Globals.ThisAddIn.Application.ActiveWorkbook.Name.Contains("Project Tracker.xlsx"))
            {
                Globals.ThisAddIn.pdf = false;
                Globals.ThisAddIn.InitializeSearchFromExcel();
            }
            else
            {
                MessageBox.Show("This application will only work in the Project Tracker.xlsx file!");
                return;
            }
            
        }

        private void FindVaultedPdf_Click(object sender, RibbonControlEventArgs e)
        {
            if (Globals.ThisAddIn.Application.ActiveWorkbook.Name == "Project Tracker.xlsx")
            {
                Globals.ThisAddIn.pdf = true;
                Globals.ThisAddIn.InitializeSearchFromExcel();
                //Globals.ThisAddIn.FindLocalPdf(Globals.ThisAddIn.InventorProjectRootFolder);
            }
            else
            {
                MessageBox.Show("This application will only work in the Project Tracker.xlsx file!");
                return;
            }
        }

        private void PrintSelectedPdfs_Click(object sender, RibbonControlEventArgs e)
        {
            if (Globals.ThisAddIn.Application.ActiveWorkbook.Name == "Project Tracker.xlsx")
            {
                Globals.ThisAddIn.PrintPDFs();
            }
        }
    }
}
