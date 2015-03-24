using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using VDF = Autodesk.DataManagement.Client.Framework;

namespace QueryVault
{
    public partial class FileSelectionForm : Form
    {
        private VDF.Vault.Currency.Connections.Connection m_connection;
        public string selectedfilename;
        public FileSelectionForm(VDF.Vault.Currency.Connections.Connection connection)
        {
            InitializeComponent();
            m_connection = connection;
        }
        /// <summary>
        /// Takes the selected file in the list and passes the relevant information about it back to our Excel document.
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void button1_Click(object sender, EventArgs e)
        {
            ListBoxFileItem selectedfileitem = (ListBoxFileItem)m_searchResultsListBox.SelectedItem;
            List<VDF.Vault.Currency.Entities.FileIteration> fileIterations = new List<VDF.Vault.Currency.Entities.FileIteration>();
            fileIterations.Add(selectedfileitem.File);
            IDictionary<long, VDF.Vault.Currency.Entities.Folder> folderIdsToFolderEntities = Globals.ThisAddIn.m_conn.FolderManager.GetFoldersByIds(fileIterations.Select(file => file.FolderId));
            selectedfileitem.folder = folderIdsToFolderEntities.Select(m => m).Where(kvp => kvp.Key == selectedfileitem.File.FolderId).Select(k => k.Value).First(); 
            //User Defined Properties - add more as necessary
            //if (!Globals.ThisAddIn.pdf)
            //{
            //    selectedfileitem.FeatureCount = Convert.ToInt32(Globals.ThisAddIn.m_conn.PropertyManager.GetPropertyValue(selectedfileitem.File, Globals.ThisAddIn.myUDP_FeatureCountPropDefinition, null));
            //    selectedfileitem.OccurrenceCount = Convert.ToInt32(Globals.ThisAddIn.m_conn.PropertyManager.GetPropertyValue(selectedfileitem.File, Globals.ThisAddIn.myUDP_OccurrenceCountPropDefinition, null));
            //    selectedfileitem.ParameterCount = Convert.ToInt32(Globals.ThisAddIn.m_conn.PropertyManager.GetPropertyValue(selectedfileitem.File, Globals.ThisAddIn.myUDP_ParameterCountPropDefinition, null));
            //}
            if (selectedfileitem !=null)
            {
                Globals.ThisAddIn.NoMatch = false;
                Globals.ThisAddIn.selectedfile = selectedfileitem;
                try
                {

                    if (Globals.ThisAddIn.FoundList != null)
                    {
                        if (!Globals.ThisAddIn.FoundList.Contains(selectedfileitem))
                        {
                            Globals.ThisAddIn.FoundList.Add(selectedfileitem);
                        }
                        else
                        {
                            int idx = Globals.ThisAddIn.FoundList.FindIndex((ListBoxFileItem f) => f.File == selectedfileitem.File);
                            Globals.ThisAddIn.selectedfile = Globals.ThisAddIn.FoundList[idx];
                        } 
                    }
                }
                catch (Exception)
                {

                }
                this.Close();
            }
            else
            {
                MessageBox.Show("You need to select a file!");
            }
        }

        private void button2_Click(object sender, EventArgs e)
        {
            Globals.ThisAddIn.NoMatch = true;
            this.Close();
        }
        private void m_openFileToolStripMenuItem2_Click(object sender, EventArgs e)
        {
            OpenFile();
        }
        private void OpenFile()
        {
            if (m_searchResultsListBox.SelectedItem != null)
            {
                ListBoxFileItem fileItem = (ListBoxFileItem)m_searchResultsListBox.SelectedItem;
                OpenFileCommand.Execute(fileItem.File, m_connection);
            }
        }

        private void m_searchResultsListBox_MouseDoubleClick(object sender, MouseEventArgs e)
        {
            OpenFile();
        }
    }
}
