using Autodesk.Connectivity.WebServices;
using Autodesk.DataManagement.Client.Framework.Vault.Currency.Entities;
//Adobe Acrobat:
using NDde;
using NDde.Client;

using System;
using System.IO;
using System.Linq;
using System.Collections;
using System.Collections.Generic;
using System.Text.RegularExpressions;
using System.Windows.Forms;
using VaultBrowserSample;
using Excel = Microsoft.Office.Interop.Excel;
using ACW = Autodesk.Connectivity.WebServices;
using Framework = Autodesk.DataManagement.Client.Framework;
using Vault = Autodesk.DataManagement.Client.Framework.Vault;
using VDF = Autodesk.DataManagement.Client.Framework;
using Autodesk.Connectivity.WebServicesTools;
using Autodesk.DataManagement.Client.Framework.Vault.Currency.Properties;
using System.Diagnostics;
//Open XML:
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;

namespace QueryVault
{
	public partial class ThisAddIn
	{
		#region Member Variables
        public bool OldStatus { get; set; }
        public WebServiceManager ServiceManager { get; set; }
		public DirectoryInfo InventorProjectRootFolder = null;
		public Vault.Currency.Connections.Connection m_conn = null;
		//private PropDefInfo[] propDefInfos = null;
		//private Vault.Forms.Models.BrowseVaultNavigationModel m_model = null;
		public bool NoMatch = true;
		public bool pdf = false;
        public PropertyDefinition propDef;
        public PropertyDefinition myUDP_FeatureCount = null;
        public PropertyDefinition myUDP_OccurrenceCount = null;
        public PropertyDefinition myUDP_ParameterCount = null;
        public PropertyDefinition myUDP_ConstraintCount = null;
		private bool printSelect = true;

		private DdeClient m_client;

		private List<Framework.Forms.Controls.GridLayout> m_availableLayouts = new List<Framework.Forms.Controls.GridLayout>();
		//private List<ToolStripMenuItem> m_viewButtons = new List<ToolStripMenuItem>();

		//private Func<Vault.Currency.Entities.IEntity, bool> m_filterCanDisplayEntity;
		private PropDef[] defs = null;
		public ListBoxFileItem selectedfile;
		public List<ListBoxFileItem> FoundList;
		#endregion

        private void ThisAddIn_Startup(object sender, System.EventArgs e)
		{
		}

		private void ThisAddIn_Shutdown(object sender, System.EventArgs e)
		{
		}

		#region VSTO generated code

		/// <summary>
		/// Required method for Designer support - do not modify
		/// the contents of this method with the code editor.
		/// </summary>
        private void InternalStartup()
		{
			this.Startup += new System.EventHandler(ThisAddIn_Startup);
			this.Shutdown += new System.EventHandler(ThisAddIn_Shutdown);
		}

		#endregion

        /// <summary>
        /// 
        /// </summary>
        public void InitializeSearchFromExcel()
		{
            OldStatus = Application.DisplayStatusBar;
			m_conn = Vault.Forms.Library.Login(null);
			ServiceManager = m_conn.WebServiceManager;
			InventorProjectRootFolder = new DirectoryInfo(ServiceManager.DocumentService.GetRequiredWorkingFolderLocation()); //need this for later.
            if (m_conn != null)
				try
			{
				Excel.Workbook wb = Globals.ThisAddIn.Application.ActiveWorkbook;
				Excel._Worksheet ws = wb.ActiveSheet;
				if (!ws.Name.Contains("MODELLING"))
				{
					MessageBox.Show("Try switching to one the tabs labelled \"MODELLING\" and try again!");
					return;
				}

				//property definitions retrieved from Vault Settings -> Behaviours Tab -> Properties screen and the "System Name" column in the resulting table.
                PropDef materialPropDef = ServiceManager.PropertyService.GetPropertyDefinitionsByEntityClassId("FILE").First(propDef => propDef.SysName.Equals("Material"));
				PropDef titlePropDef = ServiceManager.PropertyService.GetPropertyDefinitionsByEntityClassId("FILE").First(propDef => propDef.SysName.Equals("Title"));
				PropDef revNumberPropDef = ServiceManager.PropertyService.GetPropertyDefinitionsByEntityClassId("FILE").First(propDef => propDef.SysName.Equals("RevNumber"));
				PropDef legacyDwgNumPropDef = ServiceManager.PropertyService.GetPropertyDefinitionsByEntityClassId("FILE").First(propDef => propDef.SysName.Equals("Subject"));
				//an FYI for possible future usage:
				//Comments sysname = Comments
				//Description sysname = Description
				//Part Number sysname = PartNumber
				//Project sysname = Project
				//(Vault) Revision sysname = Revision

                Globals.ThisAddIn.Application.ActiveSheet.UsedRange();

                Excel.Range range = ws.UsedRange;
                int usedCount = 0;
                for (int i = 3; i < range.Rows.Count; i++)
                {
                    Excel.Range r = range.Cells[i, 2];
                    string str = r.Value2 ; //!= null ? r.Value2.ToString() : "";
                    if (str != null)
                    {
                        usedCount++;
                    }
                }
				//create a new instance of the FoundList Object
                FoundList = new List<ListBoxFileItem>();
				for (int i = 3; i < range.Rows.Count; i++) //need to start at row 3 accounting for headers and such-like.
				{
					double percent = ((double)i / usedCount);
					UpdateStatusBar(percent,"Updating File Status From Vault... Please Wait");
					Excel.Range r = range.Cells[i, 2]; //column B
                    string str = r.Value2 != null ? r.Value2.ToString() : "";
					//string str = r.get_Value(Excel.XlRangeValueDataType.xlRangeValueDefault).ToString();
					//debugging
                    if (str != string.Empty)
					{
						//MessageBox.Show("Drawing number= " + str);
						Match match;
						if (pdf)
						{
							match = Regex.Match(str, @"(.*)(\w{2}-\d{5,}|\w{2}-\w\d{5,})(.*)", RegexOptions.IgnoreCase); //pdfs don't (normally) have -000 at the end of the filename
						}
						else // .ipt or .iam
						{
							match = Regex.Match(str, @"(.*)(\w{2}-\d{5,}-000|\w{2}-\w\d{5,}-000)(.*)", RegexOptions.IgnoreCase);
						}

						if (match.Success)
						{
							//search for the drawing number in the vault
                            if (match.Groups[2].Captures.Count > 0)
							{
								Console.WriteLine("Drawing number matches the ###### or A##### pattern : " + match.Groups[2].ToString());
								str = match.Groups[2].ToString();
							}

							//else if (match.Groups[5].Captures.Count > 0)
							//{
							//    Console.WriteLine("Drawing number matches the ###### or A##### pattern : " + match.Groups[5].ToString());
							//    str = match.Groups[5].ToString();
							//}

							Excel.Range rVaultedFileName;

							if (pdf)
							{
								rVaultedFileName = range.Cells[i, 1]; //pdf
							}
							else
							{
								rVaultedFileName = range.Cells[i, 3]; //Inventor
							}

							Excel.Range rState = range.Cells[i, 4];
							Excel.Range rRevision = range.Cells[i, 5];
							Excel.Range rFileType = range.Cells[i, 6];
							//skip Assemblies when we're looking for pdfs for the time being.
                            if (pdf)
							{
								if (rFileType.Value2 == "Assembly")
								{
									continue;
								}
							}

							Excel.Range rVaulted = range.Cells[i, 7];
							Excel.Range rVaultLocation = range.Cells[i, 10];
							Excel.Range rTitle = range.Cells[i, 11];
							Excel.Range rDrawingRevision = range.Cells[i, 12];
							Excel.Range rLegacyDrawingNumber = range.Cells[i, 13];
                            Excel.Range rConstraintCount = range.Cells[i, 14];
							Excel.Range rFeatureCount = range.Cells[i, 15];
							Excel.Range rParameterCount = range.Cells[i, 16];
							Excel.Range rOccurrenceCount = range.Cells[i, 17];
							Excel.Range rMaterial = range.Cells[i, 18];

							rVaultedFileName.Select();
							//rVaultLocation.Select(); //scroll with the active cell

                            string searchedAlready = rVaultedFileName.Value2 != null ? rVaultedFileName.Value2.ToString() : string.Empty;
							DoSearch(str, searchedAlready);

							if (NoMatch == false && selectedfile != null)
							{
								FoundList.Add(selectedfile);
								if (selectedfile.File.EntityName.EndsWith(".iam"))
								{
									if (selectedfile.File.EntityName.StartsWith("AS-"))
									{
										rFileType.Value2 = "Assembly";
										rMaterial.Value2 = "No Material Assigned or Required";
									}
									else if (selectedfile.File.EntityName.StartsWith("DT-"))
									{
										rFileType.Value2 = "Detail Assembly";
										rMaterial.Value2 = "No Material Assigned or Required";
									}
								}
								else if(selectedfile.File.EntityName.EndsWith(".ipt"))
								{
									rFileType.Value2 = "Part";
									//only bother with material for part files.
                                    PropInst materialPropInst = ServiceManager.PropertyService.GetProperties("FILE", new long[] { selectedfile.File.EntityIterationId }, new long[] { materialPropDef.Id }).First();
									if (materialPropInst.Val != null)
									{
										rMaterial.Value2 = materialPropInst.Val;
									}
									else
									{
										rMaterial.Value2 = "No Material Assigned or Required";
									}
								}

								//add/update some information about the file in the Excel spreadsheet.
								//storing the filename that was selected means we don't need to prompt the user to choose again.
                                rVaultedFileName.Value2 = selectedfile.File.EntityName.ToString();
								if (pdf)
								{
									rVaultedFileName.Hyperlinks.Add(rVaultedFileName, FindLocalPdf(InventorProjectRootFolder, rVaultedFileName.Value2), Type.Missing, Type.Missing, Type.Missing);
								}
								else
								{
									rState.Value2 = selectedfile.File.LifecycleInfo.StateName;
									rRevision.Value2 = selectedfile.File.RevisionInfo.RevisionLabel;
									#region Is Vaulted
									//change the font to Wingdings
                                    rVaulted.Font.Name = "Wingdings";
									//set the value to the right character to present a tick.
                                    rVaulted.Value2 = ((char)0xFC).ToString();
									//need to perform a check here for the vault we're signed into?
									//MessageBox.Show(m_conn.Vault.ToString());
									if (m_conn.Vault.ToString() == "Legacy Vault")
									{
                                        rVaultLocation.Value2 = selectedfile.Folder.FullName.ToString().Replace("/","\\").Replace("$","C:\\Legacy Vault Working Folder") + "\\" + selectedfile.File.EntityName;
										//rVaultLocation.Value2 = selectedfile.folder.FullName.ToString().Replace("/", "\\").Replace("$", "C:\\Legacy Vault Working Folder") + "\\" + selectedfile.File.EntityName;
                                    }
                                    else
                                    {
                                        rVaultLocation.Value2 = selectedfile.folder.FullName.ToString().Replace("/", "\\").Replace("$", "C:\\Vault Working Folder") + "\\" + selectedfile.File.EntityName;
                                    }
                                    //deals with pulling title, rev number & subject values from the vaulted parts.
                                    if (rTitle.Value2 == "" || rTitle.Value2 == null)
                                    {
                                        PropInst titlePropInst = ServiceManager.PropertyService.GetProperties("FILE", new long[] { selectedfile.File.EntityIterationId }, new long[] { titlePropDef.Id }).First();
                                        if (titlePropInst.Val != null)
                                        {
                                            rTitle.Value2 = titlePropInst.Val;
                                        }
                                        else
                                        {
                                            rTitle.Value2 = "No Title iProperty Found!";
                                        }
                                    }
                                    if (rDrawingRevision.Value2 == "" || rDrawingRevision.Value2 == null)
                                    {
                                        PropInst revNumberPropInst = ServiceManager.PropertyService.GetProperties("FILE", new long[] { selectedfile.File.EntityIterationId }, new long[] { revNumberPropDef.Id }).First();
                                        if (revNumberPropInst.Val != null)
                                        {
                                            rDrawingRevision.Value2 = revNumberPropInst.Val;
                                        }
                                        else
                                        {
                                            rDrawingRevision.Value2 = "No Rev Number iProperty Found!";
                                        }
                                    }
                                    if (rLegacyDrawingNumber.Value2 == "" || rLegacyDrawingNumber.Value2 == null)
                                    {
                                        PropInst legacyDrawingNumberPropInst = ServiceManager.PropertyService.GetProperties("FILE", new long[] { selectedfile.File.EntityIterationId }, new long[] { legacyDwgNumPropDef.Id }).First();
                                        if (legacyDrawingNumberPropInst.Val != null)
                                        {
                                            rLegacyDrawingNumber.Value2 = legacyDrawingNumberPropInst.Val;
                                        }
                                        else
                                        {
                                            rLegacyDrawingNumber.Value2 = "No Legacy Drawing Number (Subject) iProperty Found!";
                                        }
                                    }
                                    if (selectedfile.FeatureCount != null)
                                    {
                                        rFeatureCount.Value2 = Convert.ToInt32(selectedfile.FeatureCount);
                                    }
                                    else
                                    {
                                        rFeatureCount.Value2 = 0;
                                    }
                                    //rFeatureCount.Value2 = selectedfile.FeatureCount.ToString();
                                    if (selectedfile.ParameterCount != null)
                                    {
                                        rParameterCount.Value2 = Convert.ToInt32(selectedfile.ParameterCount);
                                    }
                                    else
                                    {
                                        rParameterCount.Value2 = 0;
                                    }
                                    if (selectedfile.OccurrenceCount != null)
                                    {
                                        rOccurrenceCount.Value2 = Convert.ToInt32(selectedfile.OccurrenceCount);
                                    }
                                    else
                                    {
                                        rOccurrenceCount.Value2 = 0;
                                    }
                                    if (selectedfile.ConstraintCount != null)
                                    {
                                        rConstraintCount.Value2 = Convert.ToInt32(selectedfile.ConstraintCount);
                                    }
                                    else
                                    {
                                        rConstraintCount.Value2 = 0;
                                    }
                                    #endregion
                                    //reset the NoMatch bool & selectedfile
                                }
                                
                                NoMatch = true;
                                selectedfile = null;
                            }
                            else if(NoMatch == true) //need to mark the "VAULTED" Column with a Cross
                            {
                                #region Is NOT Vaulted
                                rState.Value2 = "NA";
                                rRevision.Value2 = "NA";
                                //change the font to Wingdings
                                rVaulted.Font.Name = "Wingdings";
                                //set the value to the right character to present a tick.
                                rVaulted.Value2 = ((char)0xFB).ToString();
                                #endregion
                                //reset the NoMatch bool & selectedfile
                                NoMatch = true;
                                selectedfile = null;
                            }
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("The error was: " + ex.Message + "\n" + ex.StackTrace);
                throw;
            }
            //we need to be sure to release all our connections when the app closes
            Application.StatusBar = false;
            Application.DisplayStatusBar = OldStatus;
            Vault.Library.ConnectionManager.CloseAllConnections();
        }
		///Updates the statusbar with a percentage so we can see how far along we are.
        public void UpdateStatusBar(double percent, string Message = "")
        {
            Application.StatusBar = Message + " (" + percent.ToString("P1") + ")";
        }
        
        /// <summary>
        /// 
        /// </summary>
        /// <param name="p"></param>
        /// <param name="VaultedFileName"></param>
        private void DoSearch(string p, string VaultedFileName)
        {
            SrchCond searchCondition = new SrchCond();
            searchCondition.PropDefId = 9; //filename
            bool searchedAlready = false;
            if (VaultedFileName != string.Empty)
            {
                searchedAlready = true;
                searchCondition.PropTyp = PropertySearchType.SingleProperty;
                searchCondition.SrchOper = 3; //equals
                searchCondition.SrchTxt = VaultedFileName; //our filename
            }
            else
            {
                searchCondition.PropTyp = PropertySearchType.SingleProperty;
                searchCondition.SrchOper = 1; //contains
                searchCondition.SrchTxt = p; //our drawing number
            }

            SrchCond[] conditions = new SrchCond[1];

            conditions[0] = searchCondition;

            string bookmark = string.Empty;
            SrchStatus status = null;

            //search for files
            List<Autodesk.Connectivity.WebServices.File> fileList = new List<Autodesk.Connectivity.WebServices.File>();
           
            if (!pdf)
            {
                PropertyDefinitionDictionary props =
                m_conn.PropertyManager.GetPropertyDefinitions(VDF.Vault.Currency.Entities.EntityClassIds.Files, null, PropertyDefinitionFilter.IncludeAll);
                
                foreach (var myKeyValuePair in props)
                {
                    propDef = myKeyValuePair.Value;
                    switch (propDef.DisplayName)
                    {
                        case "FeatureCount":
                            myUDP_FeatureCount = propDef;
                            break;
                        case "OccurrenceCount":
                            myUDP_OccurrenceCount = propDef;
                            break;
                        case "ParameterCount":
                            myUDP_ParameterCount = propDef;
                            break;
                        case "ConstraintCount":
                            myUDP_ConstraintCount = propDef;
                            break;
                        default:
                            break;
                    }
                }
            }
            
 
            while (status == null || fileList.Count < status.TotalHits)
            {
                Autodesk.Connectivity.WebServices.File[] files = m_conn.WebServiceManager.DocumentService.FindFilesBySearchConditions(
                    conditions, null, null, true, true,
                    ref bookmark, out status);

                if (files != null)
                    fileList.AddRange(files);
            }
            if (fileList.Count > 0)
            {
                /*we have a match in Vault for our drawing number 
                 now we need to decide whether we completely automate the next step or allow the user to select the expected file.
                 */
                /*ideally, if we search the fileList object we should find at least one .ipt or .iam if we find neither, we can safely assume the file hasn't been vaulted yet! */
                
                if (searchedAlready)
                {
                    foreach (Autodesk.Connectivity.WebServices.File file in fileList)
                    {
                        ListBoxFileItem fileItem = new ListBoxFileItem(new VDF.Vault.Currency.Entities.FileIteration(m_conn, file));
                        //long folderId = VDF.Vault.Currency.Entities
                        Autodesk.Connectivity.WebServices.Folder folder = m_conn.WebServiceManager.DocumentService.GetFolderById(file.FolderId);
                        fileItem.folder = folder;
                        //FilePath[] filepaths = m_conn.WebServiceManager.DocumentService.FindFilePathsByNameAndChecksum(file.Name, file.Cksum);
                        //String iptiampath = filepaths[0].Path.Replace("$", "C:\\Vault Working Folder");
                        //fileItem.Folder = iptiampath.Replace("/","\\");
                        //User Defined Properties - add more as necessary
                        if (!pdf)
                        {
                            fileItem.FeatureCount = m_conn.PropertyManager.GetPropertyValue(fileItem.File, myUDP_FeatureCount, null);
                            fileItem.OccurrenceCount = m_conn.PropertyManager.GetPropertyValue(fileItem.File, myUDP_OccurrenceCount, null);
                            fileItem.ParameterCount = m_conn.PropertyManager.GetPropertyValue(fileItem.File, myUDP_ParameterCount, null);
                            fileItem.ConstraintCount = m_conn.PropertyManager.GetPropertyValue(fileItem.File, myUDP_ConstraintCount, null);
                        }
                        Globals.ThisAddIn.NoMatch = false;
                        Globals.ThisAddIn.selectedfile = fileItem;
                        //we should only be returning one file because we're searching for it directly!
                        return;
                    }
                }
                else
                {
                    List<Autodesk.Connectivity.WebServices.File> results;
                    if (!Globals.ThisAddIn.pdf)
                    {
                        fileList.RemoveAll(StartsReplaceWith); //bin those files which start "Replace With"
                        fileList.RemoveAll(ContainsUnderscores);
                        results = fileList.FindAll(x => x.Name.EndsWith(".ipt") || x.Name.EndsWith(".iam")); //ignore .pdf
                    }
                    else
                    {
                        results = fileList.FindAll(x => x.Name.EndsWith(".pdf")); //specifically .pdf
                    }
                    
                    if (results.Count > 1)
                    {
                        foreach (Autodesk.Connectivity.WebServices.File file in results)
                        {
                            ListBoxFileItem findfile = new ListBoxFileItem(new VDF.Vault.Currency.Entities.FileIteration(m_conn, file));
                            if (Globals.ThisAddIn.FoundList.Contains(findfile))
                            {
                                Globals.ThisAddIn.NoMatch = false;
                                Globals.ThisAddIn.selectedfile = findfile;
                                return;
                            }
                        }
                        //we have multiple to choose from and should get the user to pick
                        //create a form object to display found items
                        string selectedfilename = string.Empty;
                        FileSelectionForm fileForm = new FileSelectionForm(m_conn);
                        //iterate through found files and display them in the search results list box
                        foreach (Autodesk.Connectivity.WebServices.File file in results)
                        {
                            ListBoxFileItem fileItem = new ListBoxFileItem(new VDF.Vault.Currency.Entities.FileIteration(m_conn, file));
                            fileForm.m_searchResultsListBox.Items.Add(fileItem);
                        }
                        //update the items count label
                        fileForm.m_SearchingForLabel.Text = "Searching for filename(s) containing: " + p;
                        fileForm.m_itemsCountLabel.Text = (results.Count > 0) ? results.Count + " Items" : "0 Items";
                        //display the form and wait for it to close using the ShowDialog() method.
                        fileForm.ShowDialog();

                    }
                    else if (results.Count == 1)
                    {
                        //get the first and only file we found whose name contains .iam or .ipt
                        Autodesk.Connectivity.WebServices.File foundfile = results[0];
                        ListBoxFileItem fileItem = new ListBoxFileItem(new VDF.Vault.Currency.Entities.FileIteration(m_conn, foundfile));
                        Autodesk.Connectivity.WebServices.Folder folder = m_conn.WebServiceManager.DocumentService.GetFolderById(foundfile.FolderId);
                        fileItem.folder = folder;
                        //User Defined Properties - add more as necessary
                        if (!pdf)
                        {
                            fileItem.FeatureCount = m_conn.PropertyManager.GetPropertyValue(fileItem.File, myUDP_FeatureCount, null);
                            fileItem.OccurrenceCount = m_conn.PropertyManager.GetPropertyValue(fileItem.File, myUDP_OccurrenceCount, null);
                            fileItem.ParameterCount = m_conn.PropertyManager.GetPropertyValue(fileItem.File, myUDP_ParameterCount, null);
                            fileItem.ConstraintCount = m_conn.PropertyManager.GetPropertyValue(fileItem.File, myUDP_ConstraintCount, null);
                        }
                        Globals.ThisAddIn.NoMatch = false;
                        Globals.ThisAddIn.selectedfile = fileItem;
                        return;
                                            }
                    else //no matches or all .pdf files
                    {
                        Globals.ThisAddIn.NoMatch = true;
                    }
                }
                
            }
        }

        private bool ContainsUnderscores(ACW.File obj)
        {
            return obj.Name.ToLower().Contains("_");
        }

        private bool StartsReplaceWith(ACW.File obj)
        {
            return obj.Name.ToLower().StartsWith("replace with");
        }
        private void InitializePropertyDefs()
        {
            defs = m_conn.WebServiceManager.PropertyService.GetPropertyDefinitionsByEntityClassId(VDF.Vault.Currency.Entities.EntityClassIds.Files);
            if (defs != null && defs.Length > 0)
            {
                Array.Sort(defs, new PropertyDefinitionSorter());
            }
        }
        /// <summary>
        /// Searches the local directory looking for pdf files that match (or nearly match) .ipt files.
        /// Could be made to look for .iam files as well but there are likely so many that the result would be an array of files.
        /// </summary>
        public string FindLocalPdf(DirectoryInfo dir,string pdfName)
        {
            try
            {
                var thispdf = (from file in dir.GetFiles("*.pdf", SearchOption.AllDirectories)
                              where file.Name == pdfName
                              select file).First();
                //string thispdf = from string[] a in System.IO.Directory.GetFiles(dir.FullName,"*.pdf",SearchOption.AllDirectories)
                //                    where a.Equals(pdfName)
                //                    select a;
                if (thispdf != null)
                {
                    return thispdf.DirectoryName + "\\" + thispdf.Name;
                }
                else
                {
                    return "";
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("Exception was: " + ex.Message + "\n\r" + ex.StackTrace);
                return "";
            }
        }

        /// <summary>
        /// Allows the user to print a bunch of filtered urls in Column A of a spreadsheet. These urls can be populated by the "Find Vaulted pdf" command in this addin, or by manually input urls.
        /// </summary>
        public void PrintPDFs()
        {
            #region Print PDFs
            MessageBox.Show("Make sure to disable the \"Protected Mode at Startup\" Option in adobe before continuing with this tool!", "Doh!", MessageBoxButtons.OK);
            MessageBox.Show("If you are running Adobe Reader 11.x please talk to Alex before running this application as it will not work without attention!", "Doh!", MessageBoxButtons.OK);
            m_client = new DdeClient("AcroViewR11", "control");
            bool tryStart = false;
            bool connected = false;
            do
            {
                try
                {
                    m_client.Connect();
                    connected = true;
                    //listBox1.Items.Add("Connected to DDE");
                    //listBox1.Refresh();
                }
                catch (DdeException)
                {
                    //Start Acrobat
                    var proc = new Process();
                    proc.StartInfo.FileName = "AcroRd32.exe";
                    proc.StartInfo.Arguments = "/n";
                    proc.Start();
                    proc.WaitForInputIdle();

                    tryStart = !tryStart;
                }
            } while (!connected && tryStart);
            //Excel.Application ExApp = Globals.ThisAddIn.Application as Microsoft.Office.Interop.Excel.Application;
            //Excel.Range SelectedRange = ExApp.Selection as Microsoft.Office.Interop.Excel.Range;
            Excel.Workbook wb = Globals.ThisAddIn.Application.ActiveWorkbook;
            Excel._Worksheet ws = wb.ActiveSheet;
            if (!ws.Name.Contains("MODELLING"))
            {
                MessageBox.Show("Try switching to one the tabs labelled \"MODELLING\" and try again!");
                return;
            }
            Excel.Range visibleCells = ws.UsedRange.SpecialCells(Excel.XlCellType.xlCellTypeVisible, Type.Missing);
            foreach (Excel.Range area in visibleCells.Areas)
            {
                foreach (Excel.Range row in area.Rows)
                {
                    foreach (Excel.Range column in row.Columns)
                    {
                        if (column.Hyperlinks.Count > 0)
                        {
                            column.Select();
                            string line = column.Hyperlinks.Item[1].Address;
                            line = wb.Path + "\\" + line;
                            if (line == null)
                                continue;
                            m_client.Execute("[DocOpen(\"" + line + "\")]", 60000);
                            if (printSelect)
                            {
                                m_client.Execute("[FilePrint(\"" + line + "\")]", 60000);
                                m_client.Execute("[DocClose(\"" + line + "\")]", 60000);
                            }
                            else
                            {
                                m_client.Execute("[FilePrintSilent(\"" + line + "\")]", 60000);
                                m_client.Execute("[DocClose(\"" + line + "\")]", 60000);
                            }
                        }
                        //only care about Column A so continue afterwards
                        continue;
                    }
                }
            }
            m_client.Execute("[AppExit]", 60000);
            #endregion
        }

        
    }
    #region "Search Condition Item Class"
    class SrchCondItem
    {
        public SrchCond SrchCond;
        public PropDef PropDef;

        public SrchCondItem(SrchCond srchCond, PropDef propDef)
        {
            this.SrchCond = srchCond;
            this.PropDef = propDef;
        }

        public override string ToString()
        {
            string conditionName = Condition.GetCondition(SrchCond.SrchOper).DisplayName;
            return String.Format("{0} {1} {2}", PropDef.DispName, conditionName, SrchCond.SrchTxt);
        }
    }
    #endregion
    #region PropertyDefinitionSorter Class
    /// <summary>
    /// Used for sorting collections of PropertyDefinition's.
    /// </summary>
    class PropertyDefinitionSorter : IComparer
    {
        /// <summary>
        /// Class (static) constructor that creates a static Comparer class instane used for sorting PropertyDefinition's.
        /// </summary>
        static PropertyDefinitionSorter()
        {

            m_comparer = new Comparer(Application.CurrentCulture);

        }

        private static Comparer m_comparer;

        public int Compare(object x, object y)
        {
            PropDef propDefX = x as PropDef;
            PropDef propDefY = y as PropDef;

            lock (m_comparer)
            {

                return m_comparer.Compare(propDefX.DispName, propDefY.DispName);

            }

        }

    }
    #endregion
    #region "ListBoxFileItem"
    /// <summary>
    /// A list box item which contains a File object
    /// </summary>
    public class ListBoxFileItem
    {
        private FileIteration file;
        public FileIteration File
        {
            get { return file; }
        }

        public ListBoxFileItem(FileIteration f)
        {
            file = f;
        }
        public Autodesk.Connectivity.WebServices.Folder folder;
        public Autodesk.Connectivity.WebServices.Folder Folder
        {
            get { return folder; }
        }
        /// <summary>
        /// Determines the text displayed in the ListBox
        /// </summary>
        public override string ToString()
        {
            return this.file.EntityName;
        }

        public object FeatureCount { get; set; }

        public object OccurrenceCount { get; set; }

        public object ParameterCount { get; set; }

        public object ConstraintCount { get; set; }

    }
    
    #endregion
#region "ExcelFile"
    public class ExcelFile
    {
        public string Name { get; set; }
        public string Folder { get; set; }
        public long FolderId { get; set; }
    }
#endregion
}
