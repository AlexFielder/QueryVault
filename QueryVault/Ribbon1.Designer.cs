namespace QueryVault
{
    partial class Ribbon1 : Microsoft.Office.Tools.Ribbon.RibbonBase
    {
        /// <summary>
        /// Required designer variable.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        public Ribbon1()
            : base(Globals.Factory.GetRibbonFactory())
        {
            InitializeComponent();
        }

        /// <summary> 
        /// Clean up any resources being used.
        /// </summary>
        /// <param name="disposing">true if managed resources should be disposed; otherwise, false.</param>
        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
            {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        #region Component Designer generated code

        /// <summary>
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        private void InitializeComponent()
        {
            this.tab1 = this.Factory.CreateRibbonTab();
            this.group1 = this.Factory.CreateRibbonGroup();
            this.QueryVault = this.Factory.CreateRibbonButton();
            this.group2 = this.Factory.CreateRibbonGroup();
            this.FindVaultedPdf = this.Factory.CreateRibbonButton();
            this.PrintSelectedPdfs = this.Factory.CreateRibbonButton();
            this.tab1.SuspendLayout();
            this.group1.SuspendLayout();
            this.group2.SuspendLayout();
            // 
            // tab1
            // 
            this.tab1.ControlId.ControlIdType = Microsoft.Office.Tools.Ribbon.RibbonControlIdType.Office;
            this.tab1.Groups.Add(this.group1);
            this.tab1.Groups.Add(this.group2);
            this.tab1.Label = "TabAddIns";
            this.tab1.Name = "tab1";
            // 
            // group1
            // 
            this.group1.Items.Add(this.QueryVault);
            this.group1.Label = "Vault Query";
            this.group1.Name = "group1";
            // 
            // QueryVault
            // 
            this.QueryVault.Label = "Run Vault Query";
            this.QueryVault.Name = "QueryVault";
            this.QueryVault.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.QueryVault_Click);
            // 
            // group2
            // 
            this.group2.Items.Add(this.FindVaultedPdf);
            this.group2.Items.Add(this.PrintSelectedPdfs);
            this.group2.Label = "PDF Functions";
            this.group2.Name = "group2";
            // 
            // FindVaultedPdf
            // 
            this.FindVaultedPdf.Enabled = false;
            this.FindVaultedPdf.Label = "Find Vaulted pdf";
            this.FindVaultedPdf.Name = "FindVaultedPdf";
            this.FindVaultedPdf.ScreenTip = "...For each .ipt file in our spreadsheet and add a link back to it from the sprea" +
    "dsheet";
            this.FindVaultedPdf.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.FindVaultedPdf_Click);
            // 
            // PrintSelectedPdfs
            // 
            this.PrintSelectedPdfs.Enabled = false;
            this.PrintSelectedPdfs.Label = "Print Selected PDFs";
            this.PrintSelectedPdfs.Name = "PrintSelectedPdfs";
            this.PrintSelectedPdfs.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.PrintSelectedPdfs_Click);
            // 
            // Ribbon1
            // 
            this.Name = "Ribbon1";
            this.RibbonType = "Microsoft.Excel.Workbook";
            this.Tabs.Add(this.tab1);
            this.Load += new Microsoft.Office.Tools.Ribbon.RibbonUIEventHandler(this.Ribbon1_Load);
            this.tab1.ResumeLayout(false);
            this.tab1.PerformLayout();
            this.group1.ResumeLayout(false);
            this.group1.PerformLayout();
            this.group2.ResumeLayout(false);
            this.group2.PerformLayout();

        }

        #endregion

        internal Microsoft.Office.Tools.Ribbon.RibbonTab tab1;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup group1;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton QueryVault;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton FindVaultedPdf;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton PrintSelectedPdfs;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup group2;
    }

    partial class ThisRibbonCollection
    {
        internal Ribbon1 Ribbon1
        {
            get { return this.GetRibbon<Ribbon1>(); }
        }
    }
}
