namespace QueryVault
{
    partial class FileSelectionForm
    {
        /// <summary>
        /// Required designer variable.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

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

        #region Windows Form Designer generated code

        /// <summary>
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        private void InitializeComponent()
        {
            this.components = new System.ComponentModel.Container();
            this.m_searchResultsListBox = new System.Windows.Forms.ListBox();
            this.button1 = new System.Windows.Forms.Button();
            this.button2 = new System.Windows.Forms.Button();
            this.contextMenuStrip1 = new System.Windows.Forms.ContextMenuStrip(this.components);
            this.m_openFileToolStripMenuItem2 = new System.Windows.Forms.ToolStripMenuItem();
            this.m_itemsCountLabel = new System.Windows.Forms.Label();
            this.m_SearchingForLabel = new System.Windows.Forms.Label();
            this.contextMenuStrip1.SuspendLayout();
            this.SuspendLayout();
            // 
            // m_searchResultsListBox
            // 
            this.m_searchResultsListBox.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.m_searchResultsListBox.Location = new System.Drawing.Point(12, 38);
            this.m_searchResultsListBox.Name = "m_searchResultsListBox";
            this.m_searchResultsListBox.Size = new System.Drawing.Size(502, 147);
            this.m_searchResultsListBox.TabIndex = 12;
            // 
            // button1
            // 
            this.button1.Location = new System.Drawing.Point(439, 219);
            this.button1.Name = "button1";
            this.button1.Size = new System.Drawing.Size(75, 23);
            this.button1.TabIndex = 13;
            this.button1.Text = "Done!";
            this.button1.UseVisualStyleBackColor = true;
            this.button1.Click += new System.EventHandler(this.button1_Click);
            // 
            // button2
            // 
            this.button2.Location = new System.Drawing.Point(12, 219);
            this.button2.Name = "button2";
            this.button2.Size = new System.Drawing.Size(101, 23);
            this.button2.TabIndex = 14;
            this.button2.Text = "No Matching File!";
            this.button2.UseVisualStyleBackColor = true;
            this.button2.Click += new System.EventHandler(this.button2_Click);
            // 
            // contextMenuStrip1
            // 
            this.contextMenuStrip1.Items.AddRange(new System.Windows.Forms.ToolStripItem[] {
            this.m_openFileToolStripMenuItem2});
            this.contextMenuStrip1.Name = "contextMenuStrip1";
            this.contextMenuStrip1.Size = new System.Drawing.Size(125, 26);
            // 
            // m_openFileToolStripMenuItem2
            // 
            this.m_openFileToolStripMenuItem2.Name = "m_openFileToolStripMenuItem2";
            this.m_openFileToolStripMenuItem2.Size = new System.Drawing.Size(124, 22);
            this.m_openFileToolStripMenuItem2.Text = "Open File";
            // 
            // m_itemsCountLabel
            // 
            this.m_itemsCountLabel.AutoSize = true;
            this.m_itemsCountLabel.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.m_itemsCountLabel.Location = new System.Drawing.Point(12, 188);
            this.m_itemsCountLabel.Name = "m_itemsCountLabel";
            this.m_itemsCountLabel.Size = new System.Drawing.Size(43, 13);
            this.m_itemsCountLabel.TabIndex = 16;
            this.m_itemsCountLabel.Text = "0 Items";
            // 
            // m_SearchingForLabel
            // 
            this.m_SearchingForLabel.AutoSize = true;
            this.m_SearchingForLabel.Location = new System.Drawing.Point(15, 13);
            this.m_SearchingForLabel.Name = "m_SearchingForLabel";
            this.m_SearchingForLabel.Size = new System.Drawing.Size(76, 13);
            this.m_SearchingForLabel.TabIndex = 17;
            this.m_SearchingForLabel.Text = "Searching for: ";
            // 
            // FileSelectionForm
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(526, 254);
            this.Controls.Add(this.m_SearchingForLabel);
            this.Controls.Add(this.m_itemsCountLabel);
            this.Controls.Add(this.button2);
            this.Controls.Add(this.button1);
            this.Controls.Add(this.m_searchResultsListBox);
            this.Name = "FileSelectionForm";
            this.Text = "FileSelectionForm";
            this.contextMenuStrip1.ResumeLayout(false);
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        public System.Windows.Forms.ListBox m_searchResultsListBox;
        private System.Windows.Forms.Button button1;
        private System.Windows.Forms.Button button2;
        private System.Windows.Forms.ContextMenuStrip contextMenuStrip1;
        private System.Windows.Forms.ToolStripMenuItem m_openFileToolStripMenuItem2;
        public System.Windows.Forms.Label m_itemsCountLabel;
        public System.Windows.Forms.Label m_SearchingForLabel;
    }
}