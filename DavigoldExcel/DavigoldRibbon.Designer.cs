namespace DavigoldExcel
{
    partial class DavigoldRibbon : Microsoft.Office.Tools.Ribbon.RibbonBase
    {
        /// <summary>
        /// Required designer variable.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        public DavigoldRibbon()
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
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(DavigoldRibbon));
            this.tab1 = this.Factory.CreateRibbonTab();
            this.group1 = this.Factory.CreateRibbonGroup();
            this.ImportButton = this.Factory.CreateRibbonButton();
            this.group2 = this.Factory.CreateRibbonGroup();
            this.VersionLabel = this.Factory.CreateRibbonLabel();
            this.UpdatedOnLabel = this.Factory.CreateRibbonLabel();
            this.SyncButton = this.Factory.CreateRibbonButton();
            this.group3 = this.Factory.CreateRibbonGroup();
            this.ShowHideButton = this.Factory.CreateRibbonButton();
            this.UploadButton = this.Factory.CreateRibbonButton();
            this.LogoutButton = this.Factory.CreateRibbonButton();
            this.StatusLabel = this.Factory.CreateRibbonLabel();
            this.tab1.SuspendLayout();
            this.group1.SuspendLayout();
            this.group2.SuspendLayout();
            this.group3.SuspendLayout();
            this.SuspendLayout();
            // 
            // tab1
            // 
            this.tab1.ControlId.ControlIdType = Microsoft.Office.Tools.Ribbon.RibbonControlIdType.Office;
            this.tab1.Groups.Add(this.group1);
            this.tab1.Groups.Add(this.group2);
            this.tab1.Groups.Add(this.group3);
            this.tab1.Label = "ONE";
            this.tab1.Name = "tab1";
            // 
            // group1
            // 
            this.group1.Items.Add(this.ImportButton);
            this.group1.Name = "group1";
            // 
            // ImportButton
            // 
            this.ImportButton.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.ImportButton.Image = ((System.Drawing.Image)(resources.GetObject("ImportButton.Image")));
            this.ImportButton.Label = "Download data into Excel";
            this.ImportButton.Name = "ImportButton";
            this.ImportButton.ShowImage = true;
            this.ImportButton.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.ImportButton_Click);
            // 
            // group2
            // 
            this.group2.Items.Add(this.VersionLabel);
            this.group2.Items.Add(this.UpdatedOnLabel);
            this.group2.Items.Add(this.SyncButton);
            this.group2.Name = "group2";
            // 
            // VersionLabel
            // 
            this.VersionLabel.Label = "Version";
            this.VersionLabel.Name = "VersionLabel";
            // 
            // UpdatedOnLabel
            // 
            this.UpdatedOnLabel.Label = "Last Updated on 04 April 2023";
            this.UpdatedOnLabel.Name = "UpdatedOnLabel";
            // 
            // SyncButton
            // 
            this.SyncButton.Label = "Publish Version";
            this.SyncButton.Name = "SyncButton";
            this.SyncButton.ShowImage = true;
            this.SyncButton.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.SyncButton_Click);
            // 
            // group3
            // 
            this.group3.Items.Add(this.ShowHideButton);
            this.group3.Items.Add(this.UploadButton);
            this.group3.Items.Add(this.LogoutButton);
            this.group3.Items.Add(this.StatusLabel);
            this.group3.Name = "group3";
            // 
            // ShowHideButton
            // 
            this.ShowHideButton.Image = ((System.Drawing.Image)(resources.GetObject("ShowHideButton.Image")));
            this.ShowHideButton.Label = "Hide/Show Fields";
            this.ShowHideButton.Name = "ShowHideButton";
            this.ShowHideButton.ShowImage = true;
            this.ShowHideButton.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.ShowHideButton_Click);
            // 
            // UploadButton
            // 
            this.UploadButton.Image = ((System.Drawing.Image)(resources.GetObject("UploadButton.Image")));
            this.UploadButton.Label = "Upload Data";
            this.UploadButton.Name = "UploadButton";
            this.UploadButton.ShowImage = true;
            this.UploadButton.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.UploadButton_Click);
            // 
            // LogoutButton
            // 
            this.LogoutButton.Image = ((System.Drawing.Image)(resources.GetObject("LogoutButton.Image")));
            this.LogoutButton.Label = "Logout";
            this.LogoutButton.Name = "LogoutButton";
            this.LogoutButton.ShowImage = true;
            this.LogoutButton.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.LogoutButton_Click);
            // 
            // StatusLabel
            // 
            this.StatusLabel.Label = " ";
            this.StatusLabel.Name = "StatusLabel";
            // 
            // DavigoldRibbon
            // 
            this.Name = "DavigoldRibbon";
            this.RibbonType = "Microsoft.Excel.Workbook";
            this.Tabs.Add(this.tab1);
            this.Load += new Microsoft.Office.Tools.Ribbon.RibbonUIEventHandler(this.DavigoldRibbon_Load);
            this.tab1.ResumeLayout(false);
            this.tab1.PerformLayout();
            this.group1.ResumeLayout(false);
            this.group1.PerformLayout();
            this.group2.ResumeLayout(false);
            this.group2.PerformLayout();
            this.group3.ResumeLayout(false);
            this.group3.PerformLayout();
            this.ResumeLayout(false);

        }

        #endregion

        internal Microsoft.Office.Tools.Ribbon.RibbonTab tab1;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup group1;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton ImportButton;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup group2;
        internal Microsoft.Office.Tools.Ribbon.RibbonLabel VersionLabel;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup group3;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton ShowHideButton;
        internal Microsoft.Office.Tools.Ribbon.RibbonLabel UpdatedOnLabel;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton LogoutButton;
        internal Microsoft.Office.Tools.Ribbon.RibbonLabel StatusLabel;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton UploadButton;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton SyncButton;
    }

    partial class ThisRibbonCollection
    {
        internal DavigoldRibbon DavigoldRibbon
        {
            get { return this.GetRibbon<DavigoldRibbon>(); }
        }
    }
}
