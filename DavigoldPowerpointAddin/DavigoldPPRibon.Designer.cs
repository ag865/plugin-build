namespace DavigoldPowerpointAddin
{
    partial class DavigoldPPRibon : Microsoft.Office.Tools.Ribbon.RibbonBase
    {
        /// <summary>
        /// Required designer variable.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        public DavigoldPPRibon()
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
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(DavigoldPPRibon));
            this.tab1 = this.Factory.CreateRibbonTab();
            this.group1 = this.Factory.CreateRibbonGroup();
            this.VersionLabel = this.Factory.CreateRibbonLabel();
            this.UpdatedOnLabel = this.Factory.CreateRibbonLabel();
            this.group2 = this.Factory.CreateRibbonGroup();
            this.ShowHideButton = this.Factory.CreateRibbonButton();
            this.LoginLogoutButton = this.Factory.CreateRibbonButton();
            this.group3 = this.Factory.CreateRibbonGroup();
            this.SyncButton = this.Factory.CreateRibbonButton();
            this.button1 = this.Factory.CreateRibbonButton();
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
            this.group1.Items.Add(this.VersionLabel);
            this.group1.Items.Add(this.UpdatedOnLabel);
            this.group1.Name = "group1";
            // 
            // VersionLabel
            // 
            this.VersionLabel.Label = "Version";
            this.VersionLabel.Name = "VersionLabel";
            // 
            // UpdatedOnLabel
            // 
            this.UpdatedOnLabel.Label = "Last Updated on";
            this.UpdatedOnLabel.Name = "UpdatedOnLabel";
            // 
            // group2
            // 
            this.group2.Items.Add(this.ShowHideButton);
            this.group2.Items.Add(this.LoginLogoutButton);
            this.group2.Name = "group2";
            // 
            // ShowHideButton
            // 
            this.ShowHideButton.Image = ((System.Drawing.Image)(resources.GetObject("ShowHideButton.Image")));
            this.ShowHideButton.Label = "Hide/Show Fields";
            this.ShowHideButton.Name = "ShowHideButton";
            this.ShowHideButton.ShowImage = true;
            this.ShowHideButton.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.ShowHideButton_Click);
            // 
            // LoginLogoutButton
            // 
            this.LoginLogoutButton.Image = ((System.Drawing.Image)(resources.GetObject("LoginLogoutButton.Image")));
            this.LoginLogoutButton.Label = "Login";
            this.LoginLogoutButton.Name = "LoginLogoutButton";
            this.LoginLogoutButton.ShowImage = true;
            this.LoginLogoutButton.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.LoginLogoutButton_Click);
            // 
            // group3
            // 
            this.group3.Items.Add(this.SyncButton);
            this.group3.Items.Add(this.button1);
            this.group3.Name = "group3";
            // 
            // SyncButton
            // 
            this.SyncButton.Label = "Publish Version";
            this.SyncButton.Name = "SyncButton";
            this.SyncButton.ShowImage = true;
            this.SyncButton.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.SyncButton_Click);
            // 
            // button1
            // 
            this.button1.Image = ((System.Drawing.Image)(resources.GetObject("button1.Image")));
            this.button1.Label = "Update Links";
            this.button1.Name = "button1";
            this.button1.ShowImage = true;
            // 
            // DavigoldPPRibon
            // 
            this.Name = "DavigoldPPRibon";
            this.RibbonType = "Microsoft.PowerPoint.Presentation";
            this.Tabs.Add(this.tab1);
            this.Load += new Microsoft.Office.Tools.Ribbon.RibbonUIEventHandler(this.DavigoldPPRibon_Load);
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
        internal Microsoft.Office.Tools.Ribbon.RibbonLabel VersionLabel;
        internal Microsoft.Office.Tools.Ribbon.RibbonLabel UpdatedOnLabel;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup group2;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton ShowHideButton;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton LoginLogoutButton;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton SyncButton;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup group3;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton button1;
    }

    partial class ThisRibbonCollection
    {
        internal DavigoldPPRibon DavigoldPPRibon
        {
            get { return this.GetRibbon<DavigoldPPRibon>(); }
        }
    }
}
