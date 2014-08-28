namespace save_to_google_drive
{
    partial class GoogleDrive : Microsoft.Office.Tools.Ribbon.RibbonBase
    {
        /// <summary>
        /// Required designer variable.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        public GoogleDrive()
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
            this.driveTab = this.Factory.CreateRibbonTab();
            this.googleDriveGroup = this.Factory.CreateRibbonGroup();
            this.buttonGD = this.Factory.CreateRibbonButton();
            this.driveTab.SuspendLayout();
            this.googleDriveGroup.SuspendLayout();
            // 
            // driveTab
            // 
            this.driveTab.ControlId.ControlIdType = Microsoft.Office.Tools.Ribbon.RibbonControlIdType.Office;
            this.driveTab.ControlId.OfficeId = "TabMail";
            this.driveTab.Groups.Add(this.googleDriveGroup);
            this.driveTab.Label = "TabMail";
            this.driveTab.Name = "driveTab";
            // 
            // googleDriveGroup
            // 
            this.googleDriveGroup.Items.Add(this.buttonGD);
            this.googleDriveGroup.Label = "Google Drive";
            this.googleDriveGroup.Name = "googleDriveGroup";
            // 
            // buttonGD
            // 
            this.buttonGD.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.buttonGD.Description = "Save to Google Drive";
            this.buttonGD.Image = global::save_to_google_drive.Resources.DriveLogoBig;
            this.buttonGD.Label = "Save to Google Drive";
            this.buttonGD.Name = "buttonGD";
            this.buttonGD.ShowImage = true;
            this.buttonGD.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.buttonGD_Click);
            // 
            // GoogleDrive
            // 
            this.Name = "GoogleDrive";
            this.RibbonType = "Microsoft.Outlook.Explorer";
            this.Tabs.Add(this.driveTab);
            this.Load += new Microsoft.Office.Tools.Ribbon.RibbonUIEventHandler(this.GoogleDrive_Load);
            this.driveTab.ResumeLayout(false);
            this.driveTab.PerformLayout();
            this.googleDriveGroup.ResumeLayout(false);
            this.googleDriveGroup.PerformLayout();

        }

        #endregion

        internal Microsoft.Office.Tools.Ribbon.RibbonTab driveTab;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup googleDriveGroup;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton buttonGD;
    }

    partial class ThisRibbonCollection
    {
        internal GoogleDrive GoogleDrive
        {
            get { return this.GetRibbon<GoogleDrive>(); }
        }
    }
}
