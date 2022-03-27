namespace FilingHelper
{
    partial class ComposeInspectorCustomRibbon : Microsoft.Office.Tools.Ribbon.RibbonBase
    {
        /// <summary>
        /// Required designer variable.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        public ComposeInspectorCustomRibbon()
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
            Microsoft.Office.Tools.Ribbon.RibbonDialogLauncher ribbonDialogLauncherImpl1 = this.Factory.CreateRibbonDialogLauncher();
            Microsoft.Office.Tools.Ribbon.RibbonDialogLauncher ribbonDialogLauncherImpl2 = this.Factory.CreateRibbonDialogLauncher();
            this.tab1 = this.Factory.CreateRibbonTab();
            this.ComposeGroup = this.Factory.CreateRibbonGroup();
            this.btnAttachments = this.Factory.CreateRibbonButton();
            this.tab2 = this.Factory.CreateRibbonTab();
            this.group1 = this.Factory.CreateRibbonGroup();
            this.button1 = this.Factory.CreateRibbonButton();
            this.tab1.SuspendLayout();
            this.ComposeGroup.SuspendLayout();
            this.tab2.SuspendLayout();
            this.group1.SuspendLayout();
            this.SuspendLayout();
            // 
            // tab1
            // 
            this.tab1.ControlId.ControlIdType = Microsoft.Office.Tools.Ribbon.RibbonControlIdType.Office;
            this.tab1.ControlId.OfficeId = "TabNewMailMessage";
            this.tab1.Groups.Add(this.ComposeGroup);
            this.tab1.Label = "TabNewMailMessage";
            this.tab1.Name = "tab1";
            // 
            // ComposeGroup
            // 
            this.ComposeGroup.DialogLauncher = ribbonDialogLauncherImpl1;
            this.ComposeGroup.Items.Add(this.btnAttachments);
            this.ComposeGroup.Label = "Mailbox Angel";
            this.ComposeGroup.Name = "ComposeGroup";
            this.ComposeGroup.DialogLauncherClick += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.ComposeGroup_DialogLauncherClick);
            // 
            // btnAttachments
            // 
            this.btnAttachments.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.btnAttachments.Label = "Attachment Helper";
            this.btnAttachments.Name = "btnAttachments";
            this.btnAttachments.OfficeImageId = "MultiplePages";
            this.btnAttachments.ShowImage = true;
            this.btnAttachments.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnAttachments_Click);
            // 
            // tab2
            // 
            this.tab2.ControlId.ControlIdType = Microsoft.Office.Tools.Ribbon.RibbonControlIdType.Office;
            this.tab2.ControlId.OfficeId = "TabMail";
            this.tab2.Groups.Add(this.group1);
            this.tab2.Label = "TabMail";
            this.tab2.Name = "tab2";
            // 
            // group1
            // 
            this.group1.DialogLauncher = ribbonDialogLauncherImpl2;
            this.group1.Items.Add(this.button1);
            this.group1.Label = "Mailbox Angel";
            this.group1.Name = "group1";
            // 
            // button1
            // 
            this.button1.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.button1.Label = "Attachment Helper";
            this.button1.Name = "button1";
            this.button1.OfficeImageId = "MultiplePages";
            this.button1.ShowImage = true;
            // 
            // ComposeInspectorCustomRibbon
            // 
            this.Name = "ComposeInspectorCustomRibbon";
            this.RibbonType = "Microsoft.Outlook.Mail.Compose";
            this.Tabs.Add(this.tab1);
            this.Tabs.Add(this.tab2);
            this.tab1.ResumeLayout(false);
            this.tab1.PerformLayout();
            this.ComposeGroup.ResumeLayout(false);
            this.ComposeGroup.PerformLayout();
            this.tab2.ResumeLayout(false);
            this.tab2.PerformLayout();
            this.group1.ResumeLayout(false);
            this.group1.PerformLayout();
            this.ResumeLayout(false);

        }

        #endregion

        internal Microsoft.Office.Tools.Ribbon.RibbonTab tab1;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup ComposeGroup;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnAttachments;
        private Microsoft.Office.Tools.Ribbon.RibbonTab tab2;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup group1;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton button1;
    }

    partial class ThisRibbonCollection
    {
        internal ComposeInspectorCustomRibbon ComposeInspectorCustomRibbon
        {
            get { return this.GetRibbon<ComposeInspectorCustomRibbon>(); }
        }
    }
}
