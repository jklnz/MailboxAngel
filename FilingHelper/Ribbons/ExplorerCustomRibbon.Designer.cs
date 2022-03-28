namespace FilingHelper
{
    partial class ExplorerCustomRibbon : Microsoft.Office.Tools.Ribbon.RibbonBase
    {
        /// <summary>
        /// Required designer variable.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        public ExplorerCustomRibbon()
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
            this.components = new System.ComponentModel.Container();
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(ExplorerCustomRibbon));
            this.tab1 = this.Factory.CreateRibbonTab();
            this.tab2 = this.Factory.CreateRibbonTab();
            this.group2 = this.Factory.CreateRibbonGroup();
            this.menu2 = this.Factory.CreateRibbonMenu();
            this.btnItemReplyAll = this.Factory.CreateRibbonButton();
            this.btnItemReply = this.Factory.CreateRibbonButton();
            this.tab3 = this.Factory.CreateRibbonTab();
            this.group1 = this.Factory.CreateRibbonGroup();
            this.btnCloseAll = this.Factory.CreateRibbonButton();
            this.imageList1 = new System.Windows.Forms.ImageList(this.components);
            this.menu1 = this.Factory.CreateRibbonMenu();
            this.tab1.SuspendLayout();
            this.tab2.SuspendLayout();
            this.group2.SuspendLayout();
            this.tab3.SuspendLayout();
            this.group1.SuspendLayout();
            this.SuspendLayout();
            // 
            // tab1
            // 
            this.tab1.ControlId.ControlIdType = Microsoft.Office.Tools.Ribbon.RibbonControlIdType.Office;
            this.tab1.ControlId.OfficeId = "TabFolder";
            this.tab1.Label = "TabFolder";
            this.tab1.Name = "tab1";
            // 
            // tab2
            // 
            this.tab2.ControlId.ControlIdType = Microsoft.Office.Tools.Ribbon.RibbonControlIdType.Office;
            this.tab2.ControlId.OfficeId = "TabMail";
            this.tab2.Groups.Add(this.group2);
            this.tab2.Label = "TabMail";
            this.tab2.Name = "tab2";
            // 
            // group2
            // 
            this.group2.Items.Add(this.menu2);
            this.group2.Label = "Repond (Adv)";
            this.group2.Name = "group2";
            this.group2.Position = this.Factory.RibbonPosition.AfterOfficeId("GroupMailRespond");
            // 
            // menu2
            // 
            this.menu2.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.menu2.Image = ((System.Drawing.Image)(resources.GetObject("menu2.Image")));
            this.menu2.Items.Add(this.btnItemReplyAll);
            this.menu2.Items.Add(this.btnItemReply);
            this.menu2.Label = "Reply w/ Attachments";
            this.menu2.Name = "menu2";
            this.menu2.ShowImage = true;
            // 
            // btnItemReplyAll
            // 
            this.btnItemReplyAll.Label = "Reply All";
            this.btnItemReplyAll.Name = "btnItemReplyAll";
            this.btnItemReplyAll.OfficeImageId = "ReplyAll";
            this.btnItemReplyAll.ShowImage = true;
            this.btnItemReplyAll.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnItemReplyAll_Click);
            // 
            // btnItemReply
            // 
            this.btnItemReply.Label = "Reply Sender";
            this.btnItemReply.Name = "btnItemReply";
            this.btnItemReply.OfficeImageId = "Reply";
            this.btnItemReply.ShowImage = true;
            this.btnItemReply.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnItemReply_Click);
            // 
            // tab3
            // 
            this.tab3.ControlId.ControlIdType = Microsoft.Office.Tools.Ribbon.RibbonControlIdType.Office;
            this.tab3.ControlId.OfficeId = "TabView";
            this.tab3.Groups.Add(this.group1);
            this.tab3.Label = "TabView";
            this.tab3.Name = "tab3";
            // 
            // group1
            // 
            this.group1.Items.Add(this.btnCloseAll);
            this.group1.Label = "Window Helper";
            this.group1.Name = "group1";
            // 
            // btnCloseAll
            // 
            this.btnCloseAll.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.btnCloseAll.Label = "Close All";
            this.btnCloseAll.Name = "btnCloseAll";
            this.btnCloseAll.OfficeImageId = "ViewDisplayInHighContrast";
            this.btnCloseAll.ShowImage = true;
            this.btnCloseAll.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnCloseAll_Click);
            // 
            // imageList1
            // 
            this.imageList1.ColorDepth = System.Windows.Forms.ColorDepth.Depth8Bit;
            this.imageList1.ImageSize = new System.Drawing.Size(16, 16);
            this.imageList1.TransparentColor = System.Drawing.Color.Transparent;
            // 
            // menu1
            // 
            this.menu1.Label = "menu1";
            this.menu1.Name = "menu1";
            this.menu1.ShowImage = true;
            // 
            // ExplorerCustomRibbon
            // 
            this.Name = "ExplorerCustomRibbon";
            // 
            // ExplorerCustomRibbon.OfficeMenu
            // 
            this.OfficeMenu.Items.Add(this.menu1);
            this.RibbonType = "Microsoft.Outlook.Explorer";
            this.Tabs.Add(this.tab1);
            this.Tabs.Add(this.tab2);
            this.Tabs.Add(this.tab3);
            this.tab1.ResumeLayout(false);
            this.tab1.PerformLayout();
            this.tab2.ResumeLayout(false);
            this.tab2.PerformLayout();
            this.group2.ResumeLayout(false);
            this.group2.PerformLayout();
            this.tab3.ResumeLayout(false);
            this.tab3.PerformLayout();
            this.group1.ResumeLayout(false);
            this.group1.PerformLayout();
            this.ResumeLayout(false);

        }

        #endregion
        internal Microsoft.Office.Tools.Ribbon.RibbonMenu menu1;
        private Microsoft.Office.Tools.Ribbon.RibbonTab tab1;
        internal Microsoft.Office.Tools.Ribbon.RibbonTab tab2;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup group2;
        internal Microsoft.Office.Tools.Ribbon.RibbonMenu menu2;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnItemReplyAll;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnItemReply;
        internal Microsoft.Office.Tools.Ribbon.RibbonTab tab3;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup group1;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnCloseAll;
        private System.Windows.Forms.ImageList imageList1;
    }

    partial class ThisRibbonCollection
    {
        internal ExplorerCustomRibbon MainRibbon
        {
            get { return this.GetRibbon<ExplorerCustomRibbon>(); }
        }
    }
}
