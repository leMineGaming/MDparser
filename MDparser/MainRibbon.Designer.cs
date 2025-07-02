namespace MDparser
{
    partial class MainRibbon : Microsoft.Office.Tools.Ribbon.RibbonBase
    {
        /// <summary>
        /// Required designer variable.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        public MainRibbon()
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
            this.mdGroup = this.Factory.CreateRibbonGroup();
            this.insertMarkdown = this.Factory.CreateRibbonButton();
            this.tab1.SuspendLayout();
            this.mdGroup.SuspendLayout();
            this.SuspendLayout();
            // 
            // tab1
            // 
            this.tab1.ControlId.ControlIdType = Microsoft.Office.Tools.Ribbon.RibbonControlIdType.Office;
            this.tab1.ControlId.OfficeId = "TabInsert";
            this.tab1.Groups.Add(this.mdGroup);
            this.tab1.Label = "TabInsert";
            this.tab1.Name = "tab1";
            // 
            // mdGroup
            // 
            this.mdGroup.Items.Add(this.insertMarkdown);
            this.mdGroup.Label = "MD Parser";
            this.mdGroup.Name = "mdGroup";
            this.mdGroup.Position = this.Factory.RibbonPosition.AfterOfficeId("TabInsert");
            // 
            // insertMarkdown
            // 
            this.insertMarkdown.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.insertMarkdown.Image = global::MDparser.Properties.Resources.icons8_markdown_100;
            this.insertMarkdown.Label = "Insert Markdown";
            this.insertMarkdown.Name = "insertMarkdown";
            this.insertMarkdown.ShowImage = true;
            this.insertMarkdown.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.insertMarkdown_Click);
            // 
            // MainRibbon
            // 
            this.Name = "MainRibbon";
            this.RibbonType = "Microsoft.PowerPoint.Presentation";
            this.Tabs.Add(this.tab1);
            this.Load += new Microsoft.Office.Tools.Ribbon.RibbonUIEventHandler(this.MainRibbon_Load);
            this.tab1.ResumeLayout(false);
            this.tab1.PerformLayout();
            this.mdGroup.ResumeLayout(false);
            this.mdGroup.PerformLayout();
            this.ResumeLayout(false);

        }

        #endregion

        internal Microsoft.Office.Tools.Ribbon.RibbonTab tab1;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup mdGroup;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton insertMarkdown;
    }

    partial class ThisRibbonCollection
    {
        internal MainRibbon MainRibbon
        {
            get { return this.GetRibbon<MainRibbon>(); }
        }
    }
}
