namespace ExcelTest
{
    partial class DemoRibbon : Microsoft.Office.Tools.Ribbon.RibbonBase
    {
        /// <summary>
        /// Required designer variable.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        public DemoRibbon()
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
            this.DemoTab = this.Factory.CreateRibbonTab();
            this.DemoGroup = this.Factory.CreateRibbonGroup();
            this.DemoButton1 = this.Factory.CreateRibbonButton();
            this.DemoButton2 = this.Factory.CreateRibbonButton();
            this.DemoButton3 = this.Factory.CreateRibbonButton();
            this.tab1 = this.Factory.CreateRibbonTab();
            this.group1 = this.Factory.CreateRibbonGroup();
            this.button1 = this.Factory.CreateRibbonButton();
            this.DemoTab.SuspendLayout();
            this.DemoGroup.SuspendLayout();
            this.tab1.SuspendLayout();
            this.group1.SuspendLayout();
            this.SuspendLayout();
            // 
            // DemoTab
            // 
            this.DemoTab.ControlId.ControlIdType = Microsoft.Office.Tools.Ribbon.RibbonControlIdType.Office;
            this.DemoTab.Groups.Add(this.DemoGroup);
            this.DemoTab.Label = "DemoTab";
            this.DemoTab.Name = "DemoTab";
            this.DemoTab.Tag = "DemoTab";
            // 
            // DemoGroup
            // 
            this.DemoGroup.Items.Add(this.DemoButton1);
            this.DemoGroup.Items.Add(this.DemoButton2);
            this.DemoGroup.Items.Add(this.DemoButton3);
            this.DemoGroup.Label = "DemoGroup";
            this.DemoGroup.Name = "DemoGroup";
            // 
            // DemoButton1
            // 
            this.DemoButton1.Description = "DemoButton1";
            this.DemoButton1.Label = "DemoButton1";
            this.DemoButton1.Name = "DemoButton1";
            this.DemoButton1.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.DemoButton1_Click);
            // 
            // DemoButton2
            // 
            this.DemoButton2.Label = "DemoButton2";
            this.DemoButton2.Name = "DemoButton2";
            this.DemoButton2.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.DemoButton2_Click);
            // 
            // DemoButton3
            // 
            this.DemoButton3.Label = "DemoButton3";
            this.DemoButton3.Name = "DemoButton3";
            this.DemoButton3.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.DemoButton3_Click);
            // 
            // tab1
            // 
            this.tab1.Groups.Add(this.group1);
            this.tab1.Label = "TestTAb";
            this.tab1.Name = "tab1";
            // 
            // group1
            // 
            this.group1.Items.Add(this.button1);
            this.group1.Label = "group1";
            this.group1.Name = "group1";
            // 
            // button1
            // 
            this.button1.Label = "button1";
            this.button1.Name = "button1";
            // 
            // DemoRibbon
            // 
            this.Name = "DemoRibbon";
            this.RibbonType = "Microsoft.Excel.Workbook";
            this.Tabs.Add(this.DemoTab);
            this.Tabs.Add(this.tab1);
            this.Load += new Microsoft.Office.Tools.Ribbon.RibbonUIEventHandler(this.DemoRibbon_Load);
            this.DemoTab.ResumeLayout(false);
            this.DemoTab.PerformLayout();
            this.DemoGroup.ResumeLayout(false);
            this.DemoGroup.PerformLayout();
            this.tab1.ResumeLayout(false);
            this.tab1.PerformLayout();
            this.group1.ResumeLayout(false);
            this.group1.PerformLayout();
            this.ResumeLayout(false);

        }

        #endregion

        internal Microsoft.Office.Tools.Ribbon.RibbonTab DemoTab;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup DemoGroup;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton DemoButton1;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton DemoButton2;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton DemoButton3;
        private Microsoft.Office.Tools.Ribbon.RibbonTab tab1;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup group1;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton button1;
    }

    partial class ThisRibbonCollection
    {
        internal DemoRibbon DemoRibbon
        {
            get { return this.GetRibbon<DemoRibbon>(); }
        }
    }
}
