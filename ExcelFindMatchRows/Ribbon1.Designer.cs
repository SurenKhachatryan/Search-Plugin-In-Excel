namespace ExcelFindMatchRows
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
            this.searchEditBox = this.Factory.CreateRibbonEditBox();
            this.ProgressLabel = this.Factory.CreateRibbonLabel();
            this.buttonCencel = this.Factory.CreateRibbonButton();
            this.buttonSearch = this.Factory.CreateRibbonButton();
            this.tab1.SuspendLayout();
            this.group1.SuspendLayout();
            this.SuspendLayout();
            // 
            // tab1
            // 
            this.tab1.ControlId.ControlIdType = Microsoft.Office.Tools.Ribbon.RibbonControlIdType.Office;
            this.tab1.Groups.Add(this.group1);
            this.tab1.Label = "Search In Document";
            this.tab1.Name = "tab1";
            // 
            // group1
            // 
            this.group1.Items.Add(this.searchEditBox);
            this.group1.Items.Add(this.ProgressLabel);
            this.group1.Items.Add(this.buttonCencel);
            this.group1.Items.Add(this.buttonSearch);
            this.group1.Label = "Search Block";
            this.group1.Name = "group1";
            // 
            // searchEditBox
            // 
            this.searchEditBox.Label = " ";
            this.searchEditBox.Name = "searchEditBox";
            this.searchEditBox.Text = null;
            // 
            // ProgressLabel
            // 
            this.ProgressLabel.Label = "In Processing...";
            this.ProgressLabel.Name = "ProgressLabel";
            this.ProgressLabel.Visible = false;
            // 
            // buttonCencel
            // 
            this.buttonCencel.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.buttonCencel.Image = global::ExcelFindMatchRows.Properties.Resources.Red_Close_Button;
            this.buttonCencel.Label = "Cencel";
            this.buttonCencel.Name = "buttonCencel";
            this.buttonCencel.ShowImage = true;
            this.buttonCencel.Visible = false;
            this.buttonCencel.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.buttonCencel_Click);
            // 
            // buttonSearch
            // 
            this.buttonSearch.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.buttonSearch.Image = global::ExcelFindMatchRows.Properties.Resources.Zerode_Plump_Search;
            this.buttonSearch.Label = "Search";
            this.buttonSearch.Name = "buttonSearch";
            this.buttonSearch.ShowImage = true;
            this.buttonSearch.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.Search_Button_Click);
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
            this.ResumeLayout(false);

        }

        #endregion

        internal Microsoft.Office.Tools.Ribbon.RibbonTab tab1;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup group1;
        internal Microsoft.Office.Tools.Ribbon.RibbonEditBox searchEditBox;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton buttonSearch;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton buttonCencel;
        internal Microsoft.Office.Tools.Ribbon.RibbonLabel ProgressLabel;
    }

    partial class ThisRibbonCollection
    {
        internal Ribbon1 Ribbon1
        {
            get { return this.GetRibbon<Ribbon1>(); }
        }
    }
}
