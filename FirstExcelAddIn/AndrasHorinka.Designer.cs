namespace FirstExcelAddIn
{
    partial class AndrasHorinka : Microsoft.Office.Tools.Ribbon.RibbonBase
    {
        /// <summary>
        /// Required designer variable.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        public AndrasHorinka()
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
            this.MNB = this.Factory.CreateRibbonTab();
            this.mnbExtract = this.Factory.CreateRibbonGroup();
            this.button1 = this.Factory.CreateRibbonButton();
            this.Log = this.Factory.CreateRibbonButton();
            this.MNB.SuspendLayout();
            this.mnbExtract.SuspendLayout();
            this.SuspendLayout();
            // 
            // MNB
            // 
            this.MNB.ControlId.ControlIdType = Microsoft.Office.Tools.Ribbon.RibbonControlIdType.Office;
            this.MNB.ControlId.OfficeId = "TabHome";
            this.MNB.Groups.Add(this.mnbExtract);
            this.MNB.Label = "TabHome";
            this.MNB.Name = "MNB";
            // 
            // mnbExtract
            // 
            this.mnbExtract.Items.Add(this.button1);
            this.mnbExtract.Items.Add(this.Log);
            this.mnbExtract.Label = "Horinka András";
            this.mnbExtract.Name = "mnbExtract";
            // 
            // button1
            // 
            this.button1.Label = "MNB adatletöltés";
            this.button1.Name = "button1";
            this.button1.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.button1_Click);
            // 
            // Log
            // 
            this.Log.Description = "Provide clarification on the request";
            this.Log.Label = "Log";
            this.Log.Name = "Log";
            this.Log.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.button2_Click);
            // 
            // AndrasHorinka
            // 
            this.Name = "AndrasHorinka";
            this.RibbonType = "Microsoft.Excel.Workbook";
            this.Tabs.Add(this.MNB);
            this.Load += new Microsoft.Office.Tools.Ribbon.RibbonUIEventHandler(this.AndrasHorinka_Load);
            this.MNB.ResumeLayout(false);
            this.MNB.PerformLayout();
            this.mnbExtract.ResumeLayout(false);
            this.mnbExtract.PerformLayout();
            this.ResumeLayout(false);

        }

        #endregion

        internal Microsoft.Office.Tools.Ribbon.RibbonTab MNB;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup mnbExtract;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton button1;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton Log;
    }

    partial class ThisRibbonCollection
    {
        internal AndrasHorinka AndrasHorinka
        {
            get { return this.GetRibbon<AndrasHorinka>(); }
        }
    }
}
