namespace Korektor
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
            this.tabKorektura = this.Factory.CreateRibbonTab();
            this.grpKorektor = this.Factory.CreateRibbonGroup();
            this.ribbonButtonCyr = this.Factory.CreateRibbonButton();
            this.ribbonButtonLat = this.Factory.CreateRibbonButton();
            this.ribbonButtonCyrLat = this.Factory.CreateRibbonButton();
            this.tabKorektura.SuspendLayout();
            this.grpKorektor.SuspendLayout();
            this.SuspendLayout();
            // 
            // tabKorektura
            // 
            this.tabKorektura.ControlId.ControlIdType = Microsoft.Office.Tools.Ribbon.RibbonControlIdType.Office;
            this.tabKorektura.ControlId.OfficeId = "TabReviewWord";
            this.tabKorektura.Groups.Add(this.grpKorektor);
            this.tabKorektura.Label = "TabReviewWord";
            this.tabKorektura.Name = "tabKorektura";
            // 
            // grpKorektor
            // 
            this.grpKorektor.Items.Add(this.ribbonButtonCyr);
            this.grpKorektor.Items.Add(this.ribbonButtonLat);
            this.grpKorektor.Items.Add(this.ribbonButtonCyrLat);
            this.grpKorektor.Label = "Коректор";
            this.grpKorektor.Name = "grpKorektor";
            this.grpKorektor.Position = this.Factory.RibbonPosition.BeforeOfficeId("GroupComments");
            // 
            // ribbonButtonCyr
            // 
            this.ribbonButtonCyr.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.ribbonButtonCyr.Image = global::Korektor.Properties.Resources.SpellCheck_Cir;
            this.ribbonButtonCyr.Label = "Ћирилица";
            this.ribbonButtonCyr.Name = "ribbonButtonCyr";
            this.ribbonButtonCyr.ShowImage = true;
            this.ribbonButtonCyr.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.ribbonButtonCyr_Click);
            // 
            // ribbonButtonLat
            // 
            this.ribbonButtonLat.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.ribbonButtonLat.Image = global::Korektor.Properties.Resources.SpellCheck_Lat;
            this.ribbonButtonLat.Label = "Latinica";
            this.ribbonButtonLat.Name = "ribbonButtonLat";
            this.ribbonButtonLat.ShowImage = true;
            this.ribbonButtonLat.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.ribbonButtonLat_Click);
            // 
            // ribbonButtonCyrLat
            // 
            this.ribbonButtonCyrLat.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.ribbonButtonCyrLat.Image = global::Korektor.Properties.Resources.SpellCheck_CirLat;
            this.ribbonButtonCyrLat.Label = "Ћирилица Latinica";
            this.ribbonButtonCyrLat.Name = "ribbonButtonCyrLat";
            this.ribbonButtonCyrLat.ShowImage = true;
            this.ribbonButtonCyrLat.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.ribbonButtonCyrLat_Click);
            // 
            // Ribbon1
            // 
            this.Name = "Ribbon1";
            this.RibbonType = "Microsoft.Word.Document";
            this.Tabs.Add(this.tabKorektura);
            this.Close += new System.EventHandler(this.Ribbon1_Close);
            this.Load += new Microsoft.Office.Tools.Ribbon.RibbonUIEventHandler(this.Ribbon1_Load);
            this.tabKorektura.ResumeLayout(false);
            this.tabKorektura.PerformLayout();
            this.grpKorektor.ResumeLayout(false);
            this.grpKorektor.PerformLayout();
            this.ResumeLayout(false);

        }

        #endregion

        internal Microsoft.Office.Tools.Ribbon.RibbonTab tabKorektura;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup grpKorektor;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton ribbonButtonCyr;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton ribbonButtonLat;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton ribbonButtonCyrLat;
    }

    partial class ThisRibbonCollection
    {
        internal Ribbon1 Ribbon1
        {
            get { return this.GetRibbon<Ribbon1>(); }
        }
    }
}
