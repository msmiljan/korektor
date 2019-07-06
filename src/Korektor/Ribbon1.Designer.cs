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
            this.splitButtonCyr = this.Factory.CreateRibbonSplitButton();
            this.buttonCyr = this.Factory.CreateRibbonButton();
            this.buttonLatToCyr = this.Factory.CreateRibbonButton();
            this.buttonCyrPodelReci = this.Factory.CreateRibbonButton();
            this.buttonCyrSastaviReci = this.Factory.CreateRibbonButton();
            this.splitButtonLat = this.Factory.CreateRibbonSplitButton();
            this.buttonLat = this.Factory.CreateRibbonButton();
            this.buttonCyrToLat = this.Factory.CreateRibbonButton();
            this.buttonLatPodeliReci = this.Factory.CreateRibbonButton();
            this.buttonLatSastaviReci = this.Factory.CreateRibbonButton();
            this.buttonLatToLat = this.Factory.CreateRibbonButton();
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
            this.grpKorektor.Items.Add(this.splitButtonCyr);
            this.grpKorektor.Items.Add(this.splitButtonLat);
            this.grpKorektor.Label = "Коректор";
            this.grpKorektor.Name = "grpKorektor";
            this.grpKorektor.Position = this.Factory.RibbonPosition.BeforeOfficeId("GroupComments");
            // 
            // splitButtonCyr
            // 
            this.splitButtonCyr.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.splitButtonCyr.Image = global::Korektor.Properties.Resources.SpellCheck_Cir;
            this.splitButtonCyr.Items.Add(this.buttonCyr);
            this.splitButtonCyr.Items.Add(this.buttonLatToCyr);
            this.splitButtonCyr.Items.Add(this.buttonCyrPodelReci);
            this.splitButtonCyr.Items.Add(this.buttonCyrSastaviReci);
            this.splitButtonCyr.Label = "Ћирилица";
            this.splitButtonCyr.Name = "splitButtonCyr";
            this.splitButtonCyr.ScreenTip = "Ћирилица";
            this.splitButtonCyr.SuperTip = "Провера речи - ћирилица";
            this.splitButtonCyr.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.splitButtonCyr_Click);
            // 
            // buttonCyr
            // 
            this.buttonCyr.Image = global::Korektor.Properties.Resources.SpellCheck_Cir;
            this.buttonCyr.Label = "Ћирилица";
            this.buttonCyr.Name = "buttonCyr";
            this.buttonCyr.ScreenTip = "Ћирилица";
            this.buttonCyr.ShowImage = true;
            this.buttonCyr.SuperTip = "Провера речи - ћирилица";
            this.buttonCyr.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.splitButtonCyr_Click);
            // 
            // buttonLatToCyr
            // 
            this.buttonLatToCyr.Image = global::Korektor.Properties.Resources.LatToCir;
            this.buttonLatToCyr.Label = "Lat  →  Ћир";
            this.buttonLatToCyr.Name = "buttonLatToCyr";
            this.buttonLatToCyr.ScreenTip = "Lat  →  Ћир";
            this.buttonLatToCyr.ShowImage = true;
            this.buttonLatToCyr.SuperTip = "Конверзија латинице у ћирилицу";
            this.buttonLatToCyr.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.buttonLatToCyr_Click);
            // 
            // buttonCyrPodelReci
            // 
            this.buttonCyrPodelReci.Image = global::Korektor.Properties.Resources.Hyphenate;
            this.buttonCyrPodelReci.Label = "Подели на слогове";
            this.buttonCyrPodelReci.Name = "buttonCyrPodelReci";
            this.buttonCyrPodelReci.ScreenTip = "Подели на слогове";
            this.buttonCyrPodelReci.ShowImage = true;
            this.buttonCyrPodelReci.SuperTip = "Подела речи на слогове - ћирилица";
            this.buttonCyrPodelReci.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.buttonCyrPodeliReci_click);
            // 
            // buttonCyrSastaviReci
            // 
            this.buttonCyrSastaviReci.Image = global::Korektor.Properties.Resources.UnHyphenate;
            this.buttonCyrSastaviReci.Label = "Састави слогове";
            this.buttonCyrSastaviReci.Name = "buttonCyrSastaviReci";
            this.buttonCyrSastaviReci.ScreenTip = "Састави слогове";
            this.buttonCyrSastaviReci.ShowImage = true;
            this.buttonCyrSastaviReci.SuperTip = "Састављање слогова у речи - ћирилица";
            this.buttonCyrSastaviReci.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.buttonCyrSastaviReci_click);
            // 
            // splitButtonLat
            // 
            this.splitButtonLat.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.splitButtonLat.Image = global::Korektor.Properties.Resources.SpellCheck_Lat;
            this.splitButtonLat.Items.Add(this.buttonLat);
            this.splitButtonLat.Items.Add(this.buttonCyrToLat);
            this.splitButtonLat.Items.Add(this.buttonLatPodeliReci);
            this.splitButtonLat.Items.Add(this.buttonLatSastaviReci);
            this.splitButtonLat.Items.Add(this.buttonLatToLat);
            this.splitButtonLat.Label = "Latinica";
            this.splitButtonLat.Name = "splitButtonLat";
            this.splitButtonLat.ScreenTip = "Latinica";
            this.splitButtonLat.SuperTip = "Provera reči - latinicа";
            this.splitButtonLat.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.splitButtonLat_Click);
            // 
            // buttonLat
            // 
            this.buttonLat.Image = global::Korektor.Properties.Resources.SpellCheck_Lat;
            this.buttonLat.Label = "Latinica";
            this.buttonLat.Name = "buttonLat";
            this.buttonLat.ScreenTip = "Latinica";
            this.buttonLat.ShowImage = true;
            this.buttonLat.SuperTip = "Provera reči - latinicа";
            this.buttonLat.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.splitButtonLat_Click);
            // 
            // buttonCyrToLat
            // 
            this.buttonCyrToLat.Image = global::Korektor.Properties.Resources.CirToLat;
            this.buttonCyrToLat.Label = "Ћир  →  Lat";
            this.buttonCyrToLat.Name = "buttonCyrToLat";
            this.buttonCyrToLat.ScreenTip = "Ћир  →  Lat";
            this.buttonCyrToLat.ShowImage = true;
            this.buttonCyrToLat.SuperTip = "Konverzija ćirilice u latinicu ";
            this.buttonCyrToLat.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.buttonCyrToLat_Click);
            // 
            // buttonLatPodeliReci
            // 
            this.buttonLatPodeliReci.Image = global::Korektor.Properties.Resources.Hyphenate;
            this.buttonLatPodeliReci.Label = "Podeli na slogove";
            this.buttonLatPodeliReci.Name = "buttonLatPodeliReci";
            this.buttonLatPodeliReci.ScreenTip = "Podeli na slogove";
            this.buttonLatPodeliReci.ShowImage = true;
            this.buttonLatPodeliReci.SuperTip = "Podela reči na slogove - latinica";
            this.buttonLatPodeliReci.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.buttonLatPodeliReci_Click);
            // 
            // buttonLatSastaviReci
            // 
            this.buttonLatSastaviReci.Image = global::Korektor.Properties.Resources.UnHyphenate;
            this.buttonLatSastaviReci.Label = "Sastavi slogove";
            this.buttonLatSastaviReci.Name = "buttonLatSastaviReci";
            this.buttonLatSastaviReci.ScreenTip = "Sastavi slogove";
            this.buttonLatSastaviReci.ShowImage = true;
            this.buttonLatSastaviReci.SuperTip = "Sastavljanje slogova u reči - latinica";
            this.buttonLatSastaviReci.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.buttonLatSastaviReci_Click);
            // 
            // buttonLatToLat
            // 
            this.buttonLatToLat.Image = global::Korektor.Properties.Resources.LatToLat;
            this.buttonLatToLat.Label = "Cccsszz  →  Cčćsšzž";
            this.buttonLatToLat.Name = "buttonLatToLat";
            this.buttonLatToLat.ScreenTip = "Cccsszz  →  Cčćsšzž";
            this.buttonLatToLat.ShowImage = true;
            this.buttonLatToLat.SuperTip = "Konverzija ošišane latinice u latinicu";
            this.buttonLatToLat.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.ribbonButtonLatToLat_Click);
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
        internal Microsoft.Office.Tools.Ribbon.RibbonSplitButton splitButtonCyr;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton buttonCyrToLat;
        internal Microsoft.Office.Tools.Ribbon.RibbonSplitButton splitButtonLat;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton buttonLatToCyr;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton buttonCyr;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton buttonLat;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton buttonCyrPodelReci;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton buttonLatPodeliReci;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton buttonCyrSastaviReci;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton buttonLatSastaviReci;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton buttonLatToLat;
    }

    partial class ThisRibbonCollection
    {
        internal Ribbon1 Ribbon1
        {
            get { return this.GetRibbon<Ribbon1>(); }
        }
    }
}
