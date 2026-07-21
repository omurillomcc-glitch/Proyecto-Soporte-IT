using Microsoft.Office.Tools.Ribbon;

namespace ExcelAddInRCP_ICG
{
    partial class Ribbon1 : Microsoft.Office.Tools.Ribbon.RibbonBase
    {
        private System.ComponentModel.IContainer components = null;

        public Ribbon1()
            : base(Globals.Factory.GetRibbonFactory())
        {
            InitializeComponent();
        }

        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
            {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        #region Código generado

        private void InitializeComponent()
        {
            Microsoft.Office.Tools.Ribbon.RibbonDialogLauncher ribbonDialogLauncherImpl1 = this.Factory.CreateRibbonDialogLauncher();
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(Ribbon1));

            this.tab1 = this.Factory.CreateRibbonTab();

            // ✅ GRUPOS
            this.group1 = this.Factory.CreateRibbonGroup();
            this.groupFiltro = this.Factory.CreateRibbonGroup();

            this.txtFechaIni = this.Factory.CreateRibbonEditBox();
            this.txtFechaFin = this.Factory.CreateRibbonEditBox();

            // ✅ NUEVOS CONTROLES
            this.boxFiltroPais = this.Factory.CreateRibbonBox();
            this.boxCheckPais = this.Factory.CreateRibbonBox();
            this.lblFiltroPais = this.Factory.CreateRibbonButton();

            this.chkHN = this.Factory.CreateRibbonCheckBox();
            this.chkNI = this.Factory.CreateRibbonCheckBox();

            this.button1 = this.Factory.CreateRibbonButton();
            this.button2 = this.Factory.CreateRibbonButton();
            this.button3 = this.Factory.CreateRibbonButton();

            this.tab1.SuspendLayout();
            this.group1.SuspendLayout();
            this.groupFiltro.SuspendLayout();
            this.SuspendLayout();

            // tab1
            this.tab1.ControlId.ControlIdType = Microsoft.Office.Tools.Ribbon.RibbonControlIdType.Office;
            this.tab1.Groups.Add(this.group1);
            this.tab1.Groups.Add(this.groupFiltro);
            this.tab1.Label = "Supply Chain";
            this.tab1.Name = "tab1";

            // =========================
            // ✅ GRUPO PRINCIPAL
            // =========================
            this.group1.DialogLauncher = ribbonDialogLauncherImpl1;

            this.group1.Items.Add(this.txtFechaIni);
            this.group1.Items.Add(this.txtFechaFin);
            this.group1.Items.Add(this.button1);
            this.group1.Items.Add(this.button2);
            this.group1.Items.Add(this.button3);

            this.group1.Label = "RCP->ICG";
            this.group1.Name = "group1";

            // =========================
            // ✅ GRUPO FILTRO PAIS (CON TITULO ARRIBA)
            // =========================
            // Grupo
            this.groupFiltro.Label = "Filtros";

            // Caja principal
            this.boxFiltroPais.BoxStyle =
                Microsoft.Office.Tools.Ribbon.RibbonBoxStyle.Vertical;

            // Etiqueta
            this.lblFiltroPais.Label = "País";
            this.lblFiltroPais.Enabled = true;

            // Caja de checks
            this.boxCheckPais.BoxStyle =
                Microsoft.Office.Tools.Ribbon.RibbonBoxStyle.Vertical;

            // Checkboxes
            this.chkHN.Label = "Honduras";
            this.chkHN.Name = "chkHN";

            this.chkNI.Label = "Nicaragua";
            this.chkNI.Name = "chkNI";

            // Estructura
            this.boxCheckPais.Items.Add(this.chkHN);
            this.boxCheckPais.Items.Add(this.chkNI);

            this.boxFiltroPais.Items.Add(this.lblFiltroPais);
            this.boxFiltroPais.Items.Add(this.boxCheckPais);

            this.groupFiltro.Items.Add(this.boxFiltroPais);

            this.chkHN.Click += new RibbonControlEventHandler(chkHN_Click);
            this.chkNI.Click += new RibbonControlEventHandler(chkNI_Click);
            //____________________ FIN FILTRO PAIS____________________________________________________________


            // txtFechaIni
            this.txtFechaIni.Label = "Fecha Inicial";
            this.txtFechaIni.Name = "txtFechaIni";
            this.txtFechaIni.SuperTip = "dd-mm-aaaa";
            this.txtFechaIni.Text = null;
            this.txtFechaIni.TextChanged += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.txtFechaIni_TextChanged);

            // txtFechaFin
            this.txtFechaFin.Label = "Fecha Final";
            this.txtFechaFin.Name = "txtFechaFin";
            this.txtFechaFin.SuperTip = "dd-mm-aaaa";
            this.txtFechaFin.Text = null;
            this.txtFechaFin.TextChanged += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.txtFechaFin_TextChanged);

            // button1
            this.button1.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.button1.Image = ((System.Drawing.Image)(resources.GetObject("button1.Image")));
            this.button1.Label = "Repartir";
            this.button1.Name = "button1";
            this.button1.ShowImage = true;
            this.button1.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.button1_Click_1);

            // button2
            this.button2.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.button2.Image = ((System.Drawing.Image)(resources.GetObject("button2.Image")));
            this.button2.Label = "Comprar";
            this.button2.Name = "button2";
            this.button2.ShowImage = true;
            this.button2.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.button2_Click);

            // button3
            this.button3.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.button3.Image = ((System.Drawing.Image)(resources.GetObject("button3.Image")));
            this.button3.Label = "Traspasos";
            this.button3.Name = "button3";
            this.button3.ShowImage = true;
            this.button3.Visible = false;
            this.button3.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.button3_Click);

            // Ribbon1
            this.Name = "Ribbon1";
            this.RibbonType = "Microsoft.Excel.Workbook";
            this.Tabs.Add(this.tab1);
            this.Load += new Microsoft.Office.Tools.Ribbon.RibbonUIEventHandler(this.Ribbon1_Load);

            this.tab1.ResumeLayout(false);
            this.tab1.PerformLayout();
            this.group1.ResumeLayout(false);
            this.group1.PerformLayout();
            this.groupFiltro.ResumeLayout(false);
            this.groupFiltro.PerformLayout();
            this.ResumeLayout(false);
        }

        #endregion

        // ✅ NUEVOS CONTROLES
        internal Microsoft.Office.Tools.Ribbon.RibbonBox boxFiltroPais;
        internal Microsoft.Office.Tools.Ribbon.RibbonBox boxCheckPais;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton lblFiltroPais;

        internal Microsoft.Office.Tools.Ribbon.RibbonCheckBox chkHN;
        internal Microsoft.Office.Tools.Ribbon.RibbonCheckBox chkNI;

        internal Microsoft.Office.Tools.Ribbon.RibbonGroup groupFiltro;
        internal Microsoft.Office.Tools.Ribbon.RibbonTab tab1;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup group1;

        internal Microsoft.Office.Tools.Ribbon.RibbonButton button1;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton button2;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton button3;

        internal Microsoft.Office.Tools.Ribbon.RibbonEditBox txtFechaIni;
        internal Microsoft.Office.Tools.Ribbon.RibbonEditBox txtFechaFin;
    }

    partial class ThisRibbonCollection
    {
        internal Ribbon1 Ribbon1
        {
            get { return this.GetRibbon<Ribbon1>(); }
        }
    }
}