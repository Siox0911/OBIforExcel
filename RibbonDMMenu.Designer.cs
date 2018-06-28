namespace OBIforExcel
{
    partial class RibbonDMMenu : Microsoft.Office.Tools.Ribbon.RibbonBase
    {
        /// <summary>
        /// Erforderliche Designervariable.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        public RibbonDMMenu()
            : base(Globals.Factory.GetRibbonFactory())
        {
            InitializeComponent();
        }

        /// <summary> 
        /// Verwendete Ressourcen bereinigen.
        /// </summary>
        /// <param name="disposing">"true", wenn verwaltete Ressourcen gelöscht werden sollen, andernfalls "false".</param>
        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
            {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        #region Vom Komponenten-Designer generierter Code

        /// <summary>
        /// Erforderliche Methode für Designerunterstützung -
        /// Der Inhalt der Methode darf nicht mit dem Code-Editor geändert werden.
        /// </summary>
        private void InitializeComponent()
        {
            Microsoft.Office.Tools.Ribbon.RibbonDropDownItem ribbonDropDownItemImpl1 = this.Factory.CreateRibbonDropDownItem();
            this.tab1 = this.Factory.CreateRibbonTab();
            this.groupBarcode = this.Factory.CreateRibbonGroup();
            this.ddCodeType = this.Factory.CreateRibbonDropDown();
            this.checkBoxCellPosition = this.Factory.CreateRibbonCheckBox();
            this.separator1 = this.Factory.CreateRibbonSeparator();
            this.btnEinfuegen = this.Factory.CreateRibbonButton();
            this.checkBoxFitToCell = this.Factory.CreateRibbonCheckBox();
            this.tab1.SuspendLayout();
            this.groupBarcode.SuspendLayout();
            this.SuspendLayout();
            // 
            // tab1
            // 
            this.tab1.ControlId.ControlIdType = Microsoft.Office.Tools.Ribbon.RibbonControlIdType.Office;
            this.tab1.Groups.Add(this.groupBarcode);
            this.tab1.Label = "OBI";
            this.tab1.Name = "tab1";
            // 
            // groupBarcode
            // 
            this.groupBarcode.Items.Add(this.ddCodeType);
            this.groupBarcode.Items.Add(this.checkBoxCellPosition);
            this.groupBarcode.Items.Add(this.checkBoxFitToCell);
            this.groupBarcode.Items.Add(this.separator1);
            this.groupBarcode.Items.Add(this.btnEinfuegen);
            this.groupBarcode.Label = "Barcode";
            this.groupBarcode.Name = "groupBarcode";
            // 
            // ddCodeType
            // 
            ribbonDropDownItemImpl1.Label = "Datamatrix";
            ribbonDropDownItemImpl1.Tag = "DMC";
            this.ddCodeType.Items.Add(ribbonDropDownItemImpl1);
            this.ddCodeType.Label = "Codetyp";
            this.ddCodeType.Name = "ddCodeType";
            this.ddCodeType.SelectionChanged += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.DdCodeType_SelectionChanged);
            // 
            // checkBoxCellPosition
            // 
            this.checkBoxCellPosition.Checked = true;
            this.checkBoxCellPosition.Enabled = false;
            this.checkBoxCellPosition.Label = "An Zellposition";
            this.checkBoxCellPosition.Name = "checkBoxCellPosition";
            this.checkBoxCellPosition.ScreenTip = "Soll der Barcode direk an der Position der Zelle erzeugt werden?";
            // 
            // separator1
            // 
            this.separator1.Name = "separator1";
            // 
            // btnEinfuegen
            // 
            this.btnEinfuegen.Label = "Einfügen";
            this.btnEinfuegen.Name = "btnEinfuegen";
            this.btnEinfuegen.ScreenTip = "Wenn eine Zelle markiert ist, wird an dieser Position der Barcode als Bild erzeug" +
    "t.";
            this.btnEinfuegen.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.BtnEinfuegen_Click);
            // 
            // checkBoxFitToCell
            // 
            this.checkBoxFitToCell.Label = "An Zellgröße anpassen";
            this.checkBoxFitToCell.Name = "checkBoxFitToCell";
            this.checkBoxFitToCell.ScreenTip = "Der Barcode wird direkt in die Zellgröße eingepasst";
            // 
            // RibbonDMMenu
            // 
            this.Name = "RibbonDMMenu";
            this.RibbonType = "Microsoft.Excel.Workbook";
            this.Tabs.Add(this.tab1);
            this.Load += new Microsoft.Office.Tools.Ribbon.RibbonUIEventHandler(this.RibbonDMMenu_Load);
            this.tab1.ResumeLayout(false);
            this.tab1.PerformLayout();
            this.groupBarcode.ResumeLayout(false);
            this.groupBarcode.PerformLayout();
            this.ResumeLayout(false);

        }

        #endregion

        internal Microsoft.Office.Tools.Ribbon.RibbonTab tab1;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup groupBarcode;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnEinfuegen;
        internal Microsoft.Office.Tools.Ribbon.RibbonCheckBox checkBoxCellPosition;
        internal Microsoft.Office.Tools.Ribbon.RibbonSeparator separator1;
        internal Microsoft.Office.Tools.Ribbon.RibbonDropDown ddCodeType;
        internal Microsoft.Office.Tools.Ribbon.RibbonCheckBox checkBoxFitToCell;
    }

    partial class ThisRibbonCollection
    {
        internal RibbonDMMenu RibbonDMMenu
        {
            get { return this.GetRibbon<RibbonDMMenu>(); }
        }
    }
}
