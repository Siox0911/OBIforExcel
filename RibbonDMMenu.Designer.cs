namespace OBIforExcel
{
    /// <summary>
    /// Ribbonmenu in Excel
    /// </summary>
    partial class RibbonDMMenu : Microsoft.Office.Tools.Ribbon.RibbonBase
    {
        /// <summary>
        /// Erforderliche Designervariable.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        /// <summary>
        /// Neuer RibbonMenu Eintrag
        /// </summary>
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
            Microsoft.Office.Tools.Ribbon.RibbonDropDownItem ribbonDropDownItemImpl2 = this.Factory.CreateRibbonDropDownItem();
            Microsoft.Office.Tools.Ribbon.RibbonDropDownItem ribbonDropDownItemImpl3 = this.Factory.CreateRibbonDropDownItem();
            Microsoft.Office.Tools.Ribbon.RibbonDropDownItem ribbonDropDownItemImpl4 = this.Factory.CreateRibbonDropDownItem();
            this.tab1 = this.Factory.CreateRibbonTab();
            this.groupBarcode = this.Factory.CreateRibbonGroup();
            this.ddCodeType = this.Factory.CreateRibbonDropDown();
            this.checkBoxFitToCell = this.Factory.CreateRibbonCheckBox();
            this.checkBoxCellToPictureSize = this.Factory.CreateRibbonCheckBox();
            this.checkBoxCellPosition = this.Factory.CreateRibbonCheckBox();
            this.separator1 = this.Factory.CreateRibbonSeparator();
            this.btnEinfuegen = this.Factory.CreateRibbonButton();
            this.ddCellBinding = this.Factory.CreateRibbonDropDown();
            this.separator2 = this.Factory.CreateRibbonSeparator();
            this.tab1.SuspendLayout();
            this.groupBarcode.SuspendLayout();
            this.SuspendLayout();
            // 
            // tab1
            // 
            this.tab1.ControlId.ControlIdType = Microsoft.Office.Tools.Ribbon.RibbonControlIdType.Office;
            this.tab1.Groups.Add(this.groupBarcode);
            this.tab1.Label = "TabAddIns";
            this.tab1.Name = "tab1";
            // 
            // groupBarcode
            // 
            this.groupBarcode.Items.Add(this.ddCodeType);
            this.groupBarcode.Items.Add(this.checkBoxFitToCell);
            this.groupBarcode.Items.Add(this.checkBoxCellToPictureSize);
            this.groupBarcode.Items.Add(this.separator2);
            this.groupBarcode.Items.Add(this.ddCellBinding);
            this.groupBarcode.Items.Add(this.checkBoxCellPosition);
            this.groupBarcode.Items.Add(this.separator1);
            this.groupBarcode.Items.Add(this.btnEinfuegen);
            this.groupBarcode.Label = "OBI for Excel";
            this.groupBarcode.Name = "groupBarcode";
            // 
            // ddCodeType
            // 
            ribbonDropDownItemImpl1.Label = "Datamatrix";
            ribbonDropDownItemImpl1.Tag = "DMC";
            this.ddCodeType.Items.Add(ribbonDropDownItemImpl1);
            this.ddCodeType.Label = "Codetype";
            this.ddCodeType.Name = "ddCodeType";
            this.ddCodeType.SelectionChanged += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.DdCodeType_SelectionChanged);
            // 
            // checkBoxFitToCell
            // 
            this.checkBoxFitToCell.Label = "Adapt to cell size";
            this.checkBoxFitToCell.Name = "checkBoxFitToCell";
            this.checkBoxFitToCell.ScreenTip = "The barcode is fitted directly into the cell size";
            this.checkBoxFitToCell.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.CheckBoxFitToCell_Click);
            // 
            // checkBoxCellToPictureSize
            // 
            this.checkBoxCellToPictureSize.Label = "Adapt cell to picture size";
            this.checkBoxCellToPictureSize.Name = "checkBoxCellToPictureSize";
            this.checkBoxCellToPictureSize.ScreenTip = "The cell is adjusted to the image size";
            this.checkBoxCellToPictureSize.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.CheckBoxCellToPictureSize_Click);
            // 
            // checkBoxCellPosition
            // 
            this.checkBoxCellPosition.Checked = true;
            this.checkBoxCellPosition.Enabled = false;
            this.checkBoxCellPosition.Label = "On cell position";
            this.checkBoxCellPosition.Name = "checkBoxCellPosition";
            this.checkBoxCellPosition.ScreenTip = "Should the barcode be generated directly at the position of the cell?";
            // 
            // separator1
            // 
            this.separator1.Name = "separator1";
            // 
            // btnEinfuegen
            // 
            this.btnEinfuegen.Label = "Insert";
            this.btnEinfuegen.Name = "btnEinfuegen";
            this.btnEinfuegen.ScreenTip = "If a cell is marked, the barcode is created as an image at this position.";
            this.btnEinfuegen.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.BtnEinfuegen_Click);
            // 
            // ddCellBinding
            // 
            ribbonDropDownItemImpl2.Label = "Free Float";
            ribbonDropDownItemImpl2.ScreenTip = "Picture is free floating";
            ribbonDropDownItemImpl3.Label = "Move";
            ribbonDropDownItemImpl3.ScreenTip = "Picture is moved with the cell";
            ribbonDropDownItemImpl4.Label = "Move and Size";
            ribbonDropDownItemImpl4.ScreenTip = "Picture is moved and sized with the cell";
            this.ddCellBinding.Items.Add(ribbonDropDownItemImpl2);
            this.ddCellBinding.Items.Add(ribbonDropDownItemImpl3);
            this.ddCellBinding.Items.Add(ribbonDropDownItemImpl4);
            this.ddCellBinding.Label = "Cell binding";
            this.ddCellBinding.Name = "ddCellBinding";
            this.ddCellBinding.ScreenTip = "How should the picture be bound to the cell?";
            // 
            // separator2
            // 
            this.separator2.Name = "separator2";
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
        internal Microsoft.Office.Tools.Ribbon.RibbonCheckBox checkBoxCellToPictureSize;
        internal Microsoft.Office.Tools.Ribbon.RibbonDropDown ddCellBinding;
        internal Microsoft.Office.Tools.Ribbon.RibbonSeparator separator2;
    }

    partial class ThisRibbonCollection
    {
        internal RibbonDMMenu RibbonDMMenu
        {
            get { return this.GetRibbon<RibbonDMMenu>(); }
        }
    }
}
