using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Microsoft.Office.Tools.Ribbon;

namespace OBIforExcel
{
    public partial class RibbonDMMenu
    {
        //Brauchen wir nicht, noch nicht
        private void RibbonDMMenu_Load(object sender, RibbonUIEventArgs e)
        {

        }

        /// <summary>
        /// Wenn der Button Einfügen gedrückt wird, dann füge einen Barcode ein, falls möglich
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void BtnEinfuegen_Click(object sender, RibbonControlEventArgs e)
        {
            //Fehler abfangen
            try
            {
                //Die aktuelle Auswahl an Zellen abrufen
                var cells = Globals.ThisAddIn.GetCurrentSelection();
                //und die Bilder der Range hinzufügen
                Globals.ThisAddIn.AddPictures(cells, checkBoxFitToCell.Checked);
            }
            catch (Exception ex)
            {
                System.Windows.Forms.MessageBox.Show(ex.Message, "Error in OBIforExcel", System.Windows.Forms.MessageBoxButtons.OK, System.Windows.Forms.MessageBoxIcon.Error);
            }
        }


        private void DdCodeType_SelectionChanged(object sender, RibbonControlEventArgs e)
        {
            //Momentan gibt es nur DataMatrix
        }
    }
}
