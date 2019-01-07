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
            ddCellBinding.SelectedItemIndex = 1;
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
                //Bildbindung an die Zelle
                var xlPlacement = Microsoft.Office.Interop.Excel.XlPlacement.xlMove;
                switch (ddCellBinding.SelectedItemIndex)
                {
                    case 0:
                        {
                            xlPlacement = Microsoft.Office.Interop.Excel.XlPlacement.xlFreeFloating;
                            break;
                        }
                    case 1:
                        {
                            xlPlacement = Microsoft.Office.Interop.Excel.XlPlacement.xlMove;
                            break;
                        }
                    case 2:
                        {
                            xlPlacement = Microsoft.Office.Interop.Excel.XlPlacement.xlMoveAndSize;
                            break;
                        }
                }

                //Die aktuelle Auswahl an Zellen abrufen
                var cells = Globals.ThisAddIn.GetCurrentSelection();
                if (cells != null)
                {
                    //und die Bilder der Range hinzufügen
                    Globals.ThisAddIn.AddPictures(cells, xlPlacement, checkBoxFitToCell.Checked, checkBoxCellToPictureSize.Checked);
                }
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

        private void CheckBoxCellToPictureSize_Click(object sender, RibbonControlEventArgs e)
        {
            //Es kann nur einen geben. Wenn beide aktiv sind, wir automatisch die andere
            //abgewählt. Entweder die Zelle wird an die Bildgröße angepasst oder das
            //Bild an die Zelle
            if (checkBoxFitToCell.Checked && checkBoxCellToPictureSize.Checked)
            {
                checkBoxFitToCell.Checked = !checkBoxCellToPictureSize.Checked;
            }
        }

        private void CheckBoxFitToCell_Click(object sender, RibbonControlEventArgs e)
        {
            //Es kann nur einen geben. Wenn beide aktiv sind, wir automatisch die andere
            //abgewählt. Entweder die Zelle wird an die Bildgröße angepasst oder das
            //Bild an die Zelle
            if (checkBoxCellToPictureSize.Checked && checkBoxFitToCell.Checked)
            {
                checkBoxCellToPictureSize.Checked = !checkBoxFitToCell.Checked;
            }
        }
    }
}
