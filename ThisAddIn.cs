using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Xml.Linq;
using Excel = Microsoft.Office.Interop.Excel;
using Office = Microsoft.Office.Core;
using DataMatrix.net;
using OBIforExcel.CellCache;

/*
 * Dieses Addin funktioniert mit Excel 2016. Es wird nicht die Bibliothek Microsoft.Office.Interop.Excel Version 15 
 * vorausgesetzt. Diese ist in den PIA Verweisen nicht enthalten. Im Manifest steht also keine Excelversion, so 
 * dass bei einer älteren Office Version wie 2010, 2013 das Addin durchaus funktionieren sollte. Es kann der PIA 
 * Verweis aber auch an eine Excelversion gekoppelt werden. Dazu einfach bei den Verweisen unter
 * Microsoft.Office.Interop.Excel in den Eigenschaften die Einstellung Interoptypen einbetten auf False stellen. 
 * Dann wird im Manifest direkt auf die Interop Version 15 verwiesen und diese muss dann installiert sein. 
 * Ansonsten kann das AddIn in Excel nicht geladen werden.
 */

namespace OBIforExcel
{
    /// <summary>
    /// Eigentliches AddIn welches in Excel geladen wird.
    /// </summary>
    public partial class ThisAddIn
    {
        /// <summary>
        /// Der aktuelle Zellen Cache im aktuellen Workbook
        /// </summary>
        private WorkbookCache workbookCache;

        /// <summary>
        /// Das aktuelle Workbook
        /// </summary>
        private string workbookName;

        /// <summary>
        /// Wird aufgerufen, sobald das AddIn geladen wurde
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void ThisAddIn_Startup(object sender, System.EventArgs e)
        {
            //Ein paar Events von Excel abgreifen

            //Unter Excel 2010 kann es passieren, dass der nächste Event nicht abgefeuert wird.
            //Dann ist meistens schon ein Workbook geladen, wir greifen hier vorher ein.
            if (this.Application.ActiveWorkbook != null)
            {
                WorkbookActivate(this.Application.ActiveWorkbook);
            }

            //Wenn ein Workbook aktiviert wird. Wird beim Laden eines Workbooks aufgerufen
            this.Application.WorkbookActivate += WorkbookActivate;

            //Wird aufgerufen, wenn eine Zelle geändert oder neu berechnet wurde.
            this.Application.AfterCalculate += AppAfterCalculate;
        }

        /// <summary>
        /// Wird aufgerufen, sobald das AddIn deaktiviert wird
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void ThisAddIn_Shutdown(object sender, System.EventArgs e)
        {
            //Wenn das AddIn heruntergefahren wird, dann lösche den Zellen Cache
            workbookCache = null;
        }

        /// <summary>
        /// Wird jedesmal aufgerufen wenn sich eine Zelle ändert, egal ob eingegeben oder berechnet.
        /// </summary>
        private void AppAfterCalculate()
        {
            try
            {
                CheckWorkbook();
            }
            catch (Exception ex)
            {
                System.Windows.Forms.MessageBox.Show(
                    ex.Message
                    , "OBIforExcel - Error on checking workbook cache"
                    , System.Windows.Forms.MessageBoxButtons.OK
                    , System.Windows.Forms.MessageBoxIcon.Error
                );
            }
        }

        /// <summary>
        /// Wird aufgerufen, wenn das Workbook aktiviert wird
        /// </summary>
        /// <param name="Wb"></param>
        private void WorkbookActivate(Excel.Workbook Wb)
        {
            //System.Windows.Forms.MessageBox.Show($"Workbook {((Excel.Workbook)Wb).Name} aktiviert");
            try
            {
                this.workbookName = Wb.Name;
                //Zellen Cache erstellen
                CreateWorkbookCache(Wb);
            }
            catch (Exception ex)
            {
                System.Windows.Forms.MessageBox.Show(
                    ex.Message
                    , "OBIforExcel - Error on creating workbook cache"
                    , System.Windows.Forms.MessageBoxButtons.OK
                    , System.Windows.Forms.MessageBoxIcon.Error
                );
            }
        }

        /// <summary>
        /// Erzeugt den Zellen Cache. Dabei werden die Prüfungen durch den CellCache und die CellShapes durchgeführt.
        /// Hier wird nur zusätzlich der Worksheet auf null geprüft
        /// </summary>
        /// <param name="workbook"></param>
        private void CreateWorkbookCache(Excel.Workbook workbook)
        {
            if (workbook != null)
            {
                //Erstelle einen neuen Workbook Cache
                workbookCache = new WorkbookCache(workbook);
            }
        }

        /// <summary>
        /// Prüft das Workbook, wenn der Cache existiert
        /// </summary>
        private void CheckWorkbook()
        {
            workbookCache?.CheckWorkbook();
        }

        /// <summary>
        /// Prüft den aktuellen Worksheet ob sich da im Zellen Cache befindlichen Adressen etwas geändert hat.
        /// </summary>
        /// <param name="worksheet"></param>
        private void CheckWorksheet(Excel.Worksheet worksheet)
        {
            if (worksheet != null)
            {
                workbookCache?.CheckWorkSheet(worksheet);
            }
        }

        /// <summary>
        /// Prüft alle Zellen ob sich die Werte darin geändert haben, gegen die im Zellen Cache vorhandenen
        /// </summary>
        /// <param name="range"></param>
        private void CheckRange(Excel.Range range)
        {
            if (range != null)
            {
                workbookCache?.CheckRange(range);
            }
        }

        /// <summary>
        /// Gibt den aktiven Worksheet zurück. Wenn keiner ausgewählt ist, wird null zurückgegeben.
        /// </summary>
        /// <returns></returns>
        private Excel.Worksheet GetActiveWorksheet()
        {
            return (Excel.Worksheet)Application.ActiveSheet;
        }

        /// <summary>
        /// Gibt die aktuelle gewählte Zelle zurück. Wenn mehr als eine Zelle merkiiert ist, dann wird 
        /// <c>null</c> zurückgegeben.
        /// </summary>
        /// <returns></returns>
        public Excel.Range GetCurrentCell()
        {
            /*
             * Wenn nur eine Zelle markiert ist, dann gib diese zurück.
             * Sofern diese nicht null ist.
             */
            if (Application.ActiveCell?.Count == 1)
            {
                return Application.ActiveCell;
            }

            //Sonst gib null zurück
            return null;
        }

        /// <summary>
        /// Gibt die aktuelle Auswahl der Zellen zurück. Sollte verwendet werden, wenn <see cref="GetCurrentCell"/> 
        /// null zurück gibt. Kann null zurückgeben, falls die Auswahl keine Zellen beinhaltet.
        /// </summary>
        /// <returns></returns>
        public Excel.Range GetCurrentSelection()
        {
            //Nur eine Range zurückgeben, falls es auch eine ist
            if(Application.Selection is Excel.Range)
            {
                return ((Excel.Range)Application.Selection).Cells;
            }

            return null;
        }

        /// <summary>
        /// Fügt den Barcode den Zellen in der Range hinzu
        /// </summary>
        /// <param name="range">Die Zellen welche den Barcode erhalten sollen.</param>
        /// <param name="xlPlacement">Die Bindung des Bildes an die Zelle: free float, move or move and size</param>
        /// <param name="fitToCell">Bild an die Zellengröße anpassen</param>
        /// <param name="cellFitToPicture">Zelle an die Bildgröße anpassen</param>
        public void AddPictures(Excel.Range range, Excel.XlPlacement xlPlacement, bool fitToCell = false, bool cellFitToPicture = false)
        {
            try
            {
                workbookCache?.AddCellShape(this.Application.Workbooks[workbookName], range, xlPlacement, fitToCell, cellFitToPicture);
            }
            catch (Exception ex)
            {
                System.Windows.Forms.MessageBox.Show(
                    ex.Message
                    , "Error on creating barcode images"
                    , System.Windows.Forms.MessageBoxButtons.OK
                    , System.Windows.Forms.MessageBoxIcon.Error);
            }
        }

        #region Von VSTO generierter Code

        /// <summary>
        /// Erforderliche Methode für die Designerunterstützung.
        /// Der Inhalt der Methode darf nicht mit dem Code-Editor geändert werden.
        /// </summary>
        private void InternalStartup()
        {
            this.Startup += new System.EventHandler(ThisAddIn_Startup);
            this.Shutdown += new System.EventHandler(ThisAddIn_Shutdown);
        }

        #endregion
    }
}
