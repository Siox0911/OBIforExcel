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
        /// Der aktuelle Zellen Cache im aktuellen Worksheet
        /// </summary>
        private CellCache.CellCache cellCache;

        /// <summary>
        /// Wird aufgerufen, sobald das AddIn geladen wurde
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void ThisAddIn_Startup(object sender, System.EventArgs e)
        {
            //Ein paar Events von Excel abgreifen

            //Ein Sheet wurde aktiviert, z.B. Tabelle1
            this.Application.SheetActivate += SheetActivate;
            //Ein Sheet wurde deaktiviert, z.B. Tabelle1 und Tabelle2 wurde dafür dann aktiviert
            this.Application.SheetDeactivate += SheetDeactivate;
            //Wenn ein Workbook aktiviert wird. Wird beim Laden eines Workbooks aufgerufen
            this.Application.WorkbookActivate += WorkbookActivate;
        }

        private void WorkbookActivate(Excel.Workbook Wb)
        {
            //System.Windows.Forms.MessageBox.Show($"Workbook {((Excel.Workbook)Wb).Name} aktiviert");
            try
            {
                //Zellen Cache erstellen
                CreateCellCache(GetActiveWorksheet());
                //Den aktuellen Sheet aktivieren, weil das nicht automatisch passiert
                SheetActivate(GetActiveWorksheet());
            }
            catch (Exception ex)
            {
                System.Windows.Forms.MessageBox.Show(
                    ex.Message
                    , "Error on checking active Workbook"
                    , System.Windows.Forms.MessageBoxButtons.OK
                    , System.Windows.Forms.MessageBoxIcon.Error
                );
            }
        }

        /// <summary>
        /// Wird aufgerufen, sobald das AddIn deaktiviert wird
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void ThisAddIn_Shutdown(object sender, System.EventArgs e)
        {
            //Wenn das AddIn heruntergefahren wird, dann lösche den Zellen Cache
            cellCache = null;
        }

        /// <summary>
        /// Der aktuelle Sheet wurde abgewählt
        /// </summary>
        /// <param name="Sh"></param>
        private void SheetDeactivate(object Sh)
        {
            //Wenn ein Sheet deaktiviert wird, dann werden die Events abgemeldet
            this.Application.SheetChange -= SheetChanged;
            this.Application.SheetCalculate -= SheetCalculate;
            cellCache = null;
        }

        /// <summary>
        /// Es wurde ein anderer Sheet ausgewählt
        /// </summary>
        /// <param name="Sh"></param>
        private void SheetActivate(object Sh)
        {
            //Wenn ein Sheet aktiviert wird, dann werden die Events registiert.
            CreateCellCache((Excel.Worksheet)Sh);
            this.Application.SheetChange += SheetChanged;
            this.Application.SheetCalculate += SheetCalculate;
        }

        /// <summary>
        /// Event welcher aufgerufen wird, sobald der Sheet neu berechnet wird
        /// </summary>
        /// <param name="Sh"></param>
        private void SheetCalculate(object Sh)
        {
            if (((Excel.Worksheet)Sh) != null)
            {
                CheckWorksheet((Excel.Worksheet)Sh);
            }
        }

        /// <summary>
        /// Zellen im Workbook haben sich geändert. Nicht davon betroffen sind berechnete und abhängige Zellen im 
        /// kompletten Workbook.
        /// </summary>
        /// <param name="Sh"></param>
        /// <param name="Target"></param>
        private void SheetChanged(object Sh, Excel.Range Target)
        {
            CheckRange(Target);
        }

        /// <summary>
        /// Erzeugt den Zellen Cache. Dabei werden die Prüfungen durch den CellCache und die CellShapes durchgeführt.
        /// Hier wird nur zusätzlich der Worksheet auf null geprüft
        /// </summary>
        /// <param name="worksheet"></param>
        private void CreateCellCache(Excel.Worksheet worksheet)
        {
            if (worksheet != null)
            {
                //Erstelle einen neuen Zellen Cache
                cellCache = new CellCache.CellCache(worksheet);
            }
        }

        /// <summary>
        /// Prüft den aktuellen Worksheet ob sich da im Zellen Cache befindlichen Adressen etwas geändert hat.
        /// </summary>
        /// <param name="worksheet"></param>
        private void CheckWorksheet(Excel.Worksheet worksheet)
        {
            //Das machen wir aber nur, wenn der Worksheetname mit dem Namen im Zellen Cache übereinstimmt.
            //Weil wir nicht wissen was in den anderen Tabellenblättern so ist, da dies im Cache nicht abgebildet ist
            if (worksheet?.Name.Equals(cellCache.WorksheetName) == true)
            {
                //Wir laufen durch unseren Zellen Cache
                foreach (var item in cellCache.CellShapes)
                {
                    //Und prüfen jeden Eintrag auf Änderungen
                    CheckRange(worksheet.Range[item.Address]);
                }
            }
        }

        /// <summary>
        /// Prüft alle Zellen ob sich die Werte darin geändert haben, gegen die im Zellen Cache vorhandenen
        /// </summary>
        /// <param name="range"></param>
        private void CheckRange(Excel.Range range)
        {
            //Ist im Zellen Cache überhaupt ein Shape vorhanden?
            if (cellCache.CellShapes?.Count > 0)
            {
                //Durch alle Zellen der Range laufen
                foreach (Excel.Range cell in range.Cells)
                {
                    //Ist im Zellen Cache irgend eine Adresse welche der Adresse der Zelle entspricht?
                    if (cellCache.CellShapes.Any(x => x.Address.Equals(cell.Address)))
                    {
                        //Die geänderte Zelle ermitteln
                        var cllShp = cellCache.CellShapes.First(x => x.Address.Equals(cell.Address));
                        //Wurde der Inhalt geändert
                        if (!cllShp?.Value?.Equals(cell.Value))
                        {
                            /*
                             * Ja hier müssen wir einen neuen Shape erstellen, die Formatierung, Wert und Position
                             * des alten übernehmen und diesen im Cache und im Tabellenblatt ersetzen.
                             * Am Ende wird der alte Shape gelöscht.
                             */
                            //System.Windows.Forms.MessageBox.Show($"Die Zelle {cell.Address} wurde geändert.\nAktueller Worksheet: {range.Worksheet.Name}\nAlter Wert: {cllShp.Value}\nNeuer Wert: {cell.Value}");
                            try
                            {
                                //Hier picken wir die Formatierungsoptionen des alten Shapes auf
                                cllShp.Shape.PickUp();
                                //Dann erzeugen wir einen neuen Shape
                                CellShape newShape = CellShape.AddShape(cell.Value?.ToString(), cell, false);
                                //Wenn der nicht null ist
                                if (newShape?.Shape != null)
                                {
                                    //dann übernehmen wir Position und Größe des alten
                                    newShape.Shape.Top = cllShp.Shape.Top;
                                    newShape.Shape.Left = cllShp.Shape.Left;
                                    newShape.Shape.Width = cllShp.Shape.Width;
                                    newShape.Shape.Height = cllShp.Shape.Height;

                                    //Und wir ersetzen die Formatierungsoptionen durch die Optionen welche oben 
                                    //gepickt wurden
                                    newShape.Shape.Apply();
                                    //Dann fügen wir den neuen Shape unserem Cache hinzu
                                    cellCache.CellShapes.Add(newShape);
                                }

                                //Wir löschen den alten Shape
                                cllShp.Shape.Delete();
                                //Der Cache wird trotzdem um den alten Cacheeintrag erleichtert.
                                cellCache.CellShapes.Remove(cllShp);
                            }
                            catch (Exception ex)
                            {
                                System.Windows.Forms.MessageBox.Show(
                                    $"Error on replace a barcode in the worksheet\n\n{ex.Message}"
                                    , "Error in OBIforExcel"
                                    , System.Windows.Forms.MessageBoxButtons.OK
                                    , System.Windows.Forms.MessageBoxIcon.Error);
                            }
                        }
                    }
                }
            }
        }

        /// <summary>
        /// Gibt den aktiven Worksheet zurück. Wenn keiner ausgewählt ist, wird null zurückgegeben.
        /// </summary>
        /// <returns></returns>
        public Excel.Worksheet GetActiveWorksheet()
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
        /// null zurück gibt.
        /// </summary>
        /// <returns></returns>
        public Excel.Range GetCurrentSelection()
        {
            return ((Excel.Range)Application.Selection).Cells;
        }

        /// <summary>
        /// Fügt einem Bereich Bilder des Barcodes hinzu
        /// </summary>
        /// <param name="range"></param>
        /// <param name="fitToCell"></param>
        public void AddPictures(Excel.Range range, bool fitToCell = false)
        {
            //Wir laufen durch jede Zelle der Range
            foreach (Excel.Range item in range.Cells)
            {
                //und fügen das Bild ein, falls möglich
                AddPicture(item, fitToCell);
            }
        }

        /// <summary>
        /// Fügt einer einzigen Zelle den Barcode hinzu, wenn mehr als eine Zelle in der Range ist, dann wird nichts getan
        /// </summary>
        /// <param name="range"></param>
        /// <param name="fitToCell"></param>
        public void AddPicture(Excel.Range range, bool fitToCell = false)
        {
            //Mehr als eine Zelle? dann weg hier
            if (range?.Cells?.Count > 1)
                return;

            //Den Inhalt der Zelle auf null prüfen
            var value = range.Value?.ToString();

            //Der Wert der Zelle darf nicht null oder leer sein.
            if (!string.IsNullOrEmpty(value))
            {
                try
                {
                    //Es existiert kein Zellen Cache, dann instanziiere einen
                    if (cellCache == null)
                    {
                        cellCache = new CellCache.CellCache(range.Worksheet);
                    }
                    var cellShape = CellShape.AddShape(value, range, fitToCell);
                    //Füge den Zellen Shape nur hinzu, wenn diese nicht null ist
                    if (cellShape != null)
                    {
                        cellCache.CellShapes.Add(cellShape);
                    }
                }
                catch (Exception ex)
                {
                    System.Windows.Forms.MessageBox.Show(
                        ex.Message
                        , "Error on creating barcode image"
                        , System.Windows.Forms.MessageBoxButtons.OK
                        , System.Windows.Forms.MessageBoxIcon.Error);
                }
            }
            else
            {
                System.Windows.Forms.MessageBox.Show("A value of a cell can't be empty!"
                    , "Error on creating barcode image"
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
