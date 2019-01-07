using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Excel = Microsoft.Office.Interop.Excel;
using Office = Microsoft.Office.Core;

namespace OBIforExcel.CellCache
{
    /// <summary>
    /// Diese Klasse cached das komplette Workbook. Sie enthält für jedes Workbook das entsprechende
    /// SheetCache, welches wiederum die entsprechenden CellShapes enthält.
    /// </summary>
    internal class WorkbookCache
    {
        /// <summary>
        /// Der Name des Workbooks
        /// </summary>
        internal string WorkbookName { get; private set; }

        /// <summary>
        /// Die Liste mit den SheetCaches. 
        /// </summary>
        internal List<SheetCache> SheetCaches { get; private set; }

        /// <summary>
        /// Der Standard Konstruktor soll von außen nicht erreichbar sein.
        /// Muss aber existieren, weil wir ihn intern brauchen.
        /// </summary>
        private WorkbookCache() { }

        /// <summary>
        /// Erstellt einen kompletten Cache über das komplette Workbook.
        /// Dieser ist mindestens leer und die darin enthaltenen SheetCaches
        /// können von außen nicht verändert werden.
        /// </summary>
        /// <param name="workbook"></param>
        internal WorkbookCache(Excel.Workbook workbook)
        {
            SetFelder(CreateWorkbookCache(workbook));
        }

        /// <summary>
        /// Setzt die Felder des Workbook Caches in der eigenen Instanz.
        /// </summary>
        /// <param name="workbookCache"></param>
        private void SetFelder(WorkbookCache workbookCache)
        {
            WorkbookName = workbookCache.WorkbookName;
            SheetCaches = workbookCache.SheetCaches;
        }

        /// <summary>
        /// Prüft das komplette Workbook auf Änderungen und ändert die entsprechenden Barcodes. Es werden keine neuen Barcodes gefunden.
        /// </summary>
        internal void CheckWorkbook()
        {
            //Wir wollen hier nur durch, wenn der WorkbookName gesetzt ist und wir auch SheetCaches haben.
            if (!string.IsNullOrEmpty(WorkbookName) && SheetCaches?.Count > 0)
            {
                //Das Workbook finden, welches zu diesem Cache gehört
                foreach (Excel.Workbook wb in Globals.ThisAddIn.Application.Workbooks)
                {
                    //Das Workbook muss den selben Namen haben wie das hier.
                    //Am Ende dieses If ist ein return, weil wir nicht weiter suchen müssen
                    if (wb.Name.Equals(WorkbookName))
                    {
                        //Wir haben das Workbook gefunden.
                        //Nun durchlaufen wir alle Sheets
                        foreach (Excel.Worksheet sheet in wb.Sheets)
                        {
                            //Wir prüfen ob sich der Sheet in unseren Cache befindet
                            if (HasSheetCache(sheet))
                            {
                                //Dann ermitteln wir den SheetCache
                                var sheetCache = SheetCaches.First(y => y.WorksheetName.Equals(sheet.Name));
                                //Und prüfen ihn auf Änderungen
                                sheetCache.CheckWorksheet(sheet);
                            }
                        }

                        //Aus der Funktion ausbrechen, wir sind fertig
                        return;
                    }
                }
            }
        }

        /// <summary>
        /// Prüft einen Worksheet auf Änderungen, werden diese vorgefunden werden diese Änderungen in den Barcodes sichtbar.
        /// Es werden keine neuen Barcodes gefunden.
        /// </summary>
        /// <param name="worksheet"></param>
        internal void CheckWorkSheet(Excel.Worksheet worksheet)
        {
            if (worksheet != null)
            {
                //SheetCache ermitteln
                var shCache = GetSheetCache(worksheet);
                //Den SheetCache prüfen, wenn er nicht null ist.
                shCache?.CheckWorksheet(worksheet);
            }
        }

        /// <summary>
        /// Prüft eine Range auf Änderungen, werden diese vorgefunden werden diese Änderungen in den Barcodes sichtbar.
        /// Es werden keine neuen Barcodes gefunden.
        /// </summary>
        /// <param name="range"></param>
        internal void CheckRange(Excel.Range range)
        {
            if (range != null)
            {
                var shCache = GetSheetCache(range.Worksheet);
                shCache?.CheckRange(range);
            }
        }

        /// <summary>
        /// Fügt einen SheetCache dem Workbook hinzu
        /// </summary>
        /// <param name="sheetCache"></param>
        internal void AddSheetCache(SheetCache sheetCache)
        {
            if (SheetCaches == null)
            {
                SheetCaches = new List<SheetCache>();
            }
            if (sheetCache != null)
            {
                SheetCaches.Add(sheetCache);
            }
        }

        /// <summary>
        /// Gibt den passenden SheetCache zurück, der den WorkSheet cached oder null wenn kein Cache gefunden wurde.
        /// </summary>
        /// <param name="worksheet"></param>
        /// <returns></returns>
        internal SheetCache GetSheetCache(Excel.Worksheet worksheet)
        {
            if (HasSheetCache(worksheet))
            {
                return SheetCaches.First(x => x.WorksheetName.Equals(worksheet.Name));
            }
            return null;
        }

        /// <summary>
        /// Ist der Worksheet hier im Cache vertreten?
        /// </summary>
        /// <param name="worksheet"></param>
        /// <returns></returns>
        internal bool HasSheetCache(Excel.Worksheet worksheet)
        {
            return SheetCaches?.Any(x => x.WorksheetName.Equals(worksheet.Name)) == true;
        }

        /// <summary>
        /// Fügt diesem WorkbookCache ein CellShape hinzu. Dieser wird als Barcode auf dem Worksheet angezeigt.
        /// Erzeugt also einen Barcode im <paramref name="workbook"/> und der <paramref name="range"/> 
        /// und passt diesen an die Zelle an <paramref name="fitToCell"/>.
        /// </summary>
        /// <param name="workbook">In diesem Workbook soll der Barcode hinzugefügt werden.</param>
        /// <param name="range">Ein Bereich an Zellen. Die Zellen sollten Inhalte besitzen.</param>
        /// <param name="xlPlacement">Wie soll der Barcode an die Zelle gebunden sein.</param>
        /// <param name="fitToCell">Soll der Barcode an die Größe der Zelle angepasst werden.</param>
        /// <param name="cellFitToPicture">Soll die Zelle an die Größe des Bildes angepasst werden.</param>
        internal void AddCellShape(Excel.Workbook workbook, Excel.Range range, Excel.XlPlacement xlPlacement = Excel.XlPlacement.xlMove, bool fitToCell = false, bool cellFitToPicture = false)
        {
            try
            {
                if (workbook?.Name.Equals(WorkbookName) == true && range != null)
                {
                    //Existiert bereits ein SheetCache zu dem Worksheet
                    if (!HasSheetCache(range.Worksheet))
                    {
                        //Nein existiert nicht
                        //Neuen SheetCache hinzufügen
                        var shCch = SheetCache.CreateSheetCache(range.Worksheet);
                        if (shCch != null)
                        {
                            AddSheetCache(shCch);
                        }
                    }

                    //Den SheetCache ermitteln
                    var shCache = GetSheetCache(range.Worksheet);

                    if (shCache != null)
                    {
                        //Durch alle Zellen in der range laufen
                        foreach (Excel.Range cell in range)
                        {
                            //Jetzt den CellShape in den SheetCache schreiben, dabei wird er auch auf der Mappe erzeugt
                            shCache.AddCellShape(CellShape.AddShape(cell, xlPlacement, fitToCell, cellFitToPicture));
                        }
                    }
                }
            }
            catch (Exception)
            {
                throw;
            }
        }

        /// <summary>
        /// Erstellt einen neuen Workbook Cache.
        /// Es wird mindestens ein leerer WorkbookCache zurückgegeben. 
        /// Die darin enthaltenen SheetCaches können von außerhalb nicht gefüllt werden.
        /// </summary>
        /// <param name="workbook"></param>
        /// <returns></returns>
        internal static WorkbookCache CreateWorkbookCache(Excel.Workbook workbook)
        {
            try
            {
                if (workbook != null)
                {
                    //Neuen WorkbookCache erzeugen
                    var wbCache = new WorkbookCache
                    {
                        //Name und SheetCache initialisieren
                        WorkbookName = workbook.Name,
                        SheetCaches = new List<SheetCache>()
                    };

                    //Durch alle Sheets laufen und die SheetCaches erzeugen
                    foreach (Excel.Worksheet sheet in workbook.Sheets)
                    {
                        wbCache.SheetCaches.Add(SheetCache.CreateSheetCache(sheet));
                    }

                    return wbCache;
                }
                return null;
            }
            catch (Exception)
            {
                throw;
            }
        }
    }
}
