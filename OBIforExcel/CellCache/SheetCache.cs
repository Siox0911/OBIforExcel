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
    /// Klasse welche den internen Zellen Cache darstellt.
    /// Der Zellen Cache arbeitet immer nur im aktuellen Worksheet.
    /// </summary>
    internal class SheetCache
    {
        /// <summary>
        /// Eine Liste der Zellen, welche einen Shape enthalten
        /// </summary>
        internal List<CellShape> CellShapes { get; private set; }

        /// <summary>
        /// Der aktuelle Name des Worksheets
        /// </summary>
        internal string WorksheetName { get; private set; }

        /// <summary>
        /// Privater Konstruktor, dieser soll nur aus dieser Klasse aufrufbar sein.
        /// </summary>
        private SheetCache() { }

        /// <summary>
        /// Konstruktor der den Zellen Cache im angegebenen Worksheet erstellt. Dabei wird wenigstens ein 
        /// SheetCache mit einer leeren Liste an Zellen Shapes erstellt.
        /// </summary>
        /// <param name="worksheet"></param>
        internal SheetCache(Excel.Worksheet worksheet)
        {
            SetFelder(CreateSheetCache(worksheet));
        }

        /// <summary>
        /// Setzt die Felder in der eigenen Instanz
        /// </summary>
        /// <param name="cellCache">der zu setzende CellCache</param>
        private void SetFelder(SheetCache cellCache)
        {
            CellShapes = cellCache?.CellShapes;
            WorksheetName = cellCache?.WorksheetName;
        }

        /// <summary>
        /// Fügt den Zellen Shape der aktuellen Liste hinzu.
        /// </summary>
        /// <param name="cellShape"></param>
        internal void AddCellShape(CellShape cellShape)
        {
            if (CellShapes == null)
            {
                CellShapes = new List<CellShape>();
            }

            if (cellShape != null)
            {
                CellShapes.Add(cellShape);
            }
        }

        /// <summary>
        /// Prüft das Worksheet auf Änderungen in den überwachten <seealso cref="CellShapes"/>.
        /// Sind Änderungen vorhanden, werden diese direkt umgesetzt. Es werden keine neuen 
        /// Barcodes gefunden.
        /// </summary>
        /// <param name="worksheet"></param>
        internal void CheckWorksheet(Excel.Worksheet worksheet)
        {
            if (!worksheet.Name.Equals(WorksheetName))
            {
                throw new ArgumentException($"Worksheet check failed: Worksheet.Name '{worksheet.Name}' is not the required name '{WorksheetName}'");
            }

            //Das machen wir aber nur, wenn der Worksheetname mit dem Namen im Zellen Cache übereinstimmt.
            //Weil wir nicht wissen was in den anderen Tabellenblättern so ist, da dies im Cache nicht abgebildet ist
            if (CellShapes?.Count > 0)
            {
                //Array aus unserem Zellen Cache erstellen
                //Müssen wir machen, weil CheckRange verändert den Cache und das führt 
                //in der For Schleife zu einer Exception weil sich die Auflistung geändert hat.
                var rng = new Excel.Range[CellShapes.Count];
                //Wir laufen durch unseren Zellen Cache
                for (int i = 0; i < CellShapes.Count; i++)
                {
                    //Dem Array hinzufügen
                    rng[i] = worksheet.Range[CellShapes[i].Address];
                }

                //Jetzt wird jeder Eintrag im Cache geprüft
                foreach (var item in rng)
                {
                    CheckRange(item);
                }
            }
        }

        /// <summary>
        /// Prüft <paramref name="range"/> auf Änderungen zu den überwachten <seealso cref="CellShapes"/>.
        /// Sind Änderungen vorhanden, werden diese direkt umgesetzt. Es werden keine neuen Barcodes gefunden.
        /// </summary>
        /// <param name="range"></param>
        internal void CheckRange(Excel.Range range)
        {
            if (!range.Worksheet.Name.Equals(WorksheetName))
            {
                throw new ArgumentException($"Worksheet range check failed: Worksheet.Name '{range.Worksheet.Name}' is not the required name '{WorksheetName}'");
            }
            //Ist im Zellen Cache überhaupt ein Shape vorhanden?
            if (CellShapes?.Count > 0)
            {
                //Durch alle Zellen der Range laufen
                foreach (Excel.Range cell in range.Cells)
                {
                    //Ist im Zellen Cache irgend eine Adresse welche der Adresse der Zelle entspricht?
                    if (CellShapes.Any(x => x.Address.Equals(cell.Address)))
                    {
                        //Die geänderte Zelle ermitteln
                        var cllShp = CellShapes.First(x => x.Address.Equals(cell.Address));
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
                                //Wenn der alte Shape != null ist
                                if (cllShp.Shape != null)
                                {
                                    //Hier picken wir die Formatierungsoptionen des alten Shapes auf
                                    //Das kann zu einem Fehler führen, weil z.B. das Bild gelöscht wurde
                                    try
                                    {
                                        cllShp.Shape.PickUp();
                                    }
                                    catch (Exception)
                                    {
                                        //Dann löschen wir dies aus dem Cache und
                                        //lassen die nächste Zelle prüfen
                                        CellShapes.Remove(cllShp);
                                        continue;
                                    }

                                    //Dann erzeugen wir einen neuen Shape
                                    CellShape newShape = CellShape.AddShape(cell, cllShp.Shape.Placement, false, false);
                                    //Wenn der nicht null ist
                                    //Diese Prüfung bewirkt, dass bei einem Löschen des Zelleninhaltes auch der Barcode gelöscht wird
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
                                        CellShapes.Add(newShape);
                                    }

                                    //Wir löschen den alten Shape
                                    cllShp.Shape.Delete();
                                }

                                //Der Cache wird trotzdem um den alten Cacheeintrag erleichtert.
                                CellShapes.Remove(cllShp);
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
        /// Versucht im Worksheet die Shapes mit den Barcodes zu finden. Dabei wird wenigstens ein Zellen 
        /// Cache mit einer leeren Liste an Zellen Shapes erstellt. Wenn der Worksheet null ist, wird null
        /// zurückgegeben.
        /// </summary>
        /// <param name="worksheet"></param>
        /// <returns></returns>
        internal static SheetCache CreateSheetCache(Excel.Worksheet worksheet)
        {
            //Ist der Sheet vorhanden
            if (worksheet != null)
            {
                //Neuen Zellen Cache erstellen
                SheetCache cellCache = new SheetCache
                {
                    CellShapes = new List<CellShape>(),
                    WorksheetName = worksheet.Name
                };

                //Durch alle Shapes im Worksheet laufen
                foreach (Excel.Shape shape in worksheet.Shapes)
                {
                    //Prüft den Shape auf Name, null usw.
                    if (CellShape.CheckShapeName(shape))
                    {
                        //Aktuelle Adresse
                        var addressCurShape = CellShape.GetAddress(shape);
                        //Aktueller Wert
                        var cellValue = worksheet.Range[addressCurShape]?.Value;
                        //Wert und Adresse auf null prüfen
                        if (addressCurShape != null && cellValue != null)
                        {
                            //Adresse, aktueller Wert und Shape im Cache ablegen
                            cellCache.CellShapes.Add(new CellShape { Shape = shape, Address = addressCurShape, Value = cellValue });
                        }
                    }
                }
                //Es wird mindestens ein instanziierter Zellen Cache mit einer leeren instanziierten Liste zurückgegeben.
                return cellCache;
            }
            //Es gibt nichts
            return null;
        }

        /// <summary>
        /// Prüft den Worksheet im angegebenen WorkbookCache auf Änderungen. Es werden keine neuen Barcodes gefunden.
        /// </summary>
        /// <param name="workbookCache"></param>
        /// <param name="worksheet"></param>
        internal static void CheckWorksheet(WorkbookCache workbookCache, Excel.Worksheet worksheet)
        {
            //Nur prüfen wenn der SheetCache auch existiert
            if (worksheet != null && workbookCache?.HasSheetCache(worksheet) == true)
            {
                //SheetCache ermitteln
                var shCache = workbookCache.SheetCaches.First(x => x.WorksheetName.Equals(worksheet.Name));
                //Prüfen
                shCache.CheckWorksheet(worksheet);
            }
        }
    }
}
