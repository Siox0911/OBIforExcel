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
    internal class CellCache
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
        private CellCache() { }

        /// <summary>
        /// Konstruktor der den Zellen Cache im angegebenen Worksheet erstellt.
        /// </summary>
        /// <param name="worksheet"></param>
        internal CellCache(Excel.Worksheet worksheet)
        {
            SetFelder(CreateCellCache(worksheet));
        }

        /// <summary>
        /// Setzt die Felder in der eigenen Instanz
        /// </summary>
        /// <param name="cellCache">der zu setzende CellCache</param>
        private void SetFelder(CellCache cellCache)
        {
            CellShapes = cellCache?.CellShapes;
            WorksheetName = cellCache?.WorksheetName;
        }

        /// <summary>
        /// Fügt den Zellen Shape der aktuellen Liste hinzu, soweit der <paramref name="cellShape"/> 
        /// ungleich null ist         /// und die eigene Liste <seealso cref="CellShapes"/> ungleich null ist.
        /// </summary>
        /// <param name="cellShape"></param>
        internal void AddCellShape(CellShape cellShape)
        {
            if(cellShape != null && CellShapes != null)
            {
                CellShapes.Add(cellShape);
            }
        }

        /// <summary>
        /// Versucht im Worksheet die Shapes mit den Barcodes zu finden. Dabei wird wenigstens ein Zellen 
        /// Cache mit einer leeren Liste an Zellen Shapes erstellt.
        /// </summary>
        /// <param name="worksheet"></param>
        /// <returns></returns>
        internal static CellCache CreateCellCache(Excel.Worksheet worksheet)
        {
            //Ist der Sheet vorhanden und sichtbar
            if (worksheet?.Visible == Excel.XlSheetVisibility.xlSheetVisible)
            {
                //Neuen Zellen Cache erstellen
                CellCache cellCache = new CellCache
                {
                    CellShapes = new List<CellShape>(),
                    WorksheetName = worksheet.Name
                };

                //Durch alle Shapes im Worksheet laufen
                foreach (Excel.Shape shape in worksheet.Shapes)
                {
                    //Prüft den Shape auf Name, null usw.
                    if (CellShape.CheckShape(shape))
                    {
                        //Aktuelle Adresse
                        var addressCurShape = CellShape.GetAddress(shape);
                        //Aktueller Wert
                        var cellValue = worksheet.Range[addressCurShape]?.Value;
                        //Wert und Adresse auf null prüfen
                        if (addressCurShape != null && cellValue != null)
                        {
                            //Adresse, aktueller Wert und Shape im Cache ablegen
                            cellCache.CellShapes.Add(new CellShape { Shape = shape, Address = addressCurShape, Value = cellValue});
                        }
                    }
                }
                //Es wird mindestens ein instanziierter Zellen Cache mit einer leeren instanziierten Liste zurückgegeben.
                return cellCache;
            }
            //Es gibt eigentlich nichts, aber instanziiere wenigestens den Zellen Cache mit einer leeren Liste aus Zellen Shapes.
            return new CellCache { CellShapes = new List<CellShape>() };
        }
    }
}
