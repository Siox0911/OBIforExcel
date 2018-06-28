using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Excel = Microsoft.Office.Interop.Excel;
using Office = Microsoft.Office.Core;
using DataMatrix.net;

namespace OBIforExcel.CellCache
{
    /// <summary>
    /// Eine Klasse welche eine Zelladdresse mit einem Shape verheitet oder darstellt.
    /// </summary>
    internal class CellShape
    {
        /// <summary>
        /// Das Shape für den Barcode
        /// </summary>
        internal Excel.Shape Shape { get; set; }

        /// <summary>
        /// Die ZellAdresse auf der das Shape basiert
        /// </summary>
        internal string Address { get; set; }

        /// <summary>
        /// Der Wert der Zelle, die auf das Shape referenziert
        /// </summary>
        internal string Value { get; set; }

        /// <summary>
        /// Erzeugt einen Barcodeshape an der Position der Zelle <paramref name="range"/> mit dem Inhalt 
        /// <paramref name="text"/> und passt die Größe an die Zelle an <paramref name="fitToCell"/>
        /// </summary>
        /// <param name="text">Text der in den Barcode geschrieben wird</param>
        /// <param name="range">Die Zelle an der dieser Shape gebunden wird und an dessen stelle der Shape 
        /// platziert wird</param>
        /// <param name="fitToCell">Soll das Shape an die Zellgröße angepasst werden.</param>
        /// <returns></returns>
        internal static CellShape AddShape(string text, Excel.Range range, bool fitToCell = false)
        {
            //Nur wenn der text nicht leer ist und die Zellen maximal 1 Zelle ist
            if (!string.IsNullOrEmpty(text) && range?.Cells.Count == 1)
            {
                //Inhalt der Zelle prüfen
                var value = range.Value?.ToString();
                if (!string.IsNullOrEmpty(value))
                {
                    try
                    {
                        //Neuen DataMatrix Encoder erstellen
                        var dmtxImageEncoder = new DmtxImageEncoder();
                        var dmtxImageEncoderOptions = new DmtxImageEncoderOptions
                        {
                            BackColor = System.Drawing.Color.White,
                            ForeColor = System.Drawing.Color.Black
                        };

                        //Encodierung einleiten
                        var img = dmtxImageEncoder.EncodeImage(value, dmtxImageEncoderOptions);

                        //Irgendwo in Temp eine Datei erzeugen, Name egal
                        var fName = System.IO.Path.GetTempFileName() + ".jpg";

                        //In diese TempDatei speichern
                        img.Save(fName);

                        //Position der Zelle
                        System.Drawing.Point point = GetPointOfCell(range);

                        //Neue Größe ermitteln
                        System.Drawing.Size size = GetPictureOrCellSize(fName, range, fitToCell);

                        /*
                         * Erstellt ein Bild und speichert es an der Position der Zelle.
                         * Dem Shape wird die Zelle angehangen, die für den Inhalt verantwortlich ist.
                         */
                        var shape = Globals
                            .ThisAddIn
                            .GetActiveWorksheet()
                            .Shapes
                            .AddPicture(
                                fName,
                                Office.MsoTriState.msoFalse,
                                Office.MsoTriState.msoCTrue,
                                point.X,
                                point.Y,
                                size.Width,
                                size.Height);
                        //Wir speichern den Namen des Shape in der Form Barcode($A$1), so können wir später 
                        //aus jedem Tabellenblatt die Verlinkung des Barcodes zur Ursprungszelle wieder herleiten
                        shape.Name = $"Barcode({range.Address})";

                        //Das CellShape zurückgeben
                        return new CellShape
                        {
                            Shape = shape,
                            Address = range.Address,
                            Value = value
                        };
                    }
                    catch (Exception ex)
                    {
                        System.Windows.Forms.MessageBox.Show(
                            $"An Error on the barcode creation was thrown, see the message below:\n\n{ex.Message}"
                            , "Error on creating barcode image"
                            , System.Windows.Forms.MessageBoxButtons.OK
                            , System.Windows.Forms.MessageBoxIcon.Error);
                    }
                }
            }
            return null;
        }

        /// <summary>
        /// Wenn der Shape ein Shape mit Barcode ist, wird die verbundene Adresse des Shape zurück gegeben.
        /// </summary>
        /// <param name="shape"></param>
        /// <returns></returns>
        internal static string GetAddress(Excel.Shape shape)
        {
            //Prüfe den Shape
            if (CheckShape(shape))
            {
                //Adresse ermitteln, also z.B. Barcode($A$1), dann wird alles zwischen den Klammern zurückgegeben
                var addressCurShape = shape.Name.Split('(')?[1].Split(')')?[0];
                //Adresse zurückgeben
                if (addressCurShape != null)
                {
                    return addressCurShape;
                }
            }
            return null;
        }

        internal static bool CheckShape(Excel.Shape shape)
        {
            //Shape auf null prüfen
            if (shape != null)
            {
                //Name des Shapes
                var name = shape.Name;
                //Wenn der Shapename Barcode enthält und nicht null ist
                if (name?.IndexOf("barcode", StringComparison.OrdinalIgnoreCase) != -1)
                {
                    //Gib true zurück
                    return true;
                }
            }
            //Sonst false
            return false;
        }

        /// <summary>
        /// Ermittelt die Position einer einzigen Zelle. Sollte die Range mehr als eine Zelle enthalten, wird ein Fehler zurückgegeben.
        /// </summary>
        /// <param name="range">Die Range mit einer einzigen Zelle</param>
        /// <returns></returns>
        /// <exception cref="NotSupportedException">Wird ausgelöst, wenn die Range mehr als eine Zelle enthält</exception>
        private static System.Drawing.Point GetPointOfCell(Excel.Range range)
        {
            if (range.Cells.Count > 1)
            {
                throw new NotSupportedException($"This isn't a user issue, it's a problem on the Addin OBIforExcel." +
                    $"\n\nOnly one cell is allowed to get the position of it. Cells get: {range.Cells.Count}");
            }

            return new System.Drawing.Point((int)((double)range.Left), (int)((double)range.Top));
        }

        /// <summary>
        /// Ermittelt die Größe des Bildes und gibt die Größe zurück.
        /// </summary>
        /// <param name="pathWithPicture">Der Pfad zur Bilddatei</param>
        /// <param name="range">Die Zelle, welche als Größenreferenz dient</param>
        /// <param name="fitToCell">Soll das Bild an die Zellgröße angepasst werden</param>
        /// <returns></returns>
        /// <exception cref="NullReferenceException">Wenn das Bild leer ist</exception>
        private static System.Drawing.Size GetPictureOrCellSize(string pathWithPicture, Excel.Range range, bool fitToCell)
        {
            //Ist der Pfad mit dem Bild überhaupt gefüllt und existiert die Datei auch?

            /*
             * Wir müssen hier eigentlich nicht prüfen, da wir dies immer nur als privat einsetzen und 
             * es eiegentlich klar sein sollte, ob dies funktioniert oder nicht.
             * Aber es ist besser wir prüfen, weil es nicht ausgeschlossen ist,
             * das die Funktion System.IO.Path.GetTempFileName() und das danach folgende 
             * Speichern in der Methode AddShape auch funktioniert hat.
             */
            if (!string.IsNullOrEmpty(pathWithPicture) && System.IO.File.Exists(pathWithPicture))
            {
                var size = new System.Drawing.Size();
                //Die Größe wird jetzt an die Zelle angepasst oder an das Bild
                if (fitToCell)
                {
                    size.Height = (int)((double)range.Height);
                    size.Width = (int)((double)range.Width);
                }
                else
                {
                    //Größe des Bildes ermitteln
                    var bitMap = new System.Drawing.Bitmap(pathWithPicture);
                    size.Height = bitMap.Height;
                    size.Width = bitMap.Width;

                    //Wichtig, dass Bild wieder auf null setzen, da es sonst als verwendet markiert ist, falls 
                    //es jemand bearbeiten oder löschen möchte usw.
                    bitMap = null;
                }

                return size;
            }

            throw new NullReferenceException($"This isn't a user issue, it's a problem on the Addin OBIforExcel." +
                $"\n\nThe path of the picture is null or the path doesn't exist. Maybe some rights are forbidden " +
                $"in the file system.\n\nPath check failed: \"{pathWithPicture}\"");
        }
    }
}
