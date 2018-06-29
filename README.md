# OBIforExcel
Office Barcode Integration for Excel is a tool to integrate barcodes to a Excel workbook.

This is a solution without invest some money for private use. 
Please check the LICENSE for commercial use.

This AddIn is using the OpenSource Project DataMatrix.Net to create DataMatrix codes.

Maybe this project does'nt work on Excel 2010. If it not run, please check the `ThisAddIn_Startup`
Methode in file `ThisAddIn.cs` and uncomment the line:
```C#
//WorkbookActivate(this.Application.ThisWorkbook);
```

Office Barcode Integration for Excel ist ein Werkzeug zur Erzeugung von Barcodes in einer
Excel Arbeitsmappe. 

Diese Lösung ist kostenlos und enthält keine Zahlungen für die private Nutzung.
Bitte prüfen Sie die LICENSE zur kommerziellen Verwendung.

Dieses AddIn verwendet das OpenSource Projekt DataMatrix.Net um die DataMatrix Codes zu erstellen.

Vielleicht funktioniert das AddIn nicht mit Excel 2010, dann schaue in die Methode `ThisAddIn_Startup`
in der Datei `ThisAddIn.cs` und unkommentiere diese Zeile:
```C#
//WorkbookActivate(this.Application.ThisWorkbook);
```
