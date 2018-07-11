# Changelog

##### _Note the current inline code documentation is in german_

- 1.1.0.6
  - reworked: the complete Cache was reworked to handle the workbook
  - reworked: the complete Cache-code was merged into the CellCache
  - removed: Events `SheetActivate`, `SheetDeaktivate`, `SheetChanged`,
`SheetCalculate` and `SheetBeforeDelete`
    - checking events for delete a sheet or add a sheet are planned
  - changed: new handle in the methode `AddIn_Startup` for Excel2010 "hope its works"
  - Add: event `AfterCalculate` to handle a cell value is changed or calculated
    - a problem with this is a manual recalculating of the workbook;
the DataMatrix changes its value only, when a manual recalculation is performed
  - Add: the cell size can be automatically adjusted to the barcode size
  - Add: the barcode can be tied to the position and size of the cell

- 1.0.0.4
  - add: Event SheetBeforeDelete, to handle a sheet is deleted by a user 
  - changed: add some null checks to avoid Excel Errors
  - fixed: rebuild a function which is called after the sheet is recalculated 

- 1.0.0.2
  - changed: editorial changes only for simpler code reading on github

- 1.0.0.1
  - first version, initial release
  - simple CellCache implemented to link a cell to a picture
  - RibbonMenu to create a DataMatrix code as picture to a Excel-Sheet 