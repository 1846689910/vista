# vista

###### VERSION 1.1.3 CLASS

## introduction

vista is a helper class module for excel VBA use. It could assit you in repeating and complex worksheet manipulation.

## How does it work?

vista will perform as a worksheet wrapper for a specific worksheet and provide many useful functions and subs to interact with worksheet.

1. **please download `Vista.cls` from [Vista](https://github.com/1846689910/vista/releases/download/v1.1.3/vista.cls)**
2. **use `import file` to import the module in your VBA project**

## initialization

```vb
Option Explicit
Sub main()
  Dim vws As New Vista ' create a new vista wrapper
  Dim startCell As Variant
  vws.init Sheet1 ' initialize the vista wrapper instance with worksheet
  startCell = vws.getFirstNonEmptyCell()
End Sub
```

## API

###### return `Void` means a `Sub`

### Methods

#### Critical

| Return    | Method                                                                                                                                                                                                                                     |
| --------- | :----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------- |
| `Void`    | **`init(Worksheet ws)`** <br/>&bull; initialize the created `Vista` instance and hook the instance with the specified Worksheet `ws`                                                                                                       |
| `Void`    | **`sync()`**<br/>&bull; manually trigger properties update for the worksheet, include `lastRow` and `lastCol`<br/>&bull; should be used especially after the worksheet was directly manipulated by user insteadof operated through `vista` |
| `Variant` | **`getFirstNonEmptyCell()`**<br/>&bull; return `Array(row as Long, col as Long)` to represent the position of the first non empty cell in worksheet<br/>&bull; if found, return `Array(row, col)`, otherwise return `Array(-1, -1)`        |

#### Row

| Return    | Method                                                                                                                                                                                                                                                                                                                                                                                                                                                                                    |
| --------- | :---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------- |
| `Long`    | **`getLastRow()`**<br/>&bull; get the latest last row number of the worksheet                                                                                                                                                                                                                                                                                                                                                                                                             |
| `Void`    | **`removeRow(Long i)`**<br/>&bull; remove the i th row                                                                                                                                                                                                                                                                                                                                                                                                                                    |
| `Void`    | **`addRow(Long r)`**<br/>&bull; insert a new row at `r` th row, new row is `r` th row                                                                                                                                                                                                                                                                                                                                                                                                     |
| `Variant` | **`rowIndexOf(Long searchRow, Long startCol, String target, Optional Boolean exactMatch = False)`**<br/>&bull; return `Array(row as Long, col as Long)` to represent the position of the first cell found in `searchRow` and searched from `startCol` that contains content `target`<br/>&bull; if `exactMatch` then cell content should be exactly equal to `target`, otherwise cell content should contain `target`<br/>&bull; if not found, return `Array(-1, -1)`                     |
| `Object`  | **`rowIndicesOf(Long searchRow, Long startCol, String target, Optional Boolean exactMatch = False)`**<br/>&bull; return `ArrayList<Array(row as Long, col as Long)>` to represent a list of position arrays of the cell found in `searchRow` and searched from `startCol` that contains content `target`<br/>&bull; if `exactMatch` then cell content should be exactly equal to `target`, otherwise cell content should contain `target`<br/>&bull; if not found, return empty ArrayList |

#### Column

| Return    | Method                                                                                                                                                                                                                                                                                                                                                                                                                                                                                  |
| --------- | :-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------- |
| `Long`    | **`getLastCol()`**<br/>&bull; get the latest last column number of the worksheet                                                                                                                                                                                                                                                                                                                                                                                                        |
| `Void`    | **`removeCol(Long i)`**<br/>&bull; remove the i th column                                                                                                                                                                                                                                                                                                                                                                                                                               |
| `Void`    | **`addCol(Long c)`**<br/>&bull; insert a new column at `c` th column, the new column is `c` th column                                                                                                                                                                                                                                                                                                                                                                                   |
| `Variant` | **`colIndexOf(Long searchCol, Long startRow, String target, Optional Boolean exactMatch = False)`**<br/>&bull; return `Array(row as Long, col as Long)` to represent the position of the first cell found in `searchCol` and searched from `startRow` that contains content `target`<br/>&bull; if `exactMatch` then cell content should be exactly equal to `target`, otherwise cell content should contain `target`<br/>&bull; if not found, return `Array(-1, -1)`                   |
| `Object`  | **`colIndexOf(Long searchCol, Long startRow, String target, Optional Boolean exactMatch = False)`**<br/>&bull; return `ArrayList<Array(row as Long, col as Long)>` to represent a list of position arrays of the cell found in `searchCol` and searched from `startRow` that contains content `target`<br/>&bull; if `exactMatch` then cell content should be exactly equal to `target`, otherwise cell content should contain `target`<br/>&bull; if not found, return empty ArrayList |

#### Worksheet

| Return    | Method                                                                                                                                                                                                                                                                              |
| --------- | :---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------- |
| `Boolean` | **`hasWorksheet(Workbook wb, String name)`**<br/>&bull; if the workbook `wb` contains a worksheet with `name`                                                                                                                                                                       |
| `Void`    | **`addWorksheet(Workbook wb, String name)`**<br/>&bull; create a worksheet with `name` in workbook `wb`                                                                                                                                                                             |
| `Void`    | **`removeWorksheet(Workbook wb, String name)`**<br/>&bull; remove a worksheet with `name` in workbook `wb`                                                                                                                                                                          |
| `Void`    | **`clearWorksheet(Optional Worksheet ws)`**<br/>&bull; clear the whole content of the specified worksheet. <br/>&bull;if worksheet is not specified, will clear the wrapped worksheet                                                                                               |
| `Void`    | **`copyRange(Worksheet wsSrc, Variant startCell, Variant endCell, Worksheet wsTarget, Variant startCellTarget)`**<br/>&bull; copy the selected range from `wsSrc` to specific position in `wsTarget`<br/>&bull; `startCell`, `endCell`, `startCellTarget` are all `Array(row, col)` |

#### Workbook

| Return     | Method                                                                                                                                                                                                  |
| ---------- | :------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------ |
| `Void`     | **`mkWorkbook(String path, Optional String TEMPLATE_PATH)`**<br/>&bull; create a workbook according to the path, path should include filename. <br/>&bull; user can also specified a template file path |
| `Workbook` | **`openWorkbook(String path)`**<br/>&bull; return the opened workbook                                                                                                                                   |
| `Void`     | **`saveasWorkbook(Workbook wb, String path)`**<br/>&bull; save the workbook `wb` in path, could also be used for rename                                                                                 |

#### Utils

| Return    | Method                                                                                                                                                                                                                                                                                                                         |
| --------- | :----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------- |
| `Long`    | **`colLetterToNum(String s)`** <br/>&bull; column letter to number                                                                                                                                                                                                                                                             |
| `String`  | **`colNumToLetter(Long n)`** <br/>&bull; column number to letter                                                                                                                                                                                                                                                               |
| `Variant` | **`getRowColFromRangeSelector(String selector)`**<br/>&bull; convert range selector to Array(row, col), like `B5` to `Array(5, 2)`                                                                                                                                                                                             |
| `String`  | **`getRangeSelectorFromRowCol(Variant rowCol)`**<br/>&bull; convert Array(row, col) to range selector, like `Array(5, 2)` to `B5`                                                                                                                                                                                              |
| `Void`    | **`mkDir(String path)`**<br/>&bull; create the directory according to `path`, recursively create non-existing sub directories                                                                                                                                                                                                  |
| `Boolean` | **`existFile(String path)`**<br/>&bull; check if the file path exists                                                                                                                                                                                                                                                          |
| `Boolean` | **`existDir(String path)`**<br/>&bull; check if the direcotory path exists                                                                                                                                                                                                                                                     |
| `String`  | **`dirname(String path)`**<br/>&bull; return the parent directory of the given `path`                                                                                                                                                                                                                                          |
| `String`  | **`basename(String path)`**<br/>&bull; return the current folder or filename in given `path`                                                                                                                                                                                                                                   |
| `Object`  | **`getAllFilenames(String path)`**<br/>&bull; return an `ArrayList<String>` of all names of files under `path` directory                                                                                                                                                                                                       |
| `Object`  | **`getAllFilePaths_R(String path, Optional Boolean needPath=false, Optional Variant level)`**<br/>&bull; return an `ArrayList<String>` of all file paths under the `path` directory and all its nested sub folders recursively<br/>&bull; if `level` is given, then only do `level` depth search. `level=0` means current path |
| `Object`  | **`getAllSubDirs(String path)`**<br/>&bull; return an `ArrayList<String>` of all names of folders under `path` directory                                                                                                                                                                                                       |
| `Object`  | **`getAllSubDirs_R(String path, Optional Boolean needPath=false, Optional Variant level)`**<br/>&bull; return an `ArrayList<String>` of all folder paths under the `path` directory and all its nest sub folders recursively<br/>&bull; if `level` is given, then only do `level` depth search. `level=0` means current path   |
| `String`  | **`openFileDialog(Optional Variant extensions, Optional String title = "Please Select File")`**<br/>&bull; open the file selection dialog to let the user choose a file.<br/>&bull; default `extensions` is `Array("*.xlsx", "*.xls", "*.xlsm", "*.xlsb")`                                                                     |

#### Data Structure

| Return   | Method                                                                                                    |
| -------- | :-------------------------------------------------------------------------------------------------------- |
| `Object` | **`newArrayList()`**<br/>&bull; return a new instance of `System.Collections.ArrayList`                   |
| `Object` | **`newDictionary()`**<br/>&bull; return a new instance of `Scripting.Dictionary`                          |
| `Object` | **`newHashtable()`**<br/>&bull; return a new instance of `System.Collections.Hashtable`                   |
| `Object` | **`hashtableKeys(Object hashtable)`**<br/>&bull; return an `ArrayList<Key>` in hashtable                  |
| `Object` | **`hashtableValues(Object hashtable)`**<br/>&bull; return an `ArrayList<Value>` in hashtable              |
| `Object` | **`hashtableEntries(Object hashtable)`**<br/>&bull; return an `ArrayList<Array(Key, Value)>` in hashtable |
| `Object` | **`newFs()`**<br/>&bull; return a new instance of `Scripting.FileSystemObject`                            |
