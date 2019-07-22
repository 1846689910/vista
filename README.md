# vista

###### VERSION 1.0.4 CLASS

## introduction

vista is a helper class module for excel VBA use. It could assit you in repeating and complex worksheet manipulation.

## How does it work?

vista will perform as a worksheet wrapper for a specific worksheet and provide many useful functions and subs to interact with worksheet.

1. **please download `Vista.cls` from [Vista](https://gist.github.com/1846689910/f1767e08f081bb11a9fc2a8d35018166)**
2. **`import` the module in your VBA project**

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

| Return    | Method                                                                                                                                                                                                                   |
| --------- | :----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------- |
| `Void`    | `init(Worksheet ws)` <br/> initialize the created `Vista` instance and hook the instance with the specified Worksheet `ws`                                                                                               |
| `Void`    | `sync()`<br/>manually trigger properties update for the worksheet, include `lastRow` and `lastCol`<br/>should be used especially after the worksheet was directly manipulated by user insteadof operated through `vista` |
| `Varaint` | `getFirstNonEmptyCell()`<br/>return `Array(row as Long, col as Long)` to represent the position of the first non empty cell in worksheet<br/>if found, return `Array(row, col)`, otherwise return `Array(-1, -1)`        |

#### Row

| Return    | Method                                                                                                                                                                                                                                                                                                                                                                                                                                                           |
| --------- | :--------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------- |
| `Long`    | `getLastRow()`<br/>get the latest last row number of the worksheet                                                                                                                                                                                                                                                                                                                                                                                               |
| `Void`    | `removeRow(Long i)`<br/>remove the i th row                                                                                                                                                                                                                                                                                                                                                                                                                      |
| `Varaint` | `rowIndexOf(Long searchRow, Long startCol, String target, Optional Boolean exactMatch = False)`<br/>return `Array(row as Long, col as Long)` to represent the position of the first cell found in `searchRow` and searched from `startCol` that contains content `target`<br/>if `exactMatch` then cell content should be exactly equal to `target`, otherwise cell content should contain `target`<br/>if not found, return `Array(-1, -1)`                     |
| `Object`  | `rowIndicesOf(Long searchRow, Long startCol, String target, Optional Boolean exactMatch = False)`<br/>return `ArrayList<Array(row as Long, col as Long)>` to represent a list of position arrays of the cell found in `searchRow` and searched from `startCol` that contains content `target`<br/>if `exactMatch` then cell content should be exactly equal to `target`, otherwise cell content should contain `target`<br/>if not found, return empty ArrayList |

#### Column

| Return    | Method                                                                                                                                                                                                                                                                                                                                                                                                                                                         |
| --------- | :------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------- |
| `Long`    | `getLastCol()`<br/>get the latest last column number of the worksheet                                                                                                                                                                                                                                                                                                                                                                                          |
| `Void`    | `removeCol(Long i)`<br/>remove the i th column                                                                                                                                                                                                                                                                                                                                                                                                                 |
| `Varaint` | `colIndexOf(Long searchCol, Long startRow, String target, Optional Boolean exactMatch = False)`<br/>return `Array(row as Long, col as Long)` to represent the position of the first cell found in `searchCol` and searched from `startRow` that contains content `target`<br/>if `exactMatch` then cell content should be exactly equal to `target`, otherwise cell content should contain `target`<br/>if not found, return `Array(-1, -1)`                   |
| `Object`  | `colIndexOf(Long searchCol, Long startRow, String target, Optional Boolean exactMatch = False)`<br/>return `ArrayList<Array(row as Long, col as Long)>` to represent a list of position arrays of the cell found in `searchCol` and searched from `startRow` that contains content `target`<br/>if `exactMatch` then cell content should be exactly equal to `target`, otherwise cell content should contain `target`<br/>if not found, return empty ArrayList |

#### Worksheet

#### Workbook

#### Utils

| Return    | Method                                                   |
| --------- | :------------------------------------------------------- |
| `Integer` | `colLetterToNum(String s)` <br/>column letter to number  |
| `String`  | `colNumToLetter(Integer n)` <br/>column number to letter |

#### Data Structure

| Return   | Method                                                                                         |
| -------- | :--------------------------------------------------------------------------------------------- |
| `Object` | `newArrayList()`<br/>return a new instance of `System.Collections.ArrayList`                   |
| `Object` | `newDictionary()`<br/>return a new instance of `Scripting.Dictionary`                          |
| `Object` | `newHashtable()`<br/>return a new instance of `System.Collections.Hashtable`                   |
| `Object` | `hashtableKeys(Object hashtable)`<br/>return an `ArrayList<Key>` in hashtable                  |
| `Object` | `hashtableValues(Object hashtable)`<br/>return an `ArrayList<Value>` in hashtable              |
| `Object` | `hashtableEntries(Object hashtable)`<br/>return an `ArrayList<Array(Key, Value)>` in hashtable |
