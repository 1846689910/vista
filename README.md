# vista

###### VERSION 1.0.3 CLASS

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
  Dim vWs As New Vista ' create a new vista wrapper
  vWs.init Sheet1 ' initialize the vista wrapper instance with worksheet
End Sub
```

## API

### Function

#### `getLastRow()`

- get the latest last row number of the worksheet

#### `getLastCol()`

- get the latest last column number of the worksheet

#### `sync()`

- manually trigger properties update for the worksheet, include `lastRow` and `lastCol`
- should be used especially after the worksheet was directly manipulated by user insteadof operated through `vista`

#### `removeRow(i as Long)`

- remove the i th row

#### `removeCol(i as Long)`

- remove the i th column

#### `getFirstNonEmptyCell() as Variant`

- return `Array(row as Long, col as Long)` to represent the position of the first non empty cell in worksheet
- if found, return `Array(row, col)`, otherwise return `Array(-1, -1)`

#### `colLetterToNum(s) as Integer`

- column letter to number

#### `colNumToLetter(n) as String`

- column number to letter

#### `rowIndexOf(searchRow As Long, startCol As Long, target As String, Optional exactMatch As Boolean = False) as Variant`
- return `Array(row as Long, col as Long)` to represent the position of the first cell found in `searchRow` and searched from `startCol` that contains content `target`
- if `exactMatch` then cell content should be exactly equal to `target`, otherwise cell content should contain `target`
- if not found, return `Array(-1, -1)`

#### `rowIndicesOf(searchRow As Long, startCol As Long, target As String, Optional exactMatch As Boolean = False) as Object`
- return `ArrayList<Array(row as Long, col as Long)>` to represent a list of position arrays of the cell found in `searchRow` and searched from `startCol` that contains content `target`
- if `exactMatch` then cell content should be exactly equal to `target`, otherwise cell content should contain `target`
- if not found, return empty ArrayList

#### `colIndexOf(searchCol As Long, startRow As Long, target As String, Optional exactMatch As Boolean = False) As Variant`
- return `Array(row as Long, col as Long)` to represent the position of the first cell found in `searchCol` and searched from `startRow` that contains content `target`
- if `exactMatch` then cell content should be exactly equal to `target`, otherwise cell content should contain `target`
- if not found, return `Array(-1, -1)`

#### `colIndicesOf(searchCol As Long, startRow As Long, target As String, Optional exactMatch As Boolean = False) As Object`
- return `ArrayList<Array(row as Long, col as Long)>` to represent a list of position arrays of the cell found in `searchCol` and searched from `startRow` that contains content `target`
- if `exactMatch` then cell content should be exactly equal to `target`, otherwise cell content should contain `target`
- if not found, return empty ArrayList

#### `newArrayList() As Object`
- return a new instance of `System.Collections.ArrayList`