# vista

###### VERSION 1.0.2 CLASS

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

+ `getLastRow()`: get the latest last row number of the worksheet
+ `getLastCol()`: get the latest last column number of the worksheet
+ `removeRow(i)`: remove the i th row
+ `removeCol(i)`: remove the i th column
+ `getFirstNonEmptyCell()`: return `Array(row as Long, col as Long)` to represent the position of the first non empty cell in worksheet
+ `colLetterToNum(s)`: column letter to number
+ `colNumToLetter(n)`: column number to letter
+ `indexOf(searchRow as Long, searchCol as Long, target as String)`: 
    - if searchRow = -1, searchCol > 0, find the first cell in `searchCol` that contains `target` in content
    - if searchRow > 0, searchCol = -1, find the first cell in `searchRow` that contains `target` in content
    - if searchRow > 0 And searchCol > 0, start from (searchRow, searchCol)find the first cell in whole worksheet that contains `target` in content
