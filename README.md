# vista

###### VERSION 1.0.1 CLASS

## introduction

vista is a helper class module for excel VBA use. It could assit you in repeating and complex worksheet manipulation.

## How does it work?

vista will perform as a worksheet wrapper for a specific worksheet and provide many useful functions and subs to interact with worksheet.

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
