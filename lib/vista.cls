VERSION 1.0.3 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Vista"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private prLastRow As Long
Private prLastCol As Long
Private prWs As Worksheet

' initialize
Public Sub init(ByRef sheet As Worksheet)
    Set prWs = sheet
    sync
End Sub

' if the user directly manipulate the worksheet, then this method could be used to sync the latest properties of worksheet
Public Sub sync()
    prLastRow = setLastRow(prWs.UsedRange.Rows(prWs.UsedRange.Rows.Count).row)
    prLastCol = setLastCol(prWs.UsedRange.Columns(prWs.UsedRange.Columns.Count).Column)
End Sub
Public Sub removeRow(r As Long)
    If r < 1 Or r > prLastRow Then Exit Sub
    prWs.Rows(r).EntireRow.Delete
    prLastRow = prLastRow - 1
End Sub

Public Sub removeCol(c As Long)
    If c < 1 Or c > prLastCol Then Exit Sub
    prWs.Columns(c).EntireColumn.Delete
    prLastCol = prLastCol - 1
End Sub

Public Function getFirstNonEmptyCell() As Variant
    ' @return Array(Long, Long) to represent (row, column)
    Dim i As Long
    Dim j As Long
    Dim cell As Variant
    cell = Array(-1, -1)
    For i = 1 To prLastRow
        For j = 1 To prLastCol
            If Len(Trim(prWs.cells(i, j).Value)) > 0 Then
                cell = Array(i, j)
                GoTo found
            End If
        Next j
    Next i
found:
    getFirstNonEmptyCell = cell
End Function

Public Function ColLetterToNum(ByVal sColLetter As String) As Integer
    ColLetterToNum = ActiveWorkbook.Worksheets(1).Columns(sColLetter).Column
End Function

Public Function ColNumToLetter(lColNum As Integer) As String
    ColNumToLetter = Split(cells(1, lColNum).Address, "$")(1)
End Function
Public Function rowIndexOf(ByVal searchRow As Long, ByVal startCol As Long, ByVal target As String, Optional ByVal exactMatch As Boolean = False) As Variant
    Dim cell As Variant
    Dim i As Long
    cell = Array(-1, -1)
    If searchRow < 1 Or searchRow > prLastRow Or startCol < 1 Or startCol > prLastCol Then GoTo found
    For i = startCol To prLastCol
        If matchContent(prWs.cells(searchRow, i).Value, target, exactMatch) Then
            cell = Array(searchRow, i)
            GoTo found
        End If
    Next i
found:
    rowIndexOf = cell
End Function
Public Function rowIndicesOf(ByVal searchRow As Long, ByVal startCol As Long, ByVal target As String, Optional ByVal exactMatch As Boolean = False) As Object
    Dim cells As Object
    Dim i As Long
    Set cells = newArrayList()
    If searchRow < 1 Or searchRow > prLastRow Or startCol < 1 Or startCol > prLastCol Then GoTo found
    For i = startCol To prLastCol
        If matchContent(prWs.cells(searchRow, i).Value, target, exactMatch) Then cells.Add Array(seachrow, i)
    Next i
found:
    Set rowIndicesOf = cells
End Function
Public Function colIndexOf(ByVal searchCol As Long, ByVal startRow As Long, ByVal target As String, Optional ByVal exactMatch As Boolean = False) As Variant
    Dim cell As Variant
    Dim i As Long
    cell = Array(-1, -1)
    If searchCol < 1 Or searchCol > prLastCol Or startRow < 1 Or startRow > prLastRow Then GoTo found
    For i = startRow To prLastRow
        If matchContent(prWs.cells(i, searchCol).Value, target, exactMatch) Then
            cell = Array(i, searchCol)
            GoTo found
        End If
    Next i
found:
    colIndexOf = cell
End Function
Public Function colIndicesOf(ByVal searchCol As Long, ByVal startRow As Long, ByVal target As String, Optional ByVal exactMatch As Boolean = False) As Object
    Dim cells As Object
    Dim i As Long
    Set cells = newArrayList()
    If searchCol < 1 Or searchCol > prLastCol Or startRow < 1 Or startRow > prLastRow Then GoTo found
    For i = startRow To prLastRow
        If matchContent(prWs.cells(i, searchCol).Value, target, exactMatch) Then cells.Add Array(i, searchCol)
    Next i
found:
    Set colIndicesOf = cells
End Function

Private Function matchContent(ByVal cellContent As String, ByVal target As String, ByVal exact As Boolean) As Boolean
    If exact Then
        matchContent = cellContent = target
    Else
        matchContent = InStr(cellContent, target) > 0
    End If
End Function

Public Function newArrayList() As Object
    Set newArrayList = CreateObject("System.Collections.ArrayList")
End Function

' ///////////////////////////////////////////////////////////
' //////////////////  setters and getters //////////////////
' //////////////////////////////////////////////////////////
Public Function getLastRow() As Long
    getLastRow = prLastRow
End Function

Private Function setLastRow(r As Long) As Long
    prLastRow = r
    setLastRow = prLastRow
End Function

Public Function getLastCol() As Long
    getLastCol = prLastCol
End Function

Private Function setLastCol(r As Long) As Long
    prLastCol = r
    setLastCol = prLastCol
End Function

Public Function getWs() As Worksheet
    Set getWs = prWs
End Function
