VERSION 1.0.2 CLASS
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
    prLastRow = setLastRow(prWs.UsedRange.Rows(prWs.UsedRange.Rows.Count).Row)
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
            If Len(Trim(prWs.Cells(i, j).Value)) > 0 Then
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
    ColNumToLetter = Split(Cells(1, lColNum).Address, "$")(1)
End Function

Public Function indexOf(ByVal searchRow As Long, ByVal searchCol As Long, ByVal target As String) As Variant
    ' @description:
    '       if searchRow = -1, searchCol > 0, find the first cell in `searchCol` that contains `target` in content
    '       if searchRow > 0, searchCol = -1, find the first cell in `searchRow` that contains `target` in content
    '       if searchRow > 0 And searchCol > 0, start from (searchRow, searchCol)find the first cell in whole worksheet that contains `target` in content
    ' @ return Array(Long, Long) to represent (row, column)
    Dim i As Long
    Dim j As Long
    Dim cell As Variant
    cell = Array(-1, -1)
    If searchRow < 0 And searchCol < 0 Or searchRow > prLastRow Or searchCol > prLastCol Then
        GoTo found
    End If
    If searchRow > 0 And searchCol < 0 Then
        For i = 1 To prLastCol
            If InStr(prWs.Cells(searchRow, i).Value, target) > 0 Then
                cell = Array(searchRow, i)
                GoTo found
            End If
        Next i
    ElseIf searchRow < 0 And searchCol > 0 Then
        For i = 1 To prLastRow
            If InStr(prWs.Cells(i, searchCol).Value, target) > 0 Then
                cell = Array(i, searchCol)
                GoTo found
            End If
        Next i
    Else
        For i = searchRow To prLastRow
            For j = searchCol To prLastCol
                If InStr(prWs.Cells(i, j).Value, target) > 0 Then
                    cell = Array(i, j)
                    GoTo found
                End If
            Next j
        Next i
    End If
found:
    indexOf = cell
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
