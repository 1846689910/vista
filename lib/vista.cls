VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "vista"
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

' setters and getters
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
