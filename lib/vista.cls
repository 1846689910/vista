VERSION 1.0 CLASS
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
Private initialized As Boolean

' initialize
Public Sub init(ByRef sheet As Worksheet)
    If initialized Then Exit Sub
    Set prWs = sheet
    sync
    initialized = True
End Sub

' if the user directly manipulate the worksheet, then this method could be used to sync the latest properties of worksheet
Public Sub sync()
    prLastRow = setLastRow(prWs.UsedRange.Rows(prWs.UsedRange.Rows.Count).row)
    prLastCol = setLastCol(prWs.UsedRange.Columns(prWs.UsedRange.Columns.Count).Column)
End Sub
Public Sub removeRow(ByVal r As Long)
    If r < 1 Or r > prLastRow Then Exit Sub
    prWs.Rows(r).EntireRow.Delete
    prLastRow = prLastRow - 1
End Sub

Public Sub addRow(ByVal r As Long)
    If r < 1 Then Exit Sub
    prWs.Rows(r).Insert
    prLastRow = setLastRow(prWs.UsedRange.Rows(prWs.UsedRange.Rows.Count).row)
End Sub

Public Sub removeCol(ByVal c As Long)
    If c < 1 Or c > prLastCol Then Exit Sub
    prWs.Columns(c).EntireColumn.Delete
    prLastCol = prLastCol - 1
End Sub

Public Sub addCol(ByVal c As Long)
    If c < 1 Then Exit Sub
    prWs.Columns(c).Insert
    prLastCol = setLastCol(prWs.UsedRange.Columns(prWs.UsedRange.Columns.Count).Column)
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

Function colNumToLetter(ByVal n As Long) As String
    Dim s As String
    s = ""
    While n > 0
        n = n - 1
        s = s & Chr(Asc("A") + n Mod 26)
        n = Application.WorksheetFunction.RoundDown(n / 26, 0)
    Wend
    colNumToLetter = StrReverse(s)
End Function

Function colLetterToNum(ByVal s As String) As Long
    Dim n As Long
    Dim i As Integer
    n = 0
    For i = 1 To Len(s)
        n = 26 * n + (Asc(Mid(s, i, 1)) - Asc(Mid("A", 1, 1)) + 1)
    Next i
    colLetterToNum = n
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

Public Function getRowColFromRangeSelector(ByVal selector As String) As Variant
    Dim rng As Range
    Set rng = prWs.Range(selector)
    getRowColFromRangeSelector = Array(rng.row, rng.Column)
End Function

Public Function getRangeSelectorFromRowCol(ByRef rowCol As Variant) As String
    getRangeSelectorFromRowCol = colNumToLetter(CLng(rowCol(1))) & rowCol(0)
End Function

Public Sub mkDir(ByVal path As String)
    If existDir(path) Then Exit Sub
    Dim arr() As String
    Dim midPath As String
    Dim i As Integer, j As Integer
    arr = Split(path, "\")
    For i = LBound(arr) To UBound(arr)
        midPath = ""
        For j = LBound(arr) To i
            midPath = midPath & arr(j) & "\"
        Next j
        If Dir(midPath, vbDirectory) = "" Then mkDir midPath
    Next i
End Sub

Public Sub mkWorkbook(ByVal path As String, Optional ByVal TEMPLATE_PATH As String)
    If existFile(path) Then Exit Sub
    Dim arr() As String
    Dim dirpath As String
    Dim filename As String
    Dim wb As Workbook
    filename = basename(path)
    dirpath = dirname(path)
    If Not existDir(dirpath) Then mkDir (dirpath)
    If IsMissing(TEMPLATE_PATH) Then
        Set wb = Workbooks.Add(xlWBATWorksheet)
    Else
        Set wb = Workbooks.Add(TEMPLATE_PATH)
    End If
    With wb
        .Title = filename
        .Subject = filename
        .SaveAs filename:=path
    End With
    wb.Close savechanges:=True
End Sub

Public Function openWorkbook(ByVal path As String) As Workbook
    Set openWorkbook = Workbooks.Open(filename:=path)
End Function

Public Sub saveasWorkbook(ByRef wb As Workbook, ByVal path As String)
    wb.SaveAs path
End Sub

Public Function existFile(ByVal path As String) As Boolean
    existFile = Len(Dir(path)) > 0
End Function

Public Function existDir(ByVal path As String) As Boolean
    existDir = Len(Dir(path, vbDirectory)) > 0
End Function

Public Function dirname(ByVal path As String) As String
    Dim arr() As String
    arr = Split(path, "\")
    If UBound(arr) < 1 Then Err.Raise 619, "dirname", "Vista.dirname: no parent path"
    ReDim Preserve arr(UBound(arr) - 1)
    dirname = Join(arr, "\")
End Function

Public Function basename(ByVal path As String) As String
    Dim arr() As String
    arr = Split(path, "\")
    basename = arr(UBound(arr))
End Function

Public Function hasWorksheet(ByRef wb As Workbook, ByVal name As String) As Boolean
    Dim found As Boolean
    found = False
    Dim ws As Worksheet
    Dim tWs As Variant
    For Each tWs In wb.Worksheets
        If tWs.name = name Then found = True
    Next tWs
    hasWorksheet = found
End Function

Public Sub addWorksheet(ByRef wb As Workbook, ByVal name As String)
    If hasWorksheet(wb, name) Then Exit Sub
    wb.Sheets.Add(After:=wb.Sheets(wb.Sheets.Count)).name = name
End Sub

Public Sub removeWorksheet(ByRef wb As Workbook, ByVal name As String)
    If Not hasWorksheet(wb, name) Then Exit Sub
    Application.DisplayAlerts = False
    wb.Worksheets(name).Delete
    Application.DisplayAlerts = True
End Sub

Public Function getAllFilenames(ByVal path As String) As Object
    Set getAllFilenames = getAllFilePaths_R(path, needPath:=False, level:=1)
End Function

Public Function getAllFilePaths_R(ByVal path As String, Optional ByVal needPath As Boolean = True, Optional ByVal level As Variant) As Object
    Dim list As Object
    Dim fs As Object
    Set list = newArrayList()
    If IsMissing(level) Then
        Set fs = newFs()
        getAllFilePathsHelper fs.getFolder(path), list, needPath
    ElseIf level < 1 Then GoTo finish
    Else
        Set fs = newFs()
        getAllFilePathsHelper fs.getFolder(path), list, needPath, level:=level
    End If
finish:
    Set getAllFilePaths_R = list
End Function

Private Sub getAllFilePathsHelper(ByRef folder As Variant, ByRef list As Object, ByVal needPath As Boolean, Optional ByRef level As Variant)
    Dim file As Variant
    If Not IsMissing(level) Then
        If level <= 1 Then GoTo enough
        level = level - 1
    End If
    Dim subFolder As Variant
    For Each subFolder In folder.subFolders
        getAllFilePathsHelper subFolder, list, needPath, level
    Next subFolder
enough:
    For Each file In folder.Files
        If needPath Then
            list.Add (file.path)
        Else
            list.Add (file.name)
        End If
    Next file
End Sub

Public Function getAllSubDirs(ByVal path As String) As Object
    Set getAllSubDirs = getAllSubDirs_R(path, needPath:=False, level:=1)
End Function

Public Function getAllSubDirs_R(ByVal path As String, Optional ByVal needPath As Boolean = True, Optional ByVal level As Variant) As Object
    Dim list As Object
    Dim fs As Object
    Set list = newArrayList()
    If IsMissing(level) Then
        Set fs = newFs()
        getAllSubDirsHelper fs.getFolder(path), list, needPath
    ElseIf level < 1 Then GoTo finish
    Else
        Set fs = newFs()
        getAllSubDirsHelper fs.getFolder(path), list, needPath, level:=level
    End If
finish:
    Set getAllSubDirs_R = list
End Function

Private Sub getAllSubDirsHelper(ByRef folder As Variant, ByRef list As Object, ByVal needPath As Boolean, Optional ByRef level As Variant)
    If Not IsMissing(level) Then
        If level < 1 Then GoTo enough
        level = level - 1
    End If
    Dim subFolder As Variant
    For Each subFolder In folder.subFolders
        If needPath Then
            list.Add subFolder.path
        Else
            list.Add subFolder.name
        End If
        getAllSubDirsHelper subFolder, list, needPath, level
    Next subFolder
enough:
End Sub

' ///////////////////////////////////////////////////////////
' ///////////////////// Data Structure /////////////////////
' //////////////////////////////////////////////////////////
Public Function newArrayList() As Object
    Set newArrayList = CreateObject("System.Collections.ArrayList")
End Function

Public Function newDictionary() As Object
    Set newDictionary = CreateObject("Scripting.Dictionary")
End Function

Public Function newHashtable() As Object
    Set newHashtable = CreateObject("System.Collections.Hashtable")
End Function

Public Function newFs() As Object
    Set newFs = CreateObject("Scripting.FileSystemObject")
End Function

Public Function hashtableKeys(ByRef hashtable As Object) As Object
    Dim keys As Object
    Set keys = newArrayList()
    keys.addRange (hashtable.keys)
    Set hashtableKeys = keys
End Function

Public Function hashtableValues(ByRef hashtable As Object) As Object
    Dim values As Object
    Set values = newArrayList()
    values.addRange (hashtable.values)
    Set hashtableValues = values
End Function

Public Function hashtableEntries(ByRef hashtable As Object) As Object
    Dim entries As Object
    Dim i As Long
    Set entries = hashtableKeys(hashtable)
    For i = 0 To entries.Count - 1
        entries(i) = Array(entries(i), hashtable(entries(i)))
    Next i
    Set hashtableEntries = entries
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
