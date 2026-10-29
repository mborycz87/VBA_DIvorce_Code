Option Explicit

Public Function GetWorksheet(ByVal sheetName As String) As Worksheet
    Dim ws As Worksheet

    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets(sheetName)
    On Error GoTo 0

    If ws Is Nothing Then
        Set ws = ThisWorkbook.Worksheets.Add(After:=ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.Count))
        ws.Name = sheetName
    End If

    Set GetWorksheet = ws
End Function

Public Function GetListObject(ByVal ws As Worksheet, ByVal tableName As String) As ListObject
    Dim tbl As ListObject

    If ws Is Nothing Then Exit Function

    On Error Resume Next
    Set tbl = ws.ListObjects(tableName)
    On Error GoTo 0

    Set GetListObject = tbl
End Function

Public Function EnsureTable(ByVal ws As Worksheet, ByVal tableName As String) As ListObject
    Dim tbl As ListObject

    If ws Is Nothing Then Exit Function

    Set tbl = GetListObject(ws, tableName)
    If tbl Is Nothing Then
        Set tbl = ws.ListObjects.Add(SourceType:=xlSrcRange, Source:=ws.Range("A1").CurrentRegion, XlListObjectHasHeaders:=xlYes)
        tbl.Name = tableName
    End If

    Set EnsureTable = tbl
End Function

Public Function GetRowCount(ByVal tbl As ListObject) As Long
    If tbl Is Nothing Then
        GetRowCount = 0
        Exit Function
    End If

    If tbl.DataBodyRange Is Nothing Then
        GetRowCount = 0
    Else
        GetRowCount = tbl.DataBodyRange.Rows.Count
    End If
End Function

Public Function NormalizeText(ByVal value As Variant) As String
    If IsNull(value) Or IsEmpty(value) Then
        NormalizeText = ""
    Else
        NormalizeText = CStr(value)
    End If
End Function
