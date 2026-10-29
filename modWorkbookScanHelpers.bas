Option Explicit

Public Function GetCurrentWorkbookState() As String
    Dim ws As Worksheet
    Dim tbl As ListObject
    Dim countSheets As Long
    Dim countTables As Long
    Dim summary As String

    countSheets = 0
    countTables = 0

    For Each ws In ThisWorkbook.Worksheets
        countSheets = countSheets + 1
        For Each tbl In ws.ListObjects
            countTables = countTables + 1
        Next tbl
    Next ws

    summary = "Sheets: " & countSheets & "; Tables: " & countTables
    GetCurrentWorkbookState = summary
End Function

Public Function GetAllListObjects() As Collection
    Dim ws As Worksheet
    Dim tbl As ListObject
    Dim results As New Collection

    For Each ws In ThisWorkbook.Worksheets
        For Each tbl In ws.ListObjects
            results.Add tbl
        Next tbl
    Next ws

    Set GetAllListObjects = results
End Function

Public Function FindTableRowByKeyword(ByVal tbl As ListObject, ByVal keyword As String, ByVal columnIndex As Long) As Long
    Dim rowIndex As Long
    Dim cellValue As String

    If tbl Is Nothing Then
        FindTableRowByKeyword = 0
        Exit Function
    End If

    If tbl.DataBodyRange Is Nothing Then
        FindTableRowByKeyword = 0
        Exit Function
    End If

    For rowIndex = 1 To tbl.DataBodyRange.Rows.Count
        cellValue = NormalizeText(tbl.DataBodyRange.Cells(rowIndex, columnIndex).Value)
        If InStr(1, cellValue, keyword, vbTextCompare) > 0 Then
            FindTableRowByKeyword = rowIndex
            Exit Function
        End If
    Next rowIndex

    FindTableRowByKeyword = 0
End Function

Public Function FindTableRowByDateRange(ByVal tbl As ListObject, ByVal startDate As Date, ByVal endDate As Date) As Collection
    Dim rowIndex As Long
    Dim rowValues As New Collection
    Dim currentDate As Date

    If tbl Is Nothing Then
        Set FindTableRowByDateRange = rowValues
        Exit Function
    End If

    If tbl.DataBodyRange Is Nothing Then
        Set FindTableRowByDateRange = rowValues
        Exit Function
    End If

    For rowIndex = 1 To tbl.DataBodyRange.Rows.Count
        If IsDate(tbl.DataBodyRange.Cells(rowIndex, RecordColumn.ColStartDate).Value) Then
            currentDate = CDate(tbl.DataBodyRange.Cells(rowIndex, RecordColumn.ColStartDate).Value)
            If currentDate >= startDate And currentDate <= endDate Then
                rowValues.Add rowIndex
            End If
        End If
    Next rowIndex

    Set FindTableRowByDateRange = rowValues
End Function

Public Function ContainsAnyKeyword(ByVal cellValue As Variant, ByVal keywords As Variant) As Boolean
    Dim arr As Variant
    Dim i As Long
    Dim textValue As String

    textValue = NormalizeText(cellValue)
    If Len(textValue) = 0 Then
        ContainsAnyKeyword = False
        Exit Function
    End If

    If IsArray(keywords) Then
        arr = keywords
        For i = LBound(arr) To UBound(arr)
            If InStr(1, textValue, CStr(arr(i)), vbTextCompare) > 0 Then
                ContainsAnyKeyword = True
                Exit Function
            End If
        Next i
    Else
        ContainsAnyKeyword = InStr(1, textValue, CStr(keywords), vbTextCompare) > 0
    End If
End Function

Public Function GetRowsInTable(ByVal tbl As ListObject) As Long
    If tbl Is Nothing Then
        GetRowsInTable = 0
        Exit Function
    End If

    If tbl.DataBodyRange Is Nothing Then
        GetRowsInTable = 0
    Else
        GetRowsInTable = tbl.DataBodyRange.Rows.Count
    End If
End Function
