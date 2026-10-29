Option Explicit

Public Function ReadRow(ByVal tbl As ListObject, ByVal rowIndex As Long) As JobRecord
    Dim rowRange As Range
    Dim startDateValue As Variant
    Dim endDateValue As Variant

    If tbl Is Nothing Then Exit Function
    If rowIndex < 1 Then Exit Function

    If tbl.DataBodyRange Is Nothing Then Exit Function
    If rowIndex > tbl.DataBodyRange.Rows.Count Then Exit Function

    Set rowRange = tbl.DataBodyRange.Rows(rowIndex)
    startDateValue = rowRange.Cells(1, RecordColumn.ColStartDate).Value
    endDateValue = rowRange.Cells(1, RecordColumn.ColEndDate).Value

    With ReadRow
        .TableName = NormalizeText(tbl.Name)
        .JobNumber = NormalizeText(rowRange.Cells(1, RecordColumn.ColJobNumber).Value)
        .StartDate = startDateValue
        .StartTime = NormalizeText(rowRange.Cells(1, RecordColumn.ColStartTime).Value)
        .EndDate = endDateValue
        .Description = NormalizeText(rowRange.Cells(1, RecordColumn.ColDescription).Value)
        .Notes = NormalizeText(rowRange.Cells(1, RecordColumn.ColNotes).Value)
        .Keywords = NormalizeText(rowRange.Cells(1, RecordColumn.ColKeywords).Value)
        .Link = NormalizeText(rowRange.Cells(1, RecordColumn.ColLink).Value)
    End With
End Function

Public Sub WriteRow(ByVal tbl As ListObject, ByVal rowIndex As Long, ByVal record As JobRecord)
    Dim rowRange As Range

    If tbl Is Nothing Then Exit Sub
    If rowIndex < 1 Then Exit Sub

    If tbl.DataBodyRange Is Nothing Then
        tbl.ListRows.Add
        Set rowRange = tbl.DataBodyRange.Rows(1)
    Else
        If rowIndex > tbl.DataBodyRange.Rows.Count Then
            tbl.ListRows.Add
        End If
        Set rowRange = tbl.DataBodyRange.Rows(rowIndex)
    End If

    rowRange.Cells(1, RecordColumn.ColJobNumber).Value = record.JobNumber
    rowRange.Cells(1, RecordColumn.ColStartDate).Value = record.StartDate
    rowRange.Cells(1, RecordColumn.ColStartTime).Value = record.StartTime
    rowRange.Cells(1, RecordColumn.ColEndDate).Value = record.EndDate
    rowRange.Cells(1, RecordColumn.ColDescription).Value = record.Description
    rowRange.Cells(1, RecordColumn.ColNotes).Value = record.Notes
    rowRange.Cells(1, RecordColumn.ColKeywords).Value = record.Keywords
    rowRange.Cells(1, RecordColumn.ColLink).Value = record.Link
End Sub

Public Function GetWorksheetTables(ByVal ws As Worksheet) As Collection
    Dim tbl As ListObject
    Dim results As New Collection

    If ws Is Nothing Then Set GetWorksheetTables = results: Exit Function

    For Each tbl In ws.ListObjects
        results.Add tbl
    Next tbl

    Set GetWorksheetTables = results
End Function

Public Function ValueIsBlank(ByVal value As Variant) As Boolean
    If IsNull(value) Or IsEmpty(value) Then
        ValueIsBlank = True
        Exit Function
    End If

    If Len(Trim(CStr(value))) = 0 Then
        ValueIsBlank = True
    Else
        ValueIsBlank = False
    End If
End Function
