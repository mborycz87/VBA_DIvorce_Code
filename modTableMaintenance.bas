Option Explicit

Public Function InsertBlankRow(ByVal tbl As ListObject) As Long
    If tbl Is Nothing Then
        InsertBlankRow = 0
        Exit Function
    End If

    tbl.ListRows.Add
    If tbl.DataBodyRange Is Nothing Then
        InsertBlankRow = 1
    Else
        InsertBlankRow = tbl.DataBodyRange.Rows.Count
    End If
End Function

Public Function ClearRowValues(ByVal tbl As ListObject, ByVal rowIndex As Long) As Boolean
    Dim colIndex As Long

    If tbl Is Nothing Then
        ClearRowValues = False
        Exit Function
    End If

    If tbl.DataBodyRange Is Nothing Then
        ClearRowValues = False
        Exit Function
    End If

    If rowIndex < 1 Or rowIndex > tbl.DataBodyRange.Rows.Count Then
        ClearRowValues = False
        Exit Function
    End If

    For colIndex = 1 To tbl.ListColumns.Count
        tbl.DataBodyRange.Cells(rowIndex, colIndex).Value = ""
    Next colIndex

    ClearRowValues = True
End Function

Public Function DeleteRow(ByVal tbl As ListObject, ByVal rowIndex As Long) As Boolean
    If tbl Is Nothing Then
        DeleteRow = False
        Exit Function
    End If

    If tbl.DataBodyRange Is Nothing Then
        DeleteRow = False
        Exit Function
    End If

    If rowIndex < 1 Or rowIndex > tbl.DataBodyRange.Rows.Count Then
        DeleteRow = False
        Exit Function
    End If

    tbl.ListRows(rowIndex).Delete
    DeleteRow = True
End Function

Public Function UpdateRowCell(ByVal tbl As ListObject, ByVal rowIndex As Long, ByVal columnIndex As Long, ByVal value As Variant) As Boolean
    If tbl Is Nothing Then
        UpdateRowCell = False
        Exit Function
    End If

    If tbl.DataBodyRange Is Nothing Then
        UpdateRowCell = False
        Exit Function
    End If

    If rowIndex < 1 Or rowIndex > tbl.DataBodyRange.Rows.Count Then
        UpdateRowCell = False
        Exit Function
    End If

    If columnIndex < 1 Or columnIndex > tbl.ListColumns.Count Then
        UpdateRowCell = False
        Exit Function
    End If

    tbl.DataBodyRange.Cells(rowIndex, columnIndex).Value = value
    UpdateRowCell = True
End Function

Public Function UpdateRowFromRecord(ByVal tbl As ListObject, ByVal rowIndex As Long, ByVal record As JobRecord) As Boolean
    If tbl Is Nothing Then
        UpdateRowFromRecord = False
        Exit Function
    End If

    If tbl.DataBodyRange Is Nothing Then
        UpdateRowFromRecord = False
        Exit Function
    End If

    If rowIndex < 1 Or rowIndex > tbl.DataBodyRange.Rows.Count Then
        UpdateRowFromRecord = False
        Exit Function
    End If

    UpdateRowCell tbl, rowIndex, RecordColumn.ColJobNumber, record.JobNumber
    UpdateRowCell tbl, rowIndex, RecordColumn.ColStartDate, record.StartDate
    UpdateRowCell tbl, rowIndex, RecordColumn.ColStartTime, record.StartTime
    UpdateRowCell tbl, rowIndex, RecordColumn.ColEndDate, record.EndDate
    UpdateRowCell tbl, rowIndex, RecordColumn.ColDescription, record.Description
    UpdateRowCell tbl, rowIndex, RecordColumn.ColNotes, record.Notes
    UpdateRowCell tbl, rowIndex, RecordColumn.ColKeywords, record.Keywords
    UpdateRowCell tbl, rowIndex, RecordColumn.ColLink, record.Link

    UpdateRowFromRecord = True
End Function

Public Function GetCellText(ByVal targetCell As Range) As String
    If targetCell Is Nothing Then
        GetCellText = ""
        Exit Function
    End If

    GetCellText = NormalizeText(targetCell.Value)
End Function

Public Function SetCellText(ByVal targetCell As Range, ByVal value As String) As Boolean
    If targetCell Is Nothing Then
        SetCellText = False
        Exit Function
    End If

    targetCell.Value = value
    SetCellText = True
End Function
