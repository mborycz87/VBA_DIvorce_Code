Option Explicit

Public Function EnsureRowExists(ByVal tbl As ListObject) As Long
    If tbl Is Nothing Then
        EnsureRowExists = 0
        Exit Function
    End If

    If tbl.DataBodyRange Is Nothing Then
        tbl.ListRows.Add
        EnsureRowExists = 1
    Else
        tbl.ListRows.Add
        EnsureRowExists = tbl.DataBodyRange.Rows.Count
    End If
End Function

Public Function SetCellValue(ByVal targetCell As Range, ByVal value As Variant) As Boolean
    If targetCell Is Nothing Then
        SetCellValue = False
        Exit Function
    End If

    targetCell.Value = value
    SetCellValue = True
End Function

Public Function SetRowField(ByVal tbl As ListObject, ByVal rowIndex As Long, ByVal columnIndex As Long, ByVal value As Variant) As Boolean
    If tbl Is Nothing Then
        SetRowField = False
        Exit Function
    End If

    If tbl.DataBodyRange Is Nothing Then
        SetRowField = False
        Exit Function
    End If

    If rowIndex < 1 Or rowIndex > tbl.DataBodyRange.Rows.Count Then
        SetRowField = False
        Exit Function
    End If

    SetRowField = SetCellValue(tbl.DataBodyRange.Cells(rowIndex, columnIndex), value)
End Function

Public Function UpdateRecordRow(ByVal tbl As ListObject, ByVal rowIndex As Long, ByVal record As JobRecord) As Boolean
    If tbl Is Nothing Then
        UpdateRecordRow = False
        Exit Function
    End If

    If tbl.DataBodyRange Is Nothing Then
        UpdateRecordRow = False
        Exit Function
    End If

    If rowIndex < 1 Or rowIndex > tbl.DataBodyRange.Rows.Count Then
        UpdateRecordRow = False
        Exit Function
    End If

    SetRowField tbl, rowIndex, RecordColumn.ColJobNumber, record.JobNumber
    SetRowField tbl, rowIndex, RecordColumn.ColStartDate, record.StartDate
    SetRowField tbl, rowIndex, RecordColumn.ColStartTime, record.StartTime
    SetRowField tbl, rowIndex, RecordColumn.ColEndDate, record.EndDate
    SetRowField tbl, rowIndex, RecordColumn.ColDescription, record.Description
    SetRowField tbl, rowIndex, RecordColumn.ColNotes, record.Notes
    SetRowField tbl, rowIndex, RecordColumn.ColKeywords, record.Keywords
    SetRowField tbl, rowIndex, RecordColumn.ColLink, record.Link

    UpdateRecordRow = True
End Function

Public Function CreateRecordFromSelection(ByVal lbTarget As MSForms.ListBox) As JobRecord
    Dim record As JobRecord
    Dim tableName As String
    Dim rowIndex As Long
    Dim tbl As ListObject

    tableName = GetTableNameFromSelection(lbTarget)
    rowIndex = GetSelectedResultsIndex(lbTarget)

    If Len(tableName) = 0 Or rowIndex < 0 Then
        CreateRecordFromSelection = record
        Exit Function
    End If

    Set tbl = FindTableByName(tableName)
    If tbl Is Nothing Then
        CreateRecordFromSelection = record
        Exit Function
    End If

    If rowIndex + 1 <= GetRowsInTable(tbl) Then
        record = ReadRow(tbl, rowIndex + 1)
    End If

    CreateRecordFromSelection = record
End Function

Public Function ReplaceSelectionWithRecord(ByVal lbTarget As MSForms.ListBox, ByVal tbl As ListObject, ByVal record As JobRecord) As Boolean
    Dim selectedIndex As Long

    If lbTarget Is Nothing Then
        ReplaceSelectionWithRecord = False
        Exit Function
    End If

    selectedIndex = GetSelectedResultsIndex(lbTarget)
    If selectedIndex < 0 Then
        ReplaceSelectionWithRecord = False
        Exit Function
    End If

    ReplaceSelectionWithRecord = UpdateRecordRow(tbl, selectedIndex + 1, record)
End Function

Public Function EnsureListBoxHasSelection(ByVal lbTarget As MSForms.ListBox) As Boolean
    If lbTarget Is Nothing Then
        EnsureListBoxHasSelection = False
        Exit Function
    End If

    EnsureListBoxHasSelection = (lbTarget.ListIndex >= 0)
End Function

Public Sub PopulateListWithTableRows(ByVal lbTarget As MSForms.ListBox, ByVal tbl As ListObject)
    Dim resultArray() As Variant
    Dim rowIndex As Long
    Dim rowCount As Long

    If lbTarget Is Nothing Then Exit Sub
    If tbl Is Nothing Then Exit Sub

    If tbl.DataBodyRange Is Nothing Then
        lbTarget.Clear
        Exit Sub
    End If

    rowCount = tbl.DataBodyRange.Rows.Count
    ReDim resultArray(0 To rowCount - 1, 0 To 9)

    For rowIndex = 1 To rowCount
        resultArray(rowIndex - 1, 0) = tbl.Name
        resultArray(rowIndex - 1, 1) = NormalizeText(tbl.DataBodyRange.Cells(rowIndex, RecordColumn.ColJobNumber).Value)
        resultArray(rowIndex - 1, 2) = NormalizeText(tbl.DataBodyRange.Cells(rowIndex, RecordColumn.ColStartDate).Value)
        resultArray(rowIndex - 1, 3) = NormalizeText(tbl.DataBodyRange.Cells(rowIndex, RecordColumn.ColStartTime).Value)
        resultArray(rowIndex - 1, 4) = NormalizeText(tbl.DataBodyRange.Cells(rowIndex, RecordColumn.ColEndDate).Value)
        resultArray(rowIndex - 1, 5) = NormalizeText(tbl.DataBodyRange.Cells(rowIndex, RecordColumn.ColDescription).Value)
        resultArray(rowIndex - 1, 6) = NormalizeText(tbl.DataBodyRange.Cells(rowIndex, RecordColumn.ColNotes).Value)
        resultArray(rowIndex - 1, 7) = NormalizeText(tbl.DataBodyRange.Cells(rowIndex, RecordColumn.ColKeywords).Value)
        resultArray(rowIndex - 1, 8) = NormalizeText(tbl.DataBodyRange.Cells(rowIndex, RecordColumn.ColLink).Value)
        resultArray(rowIndex - 1, 9) = NormalizeText(tbl.DataBodyRange.Cells(rowIndex, RecordColumn.ColTableName).Value)
    Next rowIndex

    lbTarget.Clear
    lbTarget.List = resultArray
    lbTarget.FontSize = 8
    lbTarget.ColumnWidths = "75,15,50,30,50,30,112,113,0,0"
End Sub
