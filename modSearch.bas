Option Explicit

Public Function SearchByKeyword(ByVal keyword As String) As Variant
    Dim ws As Worksheet
    Dim tbl As ListObject
    Dim rowIndex As Long
    Dim resultCount As Long
    Dim resultArray() As Variant
    Dim rowData As JobRecord

    If Len(Trim(keyword)) = 0 Then
        SearchByKeyword = Empty
        Exit Function
    End If

    resultCount = 0
    ReDim resultArray(0 To 0, 0 To 9)

    For Each ws In ThisWorkbook.Worksheets
        For Each tbl In ws.ListObjects
            If Not tbl.DataBodyRange Is Nothing Then
                For rowIndex = 1 To tbl.DataBodyRange.Rows.Count
                    rowData = ReadRow(tbl, rowIndex)
                    If InStr(1, rowData.Keywords, keyword, vbTextCompare) > 0 Then
                        If resultCount > UBound(resultArray, 1) Then
                            ReDim Preserve resultArray(0 To resultCount, 0 To 9)
                        End If

                        resultArray(resultCount, 0) = tbl.Name
                        resultArray(resultCount, 1) = rowData.JobNumber
                        resultArray(resultCount, 2) = rowData.StartDate
                        resultArray(resultCount, 3) = rowData.StartTime
                        resultArray(resultCount, 4) = rowData.EndDate
                        resultArray(resultCount, 5) = rowData.Description
                        resultArray(resultCount, 6) = rowData.Notes
                        resultArray(resultCount, 7) = rowData.Keywords
                        resultArray(resultCount, 8) = rowData.Link
                        resultArray(resultCount, 9) = rowData.TableName
                        resultCount = resultCount + 1
                    End If
                Next rowIndex
            End If
        Next tbl
    Next ws

    If resultCount = 0 Then
        SearchByKeyword = Empty
        Exit Function
    End If

    ReDim Preserve resultArray(0 To resultCount - 1, 0 To 9)
    SearchByKeyword = resultArray
End Function

Public Function SearchByDateRange(ByVal startDate As Date, ByVal endDate As Date) As Variant
    Dim ws As Worksheet
    Dim tbl As ListObject
    Dim rowIndex As Long
    Dim resultCount As Long
    Dim resultArray() As Variant
    Dim rowData As JobRecord

    If startDate > endDate Then
        SearchByDateRange = Empty
        Exit Function
    End If

    resultCount = 0
    ReDim resultArray(0 To 0, 0 To 9)

    For Each ws In ThisWorkbook.Worksheets
        For Each tbl In ws.ListObjects
            If Not tbl.DataBodyRange Is Nothing Then
                For rowIndex = 1 To tbl.DataBodyRange.Rows.Count
                    rowData = ReadRow(tbl, rowIndex)

                    If Not IsDate(rowData.StartDate) Then GoTo NextRow
                    If Not IsDate(rowData.EndDate) Then GoTo NextRow

                    If CDate(rowData.StartDate) >= startDate And CDate(rowData.EndDate) <= endDate Then
                        If resultCount > UBound(resultArray, 1) Then
                            ReDim Preserve resultArray(0 To resultCount, 0 To 9)
                        End If

                        resultArray(resultCount, 0) = tbl.Name
                        resultArray(resultCount, 1) = rowData.JobNumber
                        resultArray(resultCount, 2) = rowData.StartDate
                        resultArray(resultCount, 3) = rowData.StartTime
                        resultArray(resultCount, 4) = rowData.EndDate
                        resultArray(resultCount, 5) = rowData.Description
                        resultArray(resultCount, 6) = rowData.Notes
                        resultArray(resultCount, 7) = rowData.Keywords
                        resultArray(resultCount, 8) = rowData.Link
                        resultArray(resultCount, 9) = rowData.TableName
                        resultCount = resultCount + 1
                    End If
                Next rowIndex
            End If
NextRow:
        Next tbl
    Next ws

    If resultCount = 0 Then
        SearchByDateRange = Empty
        Exit Function
    End If

    ReDim Preserve resultArray(0 To resultCount - 1, 0 To 9)
    SearchByDateRange = resultArray
End Function

Public Function SearchByTable(ByVal tableName As String) As Variant
    Dim ws As Worksheet
    Dim tbl As ListObject
    Dim rowIndex As Long
    Dim resultCount As Long
    Dim resultArray() As Variant
    Dim rowData As JobRecord

    If Len(Trim(tableName)) = 0 Then
        SearchByTable = Empty
        Exit Function
    End If

    resultCount = 0
    ReDim resultArray(0 To 0, 0 To 9)

    For Each ws In ThisWorkbook.Worksheets
        For Each tbl In ws.ListObjects
            If StrComp(tbl.Name, tableName, vbTextCompare) = 0 Then
                If Not tbl.DataBodyRange Is Nothing Then
                    For rowIndex = 1 To tbl.DataBodyRange.Rows.Count
                        rowData = ReadRow(tbl, rowIndex)

                        If resultCount > UBound(resultArray, 1) Then
                            ReDim Preserve resultArray(0 To resultCount, 0 To 9)
                        End If

                        resultArray(resultCount, 0) = tbl.Name
                        resultArray(resultCount, 1) = rowData.JobNumber
                        resultArray(resultCount, 2) = rowData.StartDate
                        resultArray(resultCount, 3) = rowData.StartTime
                        resultArray(resultCount, 4) = rowData.EndDate
                        resultArray(resultCount, 5) = rowData.Description
                        resultArray(resultCount, 6) = rowData.Notes
                        resultArray(resultCount, 7) = rowData.Keywords
                        resultArray(resultCount, 8) = rowData.Link
                        resultArray(resultCount, 9) = rowData.TableName
                        resultCount = resultCount + 1
                    Next rowIndex
                End If
            End If
        Next tbl
    Next ws

    If resultCount = 0 Then
        SearchByTable = Empty
        Exit Function
    End If

    ReDim Preserve resultArray(0 To resultCount - 1, 0 To 9)
    SearchByTable = resultArray
End Function

Public Function SearchAndPopulateListBox(ByVal keyword As String, ByVal lbTarget As Object) As Boolean
    Dim results As Variant

    If lbTarget Is Nothing Then
        SearchAndPopulateListBox = False
        Exit Function
    End If

    results = SearchByKeyword(keyword)
    If IsEmpty(results) Then
        SearchAndPopulateListBox = False
        Exit Function
    End If

    lbTarget.Clear
    lbTarget.List = results
    SearchAndPopulateListBox = True
End Function

Public Function PopulateListFromDateRange(ByVal startDate As Variant, ByVal endDate As Variant, ByVal lbTarget As Object) As Boolean
    Dim results As Variant
    Dim startDateValue As Date
    Dim endDateValue As Date

    If lbTarget Is Nothing Then
        PopulateListFromDateRange = False
        Exit Function
    End If

    If Not ValidateDateRange(startDate, endDate) Then
        PopulateListFromDateRange = False
        Exit Function
    End If

    startDateValue = CDate(startDate)
    endDateValue = CDate(endDate)
    results = SearchByDateRange(startDateValue, endDateValue)

    If IsEmpty(results) Then
        PopulateListFromDateRange = False
        Exit Function
    End If

    lbTarget.Clear
    lbTarget.List = results
    PopulateListFromDateRange = True
End Function

Public Function SetKeywordStateOnRow(ByVal tbl As ListObject, ByVal rowIndex As Long, ByVal keyword As String, ByVal stateSuffix As String) As Boolean
    Dim existingValue As String
    Dim cleanKeyword As String
    Dim newValue As String

    If tbl Is Nothing Then
        SetKeywordStateOnRow = False
        Exit Function
    End If

    cleanKeyword = Trim(keyword)
    If Len(cleanKeyword) = 0 Then
        SetKeywordStateOnRow = False
        Exit Function
    End If

    If tbl.DataBodyRange Is Nothing Then
        SetKeywordStateOnRow = False
        Exit Function
    End If

    If rowIndex < 1 Or rowIndex > tbl.DataBodyRange.Rows.Count Then
        SetKeywordStateOnRow = False
        Exit Function
    End If

    existingValue = NormalizeText(tbl.DataBodyRange.Cells(rowIndex, RecordColumn.ColKeywords).Value)

    If Len(existingValue) = 0 Then
        newValue = cleanKeyword & stateSuffix & ";"
    ElseIf InStr(1, existingValue, cleanKeyword & stateSuffix & ";", vbTextCompare) > 0 Then
        newValue = existingValue
    Else
        newValue = existingValue & cleanKeyword & stateSuffix & ";"
    End If

    tbl.DataBodyRange.Cells(rowIndex, RecordColumn.ColKeywords).Value = newValue
    SetKeywordStateOnRow = True
End Function

Public Function RemoveKeywordStateOnRow(ByVal tbl As ListObject, ByVal rowIndex As Long, ByVal keyword As String) As Boolean
    Dim existingValue As String
    Dim cleanKeyword As String
    Dim newValue As String

    If tbl Is Nothing Then
        RemoveKeywordStateOnRow = False
        Exit Function
    End If

    cleanKeyword = Trim(keyword)
    If Len(cleanKeyword) = 0 Then
        RemoveKeywordStateOnRow = False
        Exit Function
    End If

    If tbl.DataBodyRange Is Nothing Then
        RemoveKeywordStateOnRow = False
        Exit Function
    End If

    If rowIndex < 1 Or rowIndex > tbl.DataBodyRange.Rows.Count Then
        RemoveKeywordStateOnRow = False
        Exit Function
    End If

    existingValue = NormalizeText(tbl.DataBodyRange.Cells(rowIndex, RecordColumn.ColKeywords).Value)
    If Len(existingValue) = 0 Then
        RemoveKeywordStateOnRow = False
        Exit Function
    End If

    newValue = Replace(existingValue, cleanKeyword & "-yes;", "", 1, , vbTextCompare)
    newValue = Replace(newValue, cleanKeyword & "-maybe;", "", 1, , vbTextCompare)
    newValue = Replace(newValue, cleanKeyword & "-no;", "", 1, , vbTextCompare)
    newValue = Replace(newValue, cleanKeyword & ";", "", 1, , vbTextCompare)

    tbl.DataBodyRange.Cells(rowIndex, RecordColumn.ColKeywords).Value = newValue
    RemoveKeywordStateOnRow = True
End Function
