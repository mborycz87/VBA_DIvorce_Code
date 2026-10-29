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
