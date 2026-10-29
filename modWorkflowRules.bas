Option Explicit

Public Function EnsureResultArray(ByVal results As Variant) As Variant
    Dim arr As Variant
    Dim rowCount As Long
    Dim colCount As Long
    Dim i As Long
    Dim j As Long

    If IsEmpty(results) Then
        EnsureResultArray = Empty
        Exit Function
    End If

    arr = results
    rowCount = UBound(arr, 1) - LBound(arr, 1) + 1
    colCount = UBound(arr, 2) - LBound(arr, 2) + 1

    If rowCount <= 0 Or colCount <= 0 Then
        EnsureResultArray = Empty
        Exit Function
    End If

    For i = LBound(arr, 1) To UBound(arr, 1)
        For j = LBound(arr, 2) To UBound(arr, 2)
            If IsError(arr(i, j)) Then
                arr(i, j) = ""
            End If
        Next j
    Next i

    EnsureResultArray = arr
End Function

Public Function SortResultArray(ByVal results As Variant) As Variant
    Dim arr As Variant
    Dim rowCount As Long
    Dim colCount As Long
    Dim i As Long
    Dim j As Long
    Dim tempRow As Variant

    If IsEmpty(results) Then
        SortResultArray = Empty
        Exit Function
    End If

    arr = EnsureResultArray(results)
    If IsEmpty(arr) Then
        SortResultArray = Empty
        Exit Function
    End If

    rowCount = UBound(arr, 1) - LBound(arr, 1) + 1
    colCount = UBound(arr, 2) - LBound(arr, 2) + 1
    ReDim tempRow(0 To colCount - 1)

    For i = LBound(arr, 1) To UBound(arr, 1) - 1
        For j = i + 1 To UBound(arr, 1)
            If IsDate(arr(i, 2)) And IsDate(arr(j, 2)) Then
                If CDate(arr(i, 2)) > CDate(arr(j, 2)) Then
                    For colCount = 0 To UBound(arr, 2)
                        tempRow(colCount) = arr(i, colCount)
                        arr(i, colCount) = arr(j, colCount)
                        arr(j, colCount) = tempRow(colCount)
                    Next colCount
                End If
            End If
        Next j
    Next i

    SortResultArray = arr
End Function

Public Function BuildNoResultsMessage(ByVal value As String) As String
    If Len(Trim(value)) = 0 Then
        BuildNoResultsMessage = "This data does not exist"
    Else
        BuildNoResultsMessage = value
    End If
End Function

Public Sub SafeNoResultsMessage(ByVal form As Object, ByVal messageText As String)
    If form Is Nothing Then Exit Sub
    MsgBox BuildNoResultsMessage(messageText), vbInformation
End Sub

Public Function FlattenResultRow(ByVal sourceRow As Variant) As String
    Dim i As Long
    Dim output As String

    If IsArray(sourceRow) Then
        For i = LBound(sourceRow) To UBound(sourceRow)
            If Len(output) > 0 Then output = output & " | "
            output = output & CStr(sourceRow(i))
        Next i
    Else
        output = CStr(sourceRow)
    End If

    FlattenResultRow = output
End Function

Public Function ResultSetHasRows(ByVal results As Variant) As Boolean
    If IsEmpty(results) Then
        ResultSetHasRows = False
        Exit Function
    End If

    If Not IsArray(results) Then
        ResultSetHasRows = False
        Exit Function
    End If

    ResultSetHasRows = (UBound(results, 1) >= LBound(results, 1))
End Function

Public Function NormalizeResultRow(ByVal results As Variant, ByVal rowIndex As Long) As Variant
    Dim arr As Variant
    Dim output As Variant
    Dim i As Long

    If IsEmpty(results) Then
        NormalizeResultRow = Empty
        Exit Function
    End If

    arr = results
    If rowIndex < LBound(arr, 1) Or rowIndex > UBound(arr, 1) Then
        NormalizeResultRow = Empty
        Exit Function
    End If

    ReDim output(LBound(arr, 2) To UBound(arr, 2))
    For i = LBound(arr, 2) To UBound(arr, 2)
        If IsError(arr(rowIndex, i)) Then
            output(i) = ""
        Else
            output(i) = arr(rowIndex, i)
        End If
    Next i

    NormalizeResultRow = output
End Function

Public Function SafeFormatDateValue(ByVal value As Variant) As String
    If IsNull(value) Or IsEmpty(value) Then
        SafeFormatDateValue = ""
    ElseIf IsDate(value) Then
        SafeFormatDateValue = Format(CDate(value), "m/d/yyyy")
    Else
        SafeFormatDateValue = CStr(value)
    End If
End Function

Public Function SafeFormatTimeValue(ByVal value As Variant) As String
    If IsNull(value) Or IsEmpty(value) Then
        SafeFormatTimeValue = ""
    ElseIf IsDate(value) Then
        SafeFormatTimeValue = Format(CDate(value), "hh:mm")
    Else
        SafeFormatTimeValue = CStr(value)
    End If
End Function

Public Function GetResultCount(ByVal results As Variant) As Long
    If IsEmpty(results) Then
        GetResultCount = 0
        Exit Function
    End If

    If Not IsArray(results) Then
        GetResultCount = 0
        Exit Function
    End If

    GetResultCount = UBound(results, 1) - LBound(results, 1) + 1
End Function
