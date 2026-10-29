Option Explicit

Public Function AddKeywordStateToCell(ByVal targetCell As Range, ByVal keyword As String, ByVal stateSuffix As String) As Boolean
    Dim cleanKeyword As String
    Dim existingValue As String
    Dim newValue As String

    If targetCell Is Nothing Then
        AddKeywordStateToCell = False
        Exit Function
    End If

    cleanKeyword = Trim(keyword)
    If Len(cleanKeyword) = 0 Then
        AddKeywordStateToCell = False
        Exit Function
    End If

    existingValue = NormalizeText(targetCell.Value)
    If Len(existingValue) = 0 Then
        newValue = cleanKeyword & stateSuffix & ";"
    ElseIf InStr(1, existingValue, cleanKeyword & stateSuffix & ";", vbTextCompare) > 0 Then
        newValue = existingValue
    Else
        newValue = existingValue & cleanKeyword & stateSuffix & ";"
    End If

    targetCell.Value = newValue
    AddKeywordStateToCell = True
End Function

Public Function RemoveKeywordStateFromCell(ByVal targetCell As Range, ByVal keyword As String) As Boolean
    Dim cleanKeyword As String
    Dim existingValue As String
    Dim newValue As String

    If targetCell Is Nothing Then
        RemoveKeywordStateFromCell = False
        Exit Function
    End If

    cleanKeyword = Trim(keyword)
    If Len(cleanKeyword) = 0 Then
        RemoveKeywordStateFromCell = False
        Exit Function
    End If

    existingValue = NormalizeText(targetCell.Value)
    If Len(existingValue) = 0 Then
        RemoveKeywordStateFromCell = False
        Exit Function
    End If

    newValue = Replace(existingValue, cleanKeyword & "-yes;", "", 1, , vbTextCompare)
    newValue = Replace(newValue, cleanKeyword & "-maybe;", "", 1, , vbTextCompare)
    newValue = Replace(newValue, cleanKeyword & "-no;", "", 1, , vbTextCompare)
    newValue = Replace(newValue, cleanKeyword & ";", "", 1, , vbTextCompare)

    targetCell.Value = newValue
    RemoveKeywordStateFromCell = True
End Function

Public Function AddKeywordStateToRow(ByVal tbl As ListObject, ByVal rowIndex As Long, ByVal keyword As String, ByVal stateSuffix As String) As Boolean
    Dim targetCell As Range

    If tbl Is Nothing Then
        AddKeywordStateToRow = False
        Exit Function
    End If

    If tbl.DataBodyRange Is Nothing Then
        AddKeywordStateToRow = False
        Exit Function
    End If

    If rowIndex < 1 Or rowIndex > tbl.DataBodyRange.Rows.Count Then
        AddKeywordStateToRow = False
        Exit Function
    End If

    Set targetCell = tbl.DataBodyRange.Cells(rowIndex, RecordColumn.ColKeywords)
    AddKeywordStateToRow = AddKeywordStateToCell(targetCell, keyword, stateSuffix)
End Function

Public Function RemoveKeywordStateFromRow(ByVal tbl As ListObject, ByVal rowIndex As Long, ByVal keyword As String) As Boolean
    Dim targetCell As Range

    If tbl Is Nothing Then
        RemoveKeywordStateFromRow = False
        Exit Function
    End If

    If tbl.DataBodyRange Is Nothing Then
        RemoveKeywordStateFromRow = False
        Exit Function
    End If

    If rowIndex < 1 Or rowIndex > tbl.DataBodyRange.Rows.Count Then
        RemoveKeywordStateFromRow = False
        Exit Function
    End If

    Set targetCell = tbl.DataBodyRange.Cells(rowIndex, RecordColumn.ColKeywords)
    RemoveKeywordStateFromRow = RemoveKeywordStateFromCell(targetCell, keyword)
End Function

Public Function GetKeywordStateForCell(ByVal targetCell As Range, ByVal keyword As String) As String
    Dim existingValue As String
    Dim cleanKeyword As String

    If targetCell Is Nothing Then
        GetKeywordStateForCell = ""
        Exit Function
    End If

    cleanKeyword = Trim(keyword)
    If Len(cleanKeyword) = 0 Then
        GetKeywordStateForCell = ""
        Exit Function
    End If

    existingValue = NormalizeText(targetCell.Value)
    If InStr(1, existingValue, cleanKeyword & "-yes;", vbTextCompare) > 0 Then
        GetKeywordStateForCell = "-yes"
    ElseIf InStr(1, existingValue, cleanKeyword & "-maybe;", vbTextCompare) > 0 Then
        GetKeywordStateForCell = "-maybe"
    ElseIf InStr(1, existingValue, cleanKeyword & "-no;", vbTextCompare) > 0 Then
        GetKeywordStateForCell = "-no"
    ElseIf InStr(1, existingValue, cleanKeyword & ";", vbTextCompare) > 0 Then
        GetKeywordStateForCell = ""
    Else
        GetKeywordStateForCell = ""
    End If
End Function
